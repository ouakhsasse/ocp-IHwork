import * as XLSX from "xlsx";
import type { ScadaCommand, ScadaEmsMode, ScadaTariffLabel } from "./scadaSimulationService.ts";
import { classifyEmsModeByHour, classifyTariffByHour } from "./scadaSimulationService.ts";
import { OFF_PEAK_TARIFF_DH_PER_KWH } from "../constants/ocpDefaults.ts";

export const EXCEL_SCADA_SHEET_NAME = "data base 2025(1)";
export const EXCEL_SCADA_EXPECTED_HOURS = 8760;
export const EXCEL_SCADA_EXPECTED_STEPS = EXCEL_SCADA_EXPECTED_HOURS;

export interface ExcelHourlyRecord {
  rowIndex: number;
  timestamp: string;
  date: Date;
  productionMWh: number;
  consumptionMWh: number;
}

export interface NormalizedExcelScadaRow {
  timestamp: unknown;
  production_mwh: unknown;
  consumption_mwh: unknown;
}

export interface ExcelHourScadaState {
  hourIndex: number;
  minute: number;
  timestamp: string;
  timeLabel: string;
  excelRowIndex: number;
  hourOfDay: number;
  dayNumber: number;
  pvMW: number;
  loadMW: number;
  bessPowerMW: number;
  socPercent: number;
  socMWh: number;
  gridImportMW: number;
  wheelingMW: number;
  tariff: ScadaTariffLabel;
  tariffDhPerKWh: number;
  emsMode: Exclude<ScadaEmsMode, "D">;
  supportMode: "D" | null;
  command: ScadaCommand;
  netCostDhPerHour: number;
  costDh: number;
  baselineCostDh: number;
  cumulativeBaselineCostDh: number;
  cumulativeEmsCostDh: number;
  cumulativeGainDh: number;
  pvToLoadMW: number;
  pvToBessMW: number;
  pvToWheelingMW: number;
  bessToLoadMW: number;
  gridToLoadMW: number;
  gridToBessMW: number;
  pvMWh: number;
  loadMWh: number;
  pvToLoadMWh: number;
  bessChargeMWh: number;
  bessDischargeMWh: number;
  gridImportMWh: number;
  wheelingMWh: number;
  gridToLoadMWh: number;
  gridToBessMWh: number;
  pvToBessMWh: number;
  bessPowerMWh: number;
  bessEnergyStartMWh: number;
  bessEnergyEndMWh: number;
  socStartPercent: number;
  socEndPercent: number;
  rampEvent: boolean;
  alarmCount: number;
  energyBalanceError: number;
  energyBalanceOk: boolean;
}

export type ExcelMinuteScadaState = ExcelHourScadaState;

export interface ExcelHourlySummary {
  timestamp: string;
  pvMWh: number;
  loadMWh: number;
  gridImportMWh: number;
  bessChargeMWh: number;
  bessDischargeMWh: number;
  wheelingMWh: number;
  socStartPercent: number;
  socEndPercent: number;
  tariff: ScadaTariffLabel;
  emsMode: Exclude<ScadaEmsMode, "D">;
  hourlyCostDh: number;
  baselineCostDh: number;
  hourlyGainDh: number;
}

export interface ExcelScadaKpis {
  totalPvProductionMWh: number;
  totalOcpConsumptionMWh: number;
  totalGridImportMWh: number;
  totalBessChargeMWh: number;
  totalBessDischargeMWh: number;
  totalWheelingMWh: number;
  finalSocPercent: number;
  equivalentBessCycles: number;
  totalCostWithEmsDh: number;
  baselineCostWithoutEmsDh: number;
  totalEconomicGainDh: number;
  alarmCount: number;
  energyBalanceErrorCount: number;
}

export interface ExcelScadaSimulationResult {
  sourceName: string;
  hourlyRecords: ExcelHourlyRecord[];
  hourStates: ExcelHourScadaState[];
  minuteStates: ExcelHourScadaState[];
  hourlySummaries: ExcelHourlySummary[];
  kpis: ExcelScadaKpis;
  completed: boolean;
}

export interface ExcelHourPointer {
  currentHourIndex: number;
  hourIndex: number;
  hourOfDay: number;
  totalHours: number;
  dayNumber: number;
  progressPercent: number;
  timestamp: string;
  cursorHourPosition: number;
}

export interface ExcelScadaSimulationOptions {
  allowNightGridCharging?: boolean;
  initialSocPercent?: number;
  targetSocPercent?: number;
  rampThresholdMW?: number;
}

const BESS_CAPACITY_MWH = 125;
const BESS_POWER_LIMIT_MW = 25;
const CHARGE_EFFICIENCY = Math.sqrt(0.9);
const DISCHARGE_EFFICIENCY = Math.sqrt(0.9);
const WHEELING_VALUE_DH_PER_KWH = OFF_PEAK_TARIFF_DH_PER_KWH;
const BALANCE_TOLERANCE_MWH = 1e-8;

export function parseExcelScadaWorkbook(workbookData: ArrayBuffer | Uint8Array, sourceName = "data base 2025(1)mod.xlsx"): ExcelHourlyRecord[] {
  const workbook = XLSX.read(workbookData, { type: "array", cellDates: true });
  console.log("Workbook sheets:", workbook.SheetNames);
  const sheetName = workbook.SheetNames.includes(EXCEL_SCADA_SHEET_NAME)
    ? EXCEL_SCADA_SHEET_NAME
    : workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("File reading failed: workbook does not contain any sheets.");
  }

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`File reading failed: unable to read sheet ${sheetName}.`);
  }

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: null });
  console.log("Parsed rows:", rows.length);
  if (rows.length !== EXCEL_SCADA_EXPECTED_HOURS) {
    throw new Error(`Invalid row count: expected ${EXCEL_SCADA_EXPECTED_HOURS}, got ${rows.length} in ${sourceName}.`);
  }

  const records = rows.map((row, index) => {
    const normalized = normalizeExcelRow(row);
    if (index === 0) console.log("Normalized first row:", normalized);
    const date = parseExcelDate(normalized.timestamp, index + 2);
    const productionMWh = toFiniteNumber(normalized.production_mwh, `production_mwh row ${index + 2}`);
    const consumptionMWh = toFiniteNumber(normalized.consumption_mwh, `consumption_mwh row ${index + 2}`);

    if (productionMWh < 0) throw new Error(`production_mwh must be >= 0 at row ${index + 2}.`);
    if (consumptionMWh <= 0) throw new Error(`consumption_mwh must be > 0 at row ${index + 2}.`);

    return {
      rowIndex: index,
      timestamp: formatTimestamp(date),
      date,
      productionMWh,
      consumptionMWh,
    };
  });

  validateHourlyContinuity(records);
  return records;
}

export function normalizeExcelRow(row: Record<string, unknown>): NormalizedExcelScadaRow {
  const keyMap = new Map(Object.keys(row).map((key) => [normalizeColumnKey(key), key]));
  const timestampKey = findColumn(keyMap, ["la_date", "date", "timestamp", "datetime", "time"]);
  const productionKey = findColumn(keyMap, ["production_mwh", "production", "pv_mwh", "pv_production_mwh", "production_mw", "pv"]);
  const consumptionKey = findColumn(keyMap, ["consumption_mwh", "consumption", "load_mwh", "load", "consommation", "consommation_mwh", "ocp_load_mwh"]);

  if (!timestampKey) throw new Error("Missing column: La date");
  if (!productionKey) throw new Error("Missing column: production_mwh");
  if (!consumptionKey) throw new Error("Missing column: consumption_mwh");

  return {
    timestamp: row[timestampKey],
    production_mwh: row[productionKey],
    consumption_mwh: row[consumptionKey],
  };
}

export function simulateExcelScadaFullYear(
  hourlyRecords: ExcelHourlyRecord[],
  options: ExcelScadaSimulationOptions = {},
  sourceName = "data base 2025(1)mod.xlsx",
): ExcelScadaSimulationResult {
  if (hourlyRecords.length !== EXCEL_SCADA_EXPECTED_HOURS) {
    throw new Error(`Excel SCADA simulation requires ${EXCEL_SCADA_EXPECTED_HOURS} hourly rows.`);
  }

  const allowNightGridCharging = options.allowNightGridCharging ?? false;
  const targetSocMWh = BESS_CAPACITY_MWH * ((options.targetSocPercent ?? 78) / 100);
  const rampThresholdMW = options.rampThresholdMW ?? 5;
  let bessEnergyMWh = clamp(BESS_CAPACITY_MWH * ((options.initialSocPercent ?? 0) / 100), 0, BESS_CAPACITY_MWH);
  let cumulativeGainDh = 0;
  let totalCostWithEmsDh = 0;
  let baselineCostWithoutEmsDh = 0;
  let totalPvProductionMWh = 0;
  let totalOcpConsumptionMWh = 0;
  let totalGridImportMWh = 0;
  let totalBessChargeMWh = 0;
  let totalBessDischargeMWh = 0;
  let totalWheelingMWh = 0;
  let alarmCount = 0;
  let energyBalanceErrorCount = 0;
  const hourStates: ExcelHourScadaState[] = [];
  const hourlySummaries: ExcelHourlySummary[] = [];

  hourlyRecords.forEach((record, hourIndex) => {
    const hour = record.date.getUTCHours();
    const tariff = classifyTariffByHour(hour);
    const emsMode = classifyEmsModeByHour(hour);
    const previousPvMW = hourIndex > 0 ? hourlyRecords[hourIndex - 1].productionMWh : record.productionMWh;
    const rampEvent = Math.abs(record.productionMWh - previousPvMW) > rampThresholdMW;
    const supportMode: "D" | null = rampEvent ? "D" : null;
    const bessEnergyStartMWh = bessEnergyMWh;
    const socStartPercent = toSocPercent(bessEnergyMWh);
    const pvMWh = record.productionMWh;
    const loadMWh = record.consumptionMWh;
    let remainingPvMWh = pvMWh;
    let remainingLoadMWh = loadMWh;
    const pvToLoadMWh = Math.min(remainingPvMWh, remainingLoadMWh);
    remainingPvMWh -= pvToLoadMWh;
    remainingLoadMWh -= pvToLoadMWh;
    let pvToBessMWh = 0;
    let gridToBessMWh = 0;
    let bessDischargeMWh = 0;
    let gridToLoadMWh = 0;
    let wheelingMWh = 0;

    if ((emsMode === "A" || supportMode === "D") && tariff.label !== "Pointe" && remainingPvMWh > 0) {
      pvToBessMWh = Math.min(
        remainingPvMWh,
        BESS_POWER_LIMIT_MW,
        Math.max(0, (BESS_CAPACITY_MWH - bessEnergyMWh) / CHARGE_EFFICIENCY),
      );
      remainingPvMWh -= pvToBessMWh;
      bessEnergyMWh += pvToBessMWh * CHARGE_EFFICIENCY;
    }

    if (emsMode === "B" && remainingLoadMWh > 0) {
      bessDischargeMWh = Math.min(
        remainingLoadMWh,
        BESS_POWER_LIMIT_MW,
        bessEnergyMWh * DISCHARGE_EFFICIENCY,
      );
      remainingLoadMWh -= bessDischargeMWh;
      bessEnergyMWh -= bessDischargeMWh / DISCHARGE_EFFICIENCY;
    }

    if (emsMode === "C" && allowNightGridCharging && tariff.label === "Creuses" && bessEnergyMWh < targetSocMWh) {
      gridToBessMWh = Math.min(
        BESS_POWER_LIMIT_MW,
        Math.max(0, (targetSocMWh - bessEnergyMWh) / CHARGE_EFFICIENCY),
      );
      bessEnergyMWh += gridToBessMWh * CHARGE_EFFICIENCY;
    }

    wheelingMWh = Math.max(0, remainingPvMWh);
    remainingPvMWh = 0;
    gridToLoadMWh = Math.max(0, remainingLoadMWh);
    bessEnergyMWh = clamp(bessEnergyMWh, 0, BESS_CAPACITY_MWH);
    const socEndPercent = toSocPercent(bessEnergyMWh);
    const bessChargeMWh = pvToBessMWh + gridToBessMWh;
    const gridImportMWh = gridToLoadMWh + gridToBessMWh;
    const leftBalance = pvMWh + bessDischargeMWh + gridImportMWh;
    const rightBalance = loadMWh + bessChargeMWh + wheelingMWh;
    const energyBalanceError = round(leftBalance - rightBalance, 10);
    const loadBalanceError = round(loadMWh - (pvToLoadMWh + bessDischargeMWh + gridToLoadMWh), 10);
    const hasBalanceError = Math.abs(energyBalanceError) > BALANCE_TOLERANCE_MWH || Math.abs(loadBalanceError) > BALANCE_TOLERANCE_MWH;
    const hourAlarmCount = Number(rampEvent) + Number(hasBalanceError) + Number(socEndPercent >= 99.999);
    alarmCount += hourAlarmCount;
    if (hasBalanceError) energyBalanceErrorCount += 1;

    const costDh = gridImportMWh * 1000 * tariff.rateDhPerKWh - wheelingMWh * 1000 * WHEELING_VALUE_DH_PER_KWH;
    const baselineCostDh = loadMWh * 1000 * tariff.rateDhPerKWh;
    totalCostWithEmsDh += costDh;
    baselineCostWithoutEmsDh += baselineCostDh;
    cumulativeGainDh += baselineCostDh - costDh;
    totalPvProductionMWh += pvMWh;
    totalOcpConsumptionMWh += loadMWh;
    totalGridImportMWh += gridImportMWh;
    totalBessChargeMWh += bessChargeMWh;
    totalBessDischargeMWh += bessDischargeMWh;
    totalWheelingMWh += wheelingMWh;

    const state: ExcelHourScadaState = {
      hourIndex,
      minute: hourIndex,
      timestamp: record.timestamp,
      timeLabel: formatTimeLabel(record.date),
      excelRowIndex: record.rowIndex,
      hourOfDay: hour,
      dayNumber: Math.floor(hourIndex / 24) + 1,
      pvMW: round(record.productionMWh, 6),
      loadMW: round(record.consumptionMWh, 6),
      bessPowerMW: round(bessChargeMWh - bessDischargeMWh, 6),
      socPercent: round(socEndPercent, 6),
      socMWh: round(bessEnergyMWh, 6),
      gridImportMW: round(gridImportMWh, 6),
      wheelingMW: round(wheelingMWh, 6),
      tariff: tariff.label,
      tariffDhPerKWh: tariff.rateDhPerKWh,
      emsMode,
      supportMode,
      command: getExcelCommand({ bessChargeMWh, bessDischargeMWh, gridToBessMWh, wheelingMWh, gridToLoadMWh }),
      netCostDhPerHour: round(costDh, 6),
      costDh: round(costDh, 6),
      baselineCostDh: round(baselineCostDh, 6),
      cumulativeBaselineCostDh: round(baselineCostWithoutEmsDh, 6),
      cumulativeEmsCostDh: round(totalCostWithEmsDh, 6),
      cumulativeGainDh: round(cumulativeGainDh, 6),
      pvToLoadMW: round(pvToLoadMWh, 6),
      pvToBessMW: round(pvToBessMWh, 6),
      pvToWheelingMW: round(wheelingMWh, 6),
      bessToLoadMW: round(bessDischargeMWh, 6),
      gridToLoadMW: round(gridToLoadMWh, 6),
      gridToBessMW: round(gridToBessMWh, 6),
      pvMWh: round(pvMWh, 10),
      loadMWh: round(loadMWh, 10),
      pvToLoadMWh: round(pvToLoadMWh, 10),
      bessChargeMWh: round(bessChargeMWh, 10),
      bessDischargeMWh: round(bessDischargeMWh, 10),
      gridImportMWh: round(gridImportMWh, 10),
      wheelingMWh: round(wheelingMWh, 10),
      gridToLoadMWh: round(gridToLoadMWh, 10),
      gridToBessMWh: round(gridToBessMWh, 10),
      pvToBessMWh: round(pvToBessMWh, 10),
      bessPowerMWh: round(bessChargeMWh - bessDischargeMWh, 10),
      bessEnergyStartMWh: round(bessEnergyStartMWh, 10),
      bessEnergyEndMWh: round(bessEnergyMWh, 10),
      socStartPercent: round(socStartPercent, 6),
      socEndPercent: round(socEndPercent, 6),
      rampEvent,
      alarmCount: hourAlarmCount,
      energyBalanceError: hasBalanceError ? round(energyBalanceError + loadBalanceError, 10) : 0,
      energyBalanceOk: !hasBalanceError,
    };
    hourStates.push(state);
    hourlySummaries.push({
      timestamp: record.timestamp,
      pvMWh: state.pvMWh,
      loadMWh: state.loadMWh,
      gridImportMWh: state.gridImportMWh,
      bessChargeMWh: state.bessChargeMWh,
      bessDischargeMWh: state.bessDischargeMWh,
      wheelingMWh: state.wheelingMWh,
      socStartPercent: round(socStartPercent, 6),
      socEndPercent: state.socEndPercent,
      tariff: tariff.label,
      emsMode,
      hourlyCostDh: state.costDh,
      baselineCostDh: state.baselineCostDh,
      hourlyGainDh: round(state.baselineCostDh - state.costDh, 6),
    });
  });

  return {
    sourceName,
    hourlyRecords,
    hourStates,
    minuteStates: hourStates,
    hourlySummaries,
    completed: hourStates.length === EXCEL_SCADA_EXPECTED_HOURS,
    kpis: {
      totalPvProductionMWh: round(totalPvProductionMWh, 6),
      totalOcpConsumptionMWh: round(totalOcpConsumptionMWh, 6),
      totalGridImportMWh: round(totalGridImportMWh, 6),
      totalBessChargeMWh: round(totalBessChargeMWh, 6),
      totalBessDischargeMWh: round(totalBessDischargeMWh, 6),
      totalWheelingMWh: round(totalWheelingMWh, 6),
      finalSocPercent: round(toSocPercent(bessEnergyMWh), 6),
      equivalentBessCycles: round(totalBessDischargeMWh / BESS_CAPACITY_MWH, 6),
      totalCostWithEmsDh: round(totalCostWithEmsDh, 6),
      baselineCostWithoutEmsDh: round(baselineCostWithoutEmsDh, 6),
      totalEconomicGainDh: round(baselineCostWithoutEmsDh - totalCostWithEmsDh, 6),
      alarmCount,
      energyBalanceErrorCount,
    },
  };
}

export function getExcelHourPointer(hourlyRecords: ExcelHourlyRecord[], currentHourIndex: number): ExcelHourPointer {
  const totalHours = hourlyRecords.length;
  const boundedHourIndex = clamp(Math.floor(currentHourIndex), 0, Math.max(0, totalHours - 1));
  const hourIndex = boundedHourIndex;
  const hourRecord = hourlyRecords[hourIndex];
  if (!hourRecord) {
    throw new Error(`Invalid hour index: ${currentHourIndex}`);
  }

  return {
    currentHourIndex: boundedHourIndex,
    hourIndex,
    hourOfDay: hourRecord.date.getUTCHours(),
    totalHours,
    dayNumber: Math.floor(hourIndex / 24) + 1,
    progressPercent: totalHours > 0 ? ((boundedHourIndex + 1) / totalHours) * 100 : 0,
    timestamp: hourRecord.timestamp,
    cursorHourPosition: hourRecord.date.getUTCHours(),
  };
}

export function getChartCursorHourPosition(currentHourIndex: number): number {
  const boundedHourIndex = clamp(Math.floor(currentHourIndex), 0, Math.max(0, EXCEL_SCADA_EXPECTED_HOURS - 1));
  return boundedHourIndex % 24;
}

export function exportExcelHourlyLogCsv(states: ExcelHourScadaState[]): string {
  return toCsv([
    [
      "timestamp",
      "excel_row_index",
      "pv_mwh",
      "load_mwh",
      "pv_to_load_mwh",
      "pv_to_bess_mwh",
      "bess_to_load_mwh",
      "grid_to_load_mwh",
      "grid_to_bess_mwh",
      "wheeling_mwh",
      "bess_power_mw",
      "bess_soc_start_percent",
      "bess_soc_end_percent",
      "tariff",
      "tariff_dh_kwh",
      "ems_mode",
      "active_command",
      "cost_dh",
      "baseline_cost_dh",
      "gain_dh",
      "cumulative_gain_dh",
      "energy_balance_ok",
    ],
    ...states.map((state) => [
      state.timestamp,
      state.excelRowIndex,
      state.pvMWh,
      state.loadMWh,
      state.pvToLoadMWh,
      state.pvToBessMWh,
      state.bessDischargeMWh,
      state.gridToLoadMWh,
      state.gridToBessMWh,
      state.wheelingMWh,
      state.bessPowerMW,
      state.socStartPercent,
      state.socEndPercent,
      state.tariff,
      state.tariffDhPerKWh,
      state.emsMode,
      state.command,
      state.costDh,
      state.baselineCostDh,
      state.baselineCostDh - state.costDh,
      state.cumulativeGainDh,
      state.energyBalanceOk ? "OK" : "ERROR",
    ]),
  ]);
}

export const exportExcelMinuteLogCsv = exportExcelHourlyLogCsv;

export function buildExcelHourlyChartSeries(history: ExcelHourScadaState[]): Record<string, number[]> {
  return {
    pvToLoad: history.map((state) => state.pvToLoadMWh),
    bessToLoad: history.map((state) => state.bessDischargeMWh),
    gridToLoad: history.map((state) => state.gridToLoadMWh),
    bessPower: history.map((state) => state.bessPowerMW),
    pvToBess: history.map((state) => state.pvToBessMWh),
    pvToWheeling: history.map((state) => state.wheelingMWh),
    gridToBess: history.map((state) => state.gridToBessMWh),
    gridImport: history.map((state) => state.gridImportMWh),
    socEnd: history.map((state) => state.socEndPercent),
    baselineCost: history.map((state) => state.baselineCostDh),
    emsCost: history.map((state) => state.costDh),
    cumulativeGain: history.map((state) => state.cumulativeGainDh),
  };
}

export function exportExcelHourlySummaryCsv(summaries: ExcelHourlySummary[]): string {
  return toCsv([
    [
      "timestamp",
      "pv_mwh",
      "load_mwh",
      "grid_import_mwh",
      "bess_charge_mwh",
      "bess_discharge_mwh",
      "wheeling_mwh",
      "soc_start_percent",
      "soc_end_percent",
      "tariff",
      "ems_mode",
      "hourly_cost_dh",
      "baseline_cost_dh",
      "hourly_gain_dh",
    ],
    ...summaries.map((summary) => [
      summary.timestamp,
      summary.pvMWh,
      summary.loadMWh,
      summary.gridImportMWh,
      summary.bessChargeMWh,
      summary.bessDischargeMWh,
      summary.wheelingMWh,
      summary.socStartPercent,
      summary.socEndPercent,
      summary.tariff,
      summary.emsMode,
      summary.hourlyCostDh,
      summary.baselineCostDh,
      summary.hourlyGainDh,
    ]),
  ]);
}

export function exportExcelSummaryJson(result: ExcelScadaSimulationResult): string {
  return JSON.stringify({
    mode: "EXCEL FULL-YEAR SCADA TEST",
    sourceName: result.sourceName,
    hourCount: result.hourlyRecords.length,
    stepCount: result.hourStates.length,
    firstTimestamp: result.hourStates[0]?.timestamp,
    lastTimestamp: result.hourStates[result.hourStates.length - 1]?.timestamp,
    completed: result.completed,
    kpis: result.kpis,
    hourlySummaries: result.hourlySummaries,
  });
}

function validateHourlyContinuity(records: ExcelHourlyRecord[]): void {
  for (let index = 1; index < records.length; index += 1) {
    const diffMs = toWallClockUtcMs(records[index].date) - toWallClockUtcMs(records[index - 1].date);
    if (Math.abs(diffMs - 3_600_000) > 1000) {
      throw new Error(`Timestamps must be hourly and continuous. Gap detected between rows ${index + 1} and ${index + 2}.`);
    }
  }
}

function toWallClockUtcMs(date: Date): number {
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds());
}

function parseExcelDate(value: unknown, rowNumber?: number): Date {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes(), value.getSeconds()));
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S)));
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    const isoLocal = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (isoLocal) {
      return new Date(Date.UTC(
        Number(isoLocal[1]),
        Number(isoLocal[2]) - 1,
        Number(isoLocal[3]),
        Number(isoLocal[4]),
        Number(isoLocal[5]),
        Number(isoLocal[6] ?? 0),
      ));
    }
    const native = new Date(trimmed);
    if (Number.isFinite(native.getTime())) {
      return new Date(Date.UTC(native.getFullYear(), native.getMonth(), native.getDate(), native.getHours(), native.getMinutes(), native.getSeconds()));
    }
    const match = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) {
      return new Date(Date.UTC(
        Number(match[3]),
        Number(match[2]) - 1,
        Number(match[1]),
        Number(match[4] ?? 0),
        Number(match[5] ?? 0),
        Number(match[6] ?? 0),
      ));
    }
  }

  throw new Error(`Cannot parse timestamp${rowNumber ? ` at row ${rowNumber}` : ""}: ${String(value)}`);
}

function toFiniteNumber(value: unknown, label: string): number {
  const numericValue = typeof value === "number" ? value : Number(String(value).replace(",", "."));
  if (!Number.isFinite(numericValue)) throw new Error(`Invalid numeric value for ${label}.`);
  return numericValue;
}

function findColumn(keyMap: Map<string, string>, candidates: string[]): string | undefined {
  return candidates
    .map(normalizeColumnKey)
    .map((candidate) => keyMap.get(candidate))
    .find((key): key is string => Boolean(key));
}

function normalizeColumnKey(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function getExcelCommand(flows: {
  bessChargeMWh: number;
  bessDischargeMWh: number;
  gridToBessMWh: number;
  wheelingMWh: number;
  gridToLoadMWh: number;
}): ScadaCommand {
  if (flows.bessDischargeMWh > 1e-9) return "DISCHARGE_BESS";
  if (flows.gridToBessMWh > 1e-9) return "GRID_TO_BESS_OFFPEAK";
  if (flows.bessChargeMWh > 1e-9) return "CHARGE_BESS_PV";
  if (flows.wheelingMWh > 1e-9) return "WHEELING_EXPORT";
  if (flows.gridToLoadMWh > 1e-9) return "GRID_TO_LOAD";
  return "STANDBY";
}

function toCsv(rows: Array<Array<string | number>>): string {
  return rows.map((row) => row.map(formatCsvCell).join(",")).join("\n");
}

function formatCsvCell(value: string | number): string {
  if (typeof value === "number") return Number.isFinite(value) ? String(round(value, 10)) : "";
  const escaped = value.replace(/"/g, '""');
  return escaped.includes(",") || escaped.includes("\n") || escaped.includes('"') ? `"${escaped}"` : escaped;
}

function toSocPercent(energyMWh: number): number {
  return (energyMWh / BESS_CAPACITY_MWH) * 100;
}

function formatTimestamp(date: Date): string {
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}`;
}

function formatTimeLabel(date: Date): string {
  return `${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`;
}

function pad(value: number): string {
  return String(value).padStart(2, "0");
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function round(value: number, decimals: number): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}
