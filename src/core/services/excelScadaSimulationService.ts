import * as XLSX from "xlsx";
import type { ScadaCommand, ScadaEmsMode, ScadaTariffLabel } from "./scadaSimulationService.ts";
import { classifyEmsModeByHour, classifyTariffByHour } from "./scadaSimulationService.ts";
import { OFF_PEAK_TARIFF_DH_PER_KWH } from "../constants/ocpDefaults.ts";

export const EXCEL_SCADA_SHEET_NAME = "data base 2025(1)";
export const EXCEL_SCADA_EXPECTED_HOURS = 8760;
export const EXCEL_SCADA_MINUTES_PER_HOUR = 60;
export const EXCEL_SCADA_EXPECTED_MINUTES = EXCEL_SCADA_EXPECTED_HOURS * EXCEL_SCADA_MINUTES_PER_HOUR;

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

export interface ExcelMinuteScadaState {
  minute: number;
  timestamp: string;
  timeLabel: string;
  excelRowIndex: number;
  minuteIndexInHour: number;
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
  rampEvent: boolean;
  alarmCount: number;
  energyBalanceError: number;
  energyBalanceOk: boolean;
}

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
  minuteStates: ExcelMinuteScadaState[];
  hourlySummaries: ExcelHourlySummary[];
  kpis: ExcelScadaKpis;
  completed: boolean;
}

export interface ExcelMinutePointer {
  currentMinuteIndex: number;
  hourIndex: number;
  minuteInsideHour: number;
  totalMinutes: number;
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
  const minuteStates: ExcelMinuteScadaState[] = [];
  const hourlySummaries: ExcelHourlySummary[] = [];

  hourlyRecords.forEach((record, hourIndex) => {
    const hour = record.date.getUTCHours();
    const tariff = classifyTariffByHour(hour);
    const emsMode = classifyEmsModeByHour(hour);
    const previousPvMW = hourIndex > 0 ? hourlyRecords[hourIndex - 1].productionMWh : record.productionMWh;
    const rampEvent = Math.abs(record.productionMWh - previousPvMW) > rampThresholdMW;
    const supportMode: "D" | null = rampEvent ? "D" : null;
    const socStartPercent = toSocPercent(bessEnergyMWh);
    const summaryStartIndex = minuteStates.length;

    for (let minuteInHour = 0; minuteInHour < EXCEL_SCADA_MINUTES_PER_HOUR; minuteInHour += 1) {
      const minuteTimestamp = new Date(record.date.getTime() + minuteInHour * 60_000);
      const pvMWh = record.productionMWh / EXCEL_SCADA_MINUTES_PER_HOUR;
      const loadMWh = record.consumptionMWh / EXCEL_SCADA_MINUTES_PER_HOUR;
      let remainingPvMWh = pvMWh;
      let remainingLoadMWh = loadMWh;
      let pvToLoadMWh = Math.min(remainingPvMWh, remainingLoadMWh);
      remainingPvMWh -= pvToLoadMWh;
      remainingLoadMWh -= pvToLoadMWh;
      let bessChargeMWh = 0;
      let bessDischargeMWh = 0;
      let gridToLoadMWh = 0;
      let gridToBessMWh = 0;
      let wheelingMWh = 0;

      if ((emsMode === "A" || supportMode === "D") && tariff.label !== "Pointe" && remainingPvMWh > 0) {
        bessChargeMWh = Math.min(
          remainingPvMWh,
          BESS_POWER_LIMIT_MW / EXCEL_SCADA_MINUTES_PER_HOUR,
          Math.max(0, (BESS_CAPACITY_MWH - bessEnergyMWh) / CHARGE_EFFICIENCY),
        );
        remainingPvMWh -= bessChargeMWh;
        bessEnergyMWh += bessChargeMWh * CHARGE_EFFICIENCY;
        wheelingMWh = Math.max(0, remainingPvMWh);
      }

      if (emsMode === "B" && remainingLoadMWh > 0) {
        bessDischargeMWh = Math.min(
          remainingLoadMWh,
          BESS_POWER_LIMIT_MW / EXCEL_SCADA_MINUTES_PER_HOUR,
          bessEnergyMWh * DISCHARGE_EFFICIENCY,
        );
        remainingLoadMWh -= bessDischargeMWh;
        bessEnergyMWh -= bessDischargeMWh / DISCHARGE_EFFICIENCY;
      }

      if (emsMode === "C" && allowNightGridCharging && tariff.label === "Creuses" && bessEnergyMWh < targetSocMWh) {
        gridToBessMWh = Math.min(
          BESS_POWER_LIMIT_MW / EXCEL_SCADA_MINUTES_PER_HOUR,
          Math.max(0, (targetSocMWh - bessEnergyMWh) / CHARGE_EFFICIENCY),
        );
        bessEnergyMWh += gridToBessMWh * CHARGE_EFFICIENCY;
      }

      gridToLoadMWh = Math.max(0, remainingLoadMWh);
      bessEnergyMWh = clamp(bessEnergyMWh, 0, BESS_CAPACITY_MWH);
      const gridImportMWh = gridToLoadMWh + gridToBessMWh;
      const leftBalance = pvMWh + bessDischargeMWh + gridImportMWh;
      const rightBalance = loadMWh + bessChargeMWh + wheelingMWh;
      const energyBalanceError = round(leftBalance - rightBalance, 10);
      const hasBalanceError = Math.abs(energyBalanceError) > BALANCE_TOLERANCE_MWH;
      const minuteAlarmCount = Number(rampEvent) + Number(hasBalanceError) + Number(toSocPercent(bessEnergyMWh) >= 99.999);
      alarmCount += minuteAlarmCount;
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

      minuteStates.push({
        minute: minuteStates.length,
        timestamp: formatTimestamp(minuteTimestamp),
        timeLabel: formatTimeLabel(minuteTimestamp),
        excelRowIndex: record.rowIndex,
        minuteIndexInHour: minuteInHour,
        dayNumber: Math.floor(hourIndex / 24) + 1,
        pvMW: round(record.productionMWh, 6),
        loadMW: round(record.consumptionMWh, 6),
        bessPowerMW: round((bessChargeMWh + gridToBessMWh - bessDischargeMWh) * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        socPercent: round(toSocPercent(bessEnergyMWh), 6),
        socMWh: round(bessEnergyMWh, 6),
        gridImportMW: round(gridImportMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        wheelingMW: round(wheelingMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        tariff: tariff.label,
        tariffDhPerKWh: tariff.rateDhPerKWh,
        emsMode,
        supportMode,
        command: getExcelCommand({ bessChargeMWh, bessDischargeMWh, gridToBessMWh, wheelingMWh, gridToLoadMWh }),
        netCostDhPerHour: round(costDh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        costDh: round(costDh, 6),
        baselineCostDh: round(baselineCostDh, 6),
        cumulativeBaselineCostDh: round(baselineCostWithoutEmsDh, 6),
        cumulativeEmsCostDh: round(totalCostWithEmsDh, 6),
        cumulativeGainDh: round(cumulativeGainDh, 6),
        pvToLoadMW: round(pvToLoadMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        pvToBessMW: round(bessChargeMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        pvToWheelingMW: round(wheelingMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        bessToLoadMW: round(bessDischargeMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        gridToLoadMW: round(gridToLoadMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        gridToBessMW: round(gridToBessMWh * EXCEL_SCADA_MINUTES_PER_HOUR, 6),
        pvMWh: round(pvMWh, 10),
        loadMWh: round(loadMWh, 10),
        pvToLoadMWh: round(pvToLoadMWh, 10),
        bessChargeMWh: round(bessChargeMWh + gridToBessMWh, 10),
        bessDischargeMWh: round(bessDischargeMWh, 10),
        gridImportMWh: round(gridImportMWh, 10),
        wheelingMWh: round(wheelingMWh, 10),
        rampEvent,
        alarmCount: minuteAlarmCount,
        energyBalanceError,
        energyBalanceOk: !hasBalanceError,
      });
    }

    const hourStates = minuteStates.slice(summaryStartIndex);
    hourlySummaries.push({
      timestamp: record.timestamp,
      pvMWh: round(sum(hourStates, "pvMWh"), 6),
      loadMWh: round(sum(hourStates, "loadMWh"), 6),
      gridImportMWh: round(sum(hourStates, "gridImportMWh"), 6),
      bessChargeMWh: round(sum(hourStates, "bessChargeMWh"), 6),
      bessDischargeMWh: round(sum(hourStates, "bessDischargeMWh"), 6),
      wheelingMWh: round(sum(hourStates, "wheelingMWh"), 6),
      socStartPercent: round(socStartPercent, 6),
      socEndPercent: round(hourStates[hourStates.length - 1].socPercent, 6),
      tariff: tariff.label,
      emsMode,
      hourlyCostDh: round(sum(hourStates, "costDh"), 6),
      baselineCostDh: round(sum(hourStates, "baselineCostDh"), 6),
      hourlyGainDh: round(sum(hourStates, "baselineCostDh") - sum(hourStates, "costDh"), 6),
    });
  });

  return {
    sourceName,
    hourlyRecords,
    minuteStates,
    hourlySummaries,
    completed: minuteStates.length === EXCEL_SCADA_EXPECTED_MINUTES,
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

export function getExcelMinutePointer(hourlyRecords: ExcelHourlyRecord[], currentMinuteIndex: number): ExcelMinutePointer {
  const totalMinutes = hourlyRecords.length * EXCEL_SCADA_MINUTES_PER_HOUR;
  const boundedMinuteIndex = clamp(Math.floor(currentMinuteIndex), 0, Math.max(0, totalMinutes - 1));
  const hourIndex = Math.floor(boundedMinuteIndex / EXCEL_SCADA_MINUTES_PER_HOUR);
  const minuteInsideHour = boundedMinuteIndex % EXCEL_SCADA_MINUTES_PER_HOUR;
  const hourRecord = hourlyRecords[hourIndex];
  if (!hourRecord) {
    throw new Error(`Invalid minute index: ${currentMinuteIndex}`);
  }

  const timestamp = new Date(hourRecord.date.getTime() + minuteInsideHour * 60_000);
  return {
    currentMinuteIndex: boundedMinuteIndex,
    hourIndex,
    minuteInsideHour,
    totalMinutes,
    dayNumber: Math.floor(hourIndex / 24) + 1,
    progressPercent: totalMinutes > 0 ? ((boundedMinuteIndex + 1) / totalMinutes) * 100 : 0,
    timestamp: formatTimestamp(timestamp),
    cursorHourPosition: hourIndex % 24 + minuteInsideHour / EXCEL_SCADA_MINUTES_PER_HOUR,
  };
}

export function getChartCursorHourPosition(currentMinuteIndex: number): number {
  const boundedMinuteIndex = clamp(Math.floor(currentMinuteIndex), 0, Math.max(0, EXCEL_SCADA_EXPECTED_MINUTES - 1));
  const minuteOfDay = boundedMinuteIndex % (24 * EXCEL_SCADA_MINUTES_PER_HOUR);
  return Math.floor(minuteOfDay / EXCEL_SCADA_MINUTES_PER_HOUR) + (minuteOfDay % EXCEL_SCADA_MINUTES_PER_HOUR) / EXCEL_SCADA_MINUTES_PER_HOUR;
}

export function exportExcelMinuteLogCsv(states: ExcelMinuteScadaState[]): string {
  return toCsv([
    [
      "timestamp",
      "excel_row_index",
      "minute_index",
      "pv_mw",
      "load_mw",
      "bess_power_mw",
      "bess_soc_percent",
      "bess_energy_mwh",
      "grid_import_mw",
      "wheeling_mw",
      "tariff",
      "tariff_dh_kwh",
      "ems_mode",
      "active_command",
      "cost_dh",
      "cumulative_gain_dh",
      "alarm_count",
      "energy_balance_error",
    ],
    ...states.map((state) => [
      state.timestamp,
      state.excelRowIndex,
      state.minute,
      state.pvMW,
      state.loadMW,
      state.bessPowerMW,
      state.socPercent,
      state.socMWh,
      state.gridImportMW,
      state.wheelingMW,
      state.tariff,
      state.tariffDhPerKWh,
      state.emsMode,
      state.command,
      state.costDh,
      state.cumulativeGainDh,
      state.alarmCount,
      state.energyBalanceError,
    ]),
  ]);
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
    minuteCount: result.minuteStates.length,
    firstTimestamp: result.minuteStates[0]?.timestamp,
    lastTimestamp: result.minuteStates[result.minuteStates.length - 1]?.timestamp,
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

function sum(rows: ExcelMinuteScadaState[], key: keyof ExcelMinuteScadaState): number {
  return rows.reduce((total, row) => {
    const value = row[key];
    return total + (typeof value === "number" ? value : 0);
  }, 0);
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
