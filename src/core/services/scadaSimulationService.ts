import { FULL_TARIFF_DH_PER_KWH, OCP_LOAD_MW, OFF_PEAK_TARIFF_DH_PER_KWH, PEAK_TARIFF_DH_PER_KWH } from "../constants/ocpDefaults.ts";

export type ScadaTariffLabel = "Creuses" | "Pleines" | "Pointe";
export type ScadaEmsMode = "A" | "B" | "C" | "D";
export type ScadaCommand =
  | "CHARGE_BESS_PV"
  | "DISCHARGE_BESS"
  | "GRID_TO_LOAD"
  | "GRID_TO_BESS_OFFPEAK"
  | "WHEELING_EXPORT"
  | "STANDBY";

export interface HourlyPvPoint {
  hour: number;
  pvMW: number;
}

export interface MinutePvPoint {
  minute: number;
  hour: number;
  pvMW: number;
}

export interface ScadaForecastDecision {
  allowGridToBess: boolean;
  nextDayIrradiationMjM2: number;
  nextDayCloudCoverPercent: number;
}

export interface ScadaMinuteState {
  minute: number;
  timeLabel: string;
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
  cumulativeGainDh: number;
  pvToLoadMW: number;
  pvToBessMW: number;
  pvToWheelingMW: number;
  bessToLoadMW: number;
  gridToLoadMW: number;
  gridToBessMW: number;
  rampEvent: boolean;
}

const BESS_CAPACITY_MWH = 125;
const BESS_POWER_LIMIT_MW = 25;
const MIN_SOC_PERCENT = 10;
const WHEELING_VALUE_DH_PER_KWH = OFF_PEAK_TARIFF_DH_PER_KWH;

export function classifyTariffByHour(hour: number): { label: ScadaTariffLabel; rateDhPerKWh: number } {
  const normalizedHour = normalizeHour(hour);

  if (normalizedHour >= 17 && normalizedHour < 22) {
    return { label: "Pointe", rateDhPerKWh: PEAK_TARIFF_DH_PER_KWH };
  }

  if ((normalizedHour >= 7 && normalizedHour < 17) || normalizedHour === 22) {
    return { label: "Pleines", rateDhPerKWh: FULL_TARIFF_DH_PER_KWH };
  }

  return { label: "Creuses", rateDhPerKWh: OFF_PEAK_TARIFF_DH_PER_KWH };
}

export function classifyEmsModeByHour(hour: number): Exclude<ScadaEmsMode, "D"> {
  const normalizedHour = normalizeHour(hour);

  if (normalizedHour >= 7 && normalizedHour < 17) return "A";
  if (normalizedHour >= 17 && normalizedHour < 22) return "B";
  return "C";
}

export function shouldAllowGridToBessOffPeak(nextDayIrradiationMjM2: number, cloudCoverPercent: number): boolean {
  return nextDayIrradiationMjM2 < 14 || cloudCoverPercent > 60;
}

export function buildFallbackHourlyPvProfile(): HourlyPvPoint[] {
  const shape = [
    0, 0, 0, 0, 0, 0, 2.5, 10, 22, 35, 47, 57,
    63, 61, 52, 39, 23, 8, 0, 0, 0, 0, 0, 0,
  ];

  return shape.map((pvMW, hour) => ({ hour, pvMW }));
}

export function interpolateHourlyProfileToMinutes(hourlyProfile: HourlyPvPoint[]): MinutePvPoint[] {
  if (hourlyProfile.length !== 24) {
    throw new Error("Hourly PV profile must contain exactly 24 points.");
  }

  return Array.from({ length: 1440 }, (_, minute) => {
    const hour = Math.floor(minute / 60);
    const minuteInHour = minute % 60;
    const current = hourlyProfile[hour].pvMW;
    const next = hourlyProfile[(hour + 1) % 24].pvMW;
    const ratio = minuteInHour / 60;
    return {
      minute,
      hour,
      pvMW: round(Math.max(0, current + (next - current) * ratio), 3),
    };
  });
}

export function simulateScadaDay(
  hourlyProfile: HourlyPvPoint[] = buildFallbackHourlyPvProfile(),
  forecastDecision: ScadaForecastDecision = {
    allowGridToBess: false,
    nextDayIrradiationMjM2: 18,
    nextDayCloudCoverPercent: 35,
  },
): ScadaMinuteState[] {
  const minuteProfile = interpolateHourlyProfileToMinutes(hourlyProfile);
  const states: ScadaMinuteState[] = [];
  let socMWh = BESS_CAPACITY_MWH * 0.5;
  let cumulativeGainDh = 0;

  minuteProfile.forEach((point, index) => {
    const tariff = classifyTariffByHour(point.hour);
    const baseMode = classifyEmsModeByHour(point.hour);
    const previousPv = index > 0 ? minuteProfile[index - 1].pvMW : point.pvMW;
    const rampEvent = Math.abs(point.pvMW - previousPv) >= 0.12 && point.pvMW > 0;
    const supportMode: "D" | null = rampEvent ? "D" : null;
    const loadMW = OCP_LOAD_MW;

    let pvToLoadMW = Math.min(point.pvMW, loadMW);
    let remainingPvMW = Math.max(0, point.pvMW - pvToLoadMW);
    let remainingLoadMW = Math.max(0, loadMW - pvToLoadMW);
    let pvToBessMW = 0;
    let pvToWheelingMW = 0;
    let bessToLoadMW = 0;
    let gridToLoadMW = 0;
    let gridToBessMW = 0;

    if ((baseMode === "A" || supportMode === "D") && tariff.label !== "Pointe" && remainingPvMW > 0) {
      const chargeRoomMW = getChargeRoomMW(socMWh);
      pvToBessMW = Math.min(remainingPvMW, BESS_POWER_LIMIT_MW, chargeRoomMW);
      remainingPvMW -= pvToBessMW;
      pvToWheelingMW = Math.max(0, remainingPvMW);
    }

    if (baseMode === "B" && remainingLoadMW > 0) {
      const dischargeAvailableMW = getDischargeAvailableMW(socMWh);
      bessToLoadMW = Math.min(remainingLoadMW, BESS_POWER_LIMIT_MW, dischargeAvailableMW);
      remainingLoadMW -= bessToLoadMW;
    }

    if (baseMode === "C" && tariff.label === "Creuses" && forecastDecision.allowGridToBess) {
      const targetSocMWh = BESS_CAPACITY_MWH * 0.78;
      if (socMWh < targetSocMWh) {
        gridToBessMW = Math.min(12, BESS_POWER_LIMIT_MW, getChargeRoomMW(socMWh), (targetSocMWh - socMWh) * 60);
      }
    }

    gridToLoadMW = Math.max(0, remainingLoadMW);
    const bessPowerMW = round(pvToBessMW + gridToBessMW - bessToLoadMW, 3);
    socMWh = clamp(socMWh + (bessPowerMW / 60), BESS_CAPACITY_MWH * (MIN_SOC_PERCENT / 100), BESS_CAPACITY_MWH);

    const gridImportMW = gridToLoadMW + gridToBessMW;
    const actualCostDh = gridImportMW * 1000 * tariff.rateDhPerKWh / 60;
    const baselineCostDh = loadMW * 1000 * tariff.rateDhPerKWh / 60;
    const wheelingValueDh = pvToWheelingMW * 1000 * WHEELING_VALUE_DH_PER_KWH / 60;
    cumulativeGainDh += baselineCostDh - actualCostDh + wheelingValueDh;

    states.push({
      minute: point.minute,
      timeLabel: formatMinuteLabel(point.minute),
      pvMW: point.pvMW,
      loadMW,
      bessPowerMW,
      socPercent: round((socMWh / BESS_CAPACITY_MWH) * 100, 2),
      socMWh: round(socMWh, 3),
      gridImportMW: round(gridImportMW, 3),
      wheelingMW: round(pvToWheelingMW, 3),
      tariff: tariff.label,
      tariffDhPerKWh: tariff.rateDhPerKWh,
      emsMode: baseMode,
      supportMode,
      command: getActiveCommand({ pvToBessMW, bessToLoadMW, gridToBessMW, pvToWheelingMW, gridToLoadMW }),
      netCostDhPerHour: round(gridImportMW * 1000 * tariff.rateDhPerKWh - pvToWheelingMW * 1000 * WHEELING_VALUE_DH_PER_KWH, 2),
      cumulativeGainDh: round(cumulativeGainDh, 2),
      pvToLoadMW: round(pvToLoadMW, 3),
      pvToBessMW: round(pvToBessMW, 3),
      pvToWheelingMW: round(pvToWheelingMW, 3),
      bessToLoadMW: round(bessToLoadMW, 3),
      gridToLoadMW: round(gridToLoadMW, 3),
      gridToBessMW: round(gridToBessMW, 3),
      rampEvent,
    });
  });

  return states;
}

function getChargeRoomMW(socMWh: number): number {
  return Math.max(0, (BESS_CAPACITY_MWH - socMWh) * 60);
}

function getDischargeAvailableMW(socMWh: number): number {
  const reserveMWh = BESS_CAPACITY_MWH * (MIN_SOC_PERCENT / 100);
  return Math.max(0, (socMWh - reserveMWh) * 60);
}

function getActiveCommand(flows: {
  pvToBessMW: number;
  bessToLoadMW: number;
  gridToBessMW: number;
  pvToWheelingMW: number;
  gridToLoadMW: number;
}): ScadaCommand {
  if (flows.bessToLoadMW > 0.01) return "DISCHARGE_BESS";
  if (flows.gridToBessMW > 0.01) return "GRID_TO_BESS_OFFPEAK";
  if (flows.pvToBessMW > 0.01) return "CHARGE_BESS_PV";
  if (flows.pvToWheelingMW > 0.01) return "WHEELING_EXPORT";
  if (flows.gridToLoadMW > 0.01) return "GRID_TO_LOAD";
  return "STANDBY";
}

function normalizeHour(hour: number): number {
  return ((Math.floor(hour) % 24) + 24) % 24;
}

function formatMinuteLabel(minute: number): string {
  const hour = Math.floor(minute / 60);
  const minutePart = minute % 60;
  return `${String(hour).padStart(2, "0")}:${String(minutePart).padStart(2, "0")}`;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function round(value: number, decimals: number): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}
