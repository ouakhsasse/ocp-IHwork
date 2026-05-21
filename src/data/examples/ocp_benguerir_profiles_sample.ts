import type { EnergyProfilePoint } from "../../core/utils/time.ts";
import { buildHourlyTimestamps } from "../../core/utils/time.ts";
import { OCP_LOAD_MW } from "../../core/constants/ocpDefaults.ts";

const timestamps = buildHourlyTimestamps("2026-01-01T00:00:00.000Z", 48);

const twoDayPvShapeMWh = [
  0, 0, 0, 0, 0, 0, 4, 12, 24, 36, 46, 54,
  58, 55, 47, 34, 18, 5, 0, 0, 0, 0, 0, 0,
  0, 0, 0, 0, 0, 0, 3, 10, 22, 34, 44, 52,
  56, 53, 45, 32, 16, 4, 0, 0, 0, 0, 0, 0,
];

export const ocpBenguerirPvProfileSample: EnergyProfilePoint[] = timestamps.map((timestamp, index) => ({
  timestamp,
  valueMWh: twoDayPvShapeMWh[index],
}));

export const ocpBenguerirLoadProfileSample: EnergyProfilePoint[] = timestamps.map((timestamp) => ({
  timestamp,
  valueMWh: OCP_LOAD_MW,
}));
