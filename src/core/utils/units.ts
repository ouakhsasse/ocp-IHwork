export const KWH_PER_MWH = 1000;

export function mwhToKwh(valueMWh: number): number {
  return valueMWh * KWH_PER_MWH;
}

export function kwhToMwh(valueKWh: number): number {
  return valueKWh / KWH_PER_MWH;
}

export function percentToRatio(percent: number): number {
  return percent / 100;
}

export function ratioToPercent(ratio: number): number {
  return ratio * 100;
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

export function clampPercent(value: number): number {
  return clamp(Number.isFinite(value) ? value : 0, 0, 100);
}
