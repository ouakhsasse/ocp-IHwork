import type { BatterySystem } from "../models/battery.ts";
import type { GridTariff } from "../models/gridTariff.ts";
import type { IndustrialEnergyProjectConfig } from "../models/scenario.ts";
import type { CombinedProfilePoint, EnergyProfilePoint } from "./time.ts";

export interface ValidationResult {
  isValid: boolean;
  errors: string[];
}

export function createValidationResult(errors: string[]): ValidationResult {
  return {
    isValid: errors.length === 0,
    errors,
  };
}

export function validateBattery(battery: BatterySystem): ValidationResult {
  const errors: string[] = [];
  if (battery.capacityMWh <= 0) errors.push("Battery capacity must be greater than 0 MWh.");
  if (battery.powerMW <= 0) errors.push("Battery power must be greater than 0 MW.");
  if (battery.minSocPercent > battery.maxSocPercent) errors.push("Battery min SoC must be lower than or equal to max SoC.");
  if (battery.initialSocPercent < battery.minSocPercent || battery.initialSocPercent > battery.maxSocPercent) {
    errors.push("Battery initial SoC must be between min and max SoC.");
  }
  if (battery.roundTripEfficiency <= 0 || battery.roundTripEfficiency > 1) {
    errors.push("Battery round-trip efficiency must be in the range (0, 1].");
  }
  return createValidationResult(errors);
}

export function validateTariff(tariff: GridTariff): ValidationResult {
  const values = [
    tariff.peakTariffDhPerKWh,
    tariff.fullTariffDhPerKWh,
    tariff.offPeakTariffDhPerKWh,
    tariff.feedInTariffDhPerKWh ?? 0,
    tariff.wheelingValueDhPerKWh ?? 0,
  ];
  const errors = values.some((value) => value < 0) ? ["Tariff values must be greater than or equal to 0."] : [];
  return createValidationResult(errors);
}

export function validateProfileLengths(pvProfile: EnergyProfilePoint[], loadProfile: EnergyProfilePoint[]): ValidationResult {
  return createValidationResult(pvProfile.length === loadProfile.length ? [] : ["PV and load profile lengths must match."]);
}

export function validateTimestampsOrdered(points: Array<Pick<EnergyProfilePoint, "timestamp">>): ValidationResult {
  const errors: string[] = [];
  for (let index = 1; index < points.length; index += 1) {
    const previous = new Date(points[index - 1].timestamp).getTime();
    const current = new Date(points[index].timestamp).getTime();
    if (!Number.isFinite(current) || !Number.isFinite(previous)) {
      errors.push("Profile timestamps must be valid dates.");
      break;
    }
    if (current <= previous) {
      errors.push("Profile timestamps must be strictly ordered.");
      break;
    }
  }
  return createValidationResult(errors);
}

export function validateSimulationStepMinutes(simulationStepMinutes: number): ValidationResult {
  const validSteps = [5, 10, 15, 30, 60];
  return createValidationResult(
    validSteps.includes(simulationStepMinutes) ? [] : ["Simulation step must be one of 5, 10, 15, 30, or 60 minutes."],
  );
}

export function validateCombinedProfile(points: CombinedProfilePoint[], expectedLength?: number): ValidationResult {
  const errors: string[] = [];
  if (points.length === 0) errors.push("Profile must contain at least one row.");
  if (expectedLength !== undefined && points.length !== expectedLength) {
    errors.push(`Profile must contain exactly ${expectedLength} rows.`);
  }

  points.forEach((point, index) => {
    if (!point.timestamp) errors.push(`Row ${index + 1} is missing a timestamp.`);
    if (!Number.isFinite(point.pvProductionMWh)) errors.push(`Row ${index + 1} has non-numeric PV production.`);
    if (!Number.isFinite(point.loadMWh)) errors.push(`Row ${index + 1} has non-numeric load.`);
    if (point.pvProductionMWh < 0) errors.push(`Row ${index + 1} has negative PV production.`);
    if (point.loadMWh < 0) errors.push(`Row ${index + 1} has negative load.`);
  });

  const timestampValidation = validateTimestampsOrdered(points);
  return createValidationResult([...errors, ...timestampValidation.errors]);
}

export function validateProjectConfig(config: IndustrialEnergyProjectConfig): ValidationResult {
  const errors = [
    ...validateBattery(config.battery).errors,
    ...validateTariff(config.tariff).errors,
    ...validateSimulationStepMinutes(config.project.simulationStepMinutes).errors,
  ];
  return createValidationResult(errors);
}
