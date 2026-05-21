import type { BatterySystem } from "../../core/models/battery.ts";
import type { EMSStrategy } from "../../core/models/ems.ts";
import type { IndustrialEnergyProject } from "../../core/models/project.ts";
import type { PVSystem } from "../../core/models/pv.ts";
import type { IndustrialEnergyProjectConfig } from "../../core/models/scenario.ts";
import type { SiteEnergyDemand } from "../../core/models/site.ts";
import { ocpBenguerirTariff } from "./ocp_benguerir_tariffs.ts";
import { ocpBenguerirLoadProfileSample, ocpBenguerirPvProfileSample } from "./ocp_benguerir_profiles_sample.ts";
import { DEFAULT_MIN_SOC_PERCENT, OCP_LOAD_MW } from "../../core/constants/ocpDefaults.ts";

export const ocpBenguerirProject: IndustrialEnergyProject = {
  projectId: "ocp-benguerir-gantour",
  projectName: "OCP Benguerir PV-BESS EMS Study",
  clientName: "OCP Green Energy",
  country: "Morocco",
  siteName: "Benguerir / Gantour",
  sector: "Industrial mining and processing",
  description: "Reusable first case study for a 67 MWc PV plant, 25 MW / 125 MWh BESS, and 15 MW continuous industrial load.",
  createdAt: "2026-05-09T00:00:00.000Z",
  updatedAt: "2026-05-09T00:00:00.000Z",
  currency: "DH",
  simulationStepMinutes: 60,
  studyHorizonYears: 25,
};

export const ocpBenguerirSiteDemand: SiteEnergyDemand = {
  loadProfile: ocpBenguerirLoadProfileSample,
  constantLoadMW: OCP_LOAD_MW,
  annualConsumptionMWh: OCP_LOAD_MW * 8760,
  peakLoadMW: OCP_LOAD_MW,
  operatingHoursPerYear: 8760,
  criticalLoadMW: OCP_LOAD_MW,
  securityAutonomyMinutes: 30,
};

export const ocpBenguerirPvSystem: PVSystem = {
  installedCapacityMWp: 67,
  inverterEfficiency: 0.99,
  poiLossesPercent: 0.95,
  degradationPercentPerYear: 0.5,
  productionProfileMWh: ocpBenguerirPvProfileSample,
};

export const ocpBenguerirBattery: BatterySystem = {
  powerMW: 25,
  capacityMWh: 125,
  minSocPercent: DEFAULT_MIN_SOC_PERCENT,
  maxSocPercent: 100,
  initialSocPercent: 50,
  roundTripEfficiency: 0.9,
  availabilityPercent: 98,
  maxCyclesPerYear: 365,
  cycleCostDhPerKWh: 0.5,
  lifetimeYears: 25,
  degradationPercentPerYear: 1,
};

export const ocpBenguerirEmsStrategy: EMSStrategy = {
  strategyName: "OCP heuristic EMS with wheeling",
  allowPvToLoad: true,
  allowPvToBattery: true,
  allowBatteryToLoad: true,
  allowGridToBatteryOffPeak: true,
  allowGridToBatteryFullHours: false,
  allowGridToBatteryPeakHours: false,
  allowWheeling: true,
  allowCurtailment: true,
  enableWeatherForecastFlag: true,
  enableAdaptiveSocFlag: false,
  enableRampRateControlFlag: true,
  enableOptimizationFlag: false,
  socSecurityReservePercent: DEFAULT_MIN_SOC_PERCENT,
  socTargetSummerPercent: 80,
  socTargetWinterPercent: 70,
  rampRateThresholdMWPerMin: 5,
  gridChargeWhenPvForecastWeak: true,
};

export const ocpBenguerirProjectConfig: IndustrialEnergyProjectConfig = {
  project: ocpBenguerirProject,
  site: ocpBenguerirSiteDemand,
  pv: ocpBenguerirPvSystem,
  battery: ocpBenguerirBattery,
  tariff: ocpBenguerirTariff,
  emsStrategy: ocpBenguerirEmsStrategy,
};
