import type { ScenarioDefinition } from "../models/scenario.ts";
import type { ProjectStudyResult } from "../models/study.ts";
import { ocpBenguerirProjectConfig } from "../../data/examples/ocp_benguerir_project.ts";
import { OCP_LOAD_MW } from "../constants/ocpDefaults.ts";
import { clampPercent } from "../utils/units.ts";
import { exportScenarioComparisonToCsv, exportStudySummaryToJson } from "./exportService.ts";
import { validateAndNormalizeOcpProfile, type NormalizedOcpProfileResult } from "./profileNormalizationService.ts";
import { runProjectStudy } from "./projectStudyService.ts";
import { professionalScenarioDefinitions } from "./scenarioService.ts";

export interface RunOcpStudyFromCsvOptions {
  startDate?: string;
  constantLoadMW?: number;
  minSocPercent?: number;
  allowShortProfileForTesting?: boolean;
  selectedScenarios?: ScenarioDefinition[];
}

export interface OcpCsvStudyResult {
  normalizedProfileSummary: NormalizedOcpProfileResult["summary"] & {
    warnings: string[];
    detectedColumns: NormalizedOcpProfileResult["detectedColumns"];
  };
  studyResult: ProjectStudyResult;
  scenarioComparisonCsv: string;
  studySummaryJson: string;
}

export function runOcpStudyFromCsv(csvText: string, options: RunOcpStudyFromCsvOptions = {}): OcpCsvStudyResult {
  const normalized = validateAndNormalizeOcpProfile(csvText, {
    startDate: options.startDate ?? "2025-01-01T00:00:00",
    constantLoadMW: options.constantLoadMW ?? OCP_LOAD_MW,
    allowShortProfileForTesting: options.allowShortProfileForTesting ?? false,
    simulationStepMinutes: ocpBenguerirProjectConfig.project.simulationStepMinutes,
  });
  const selectedScenarios = options.selectedScenarios ?? professionalScenarioDefinitions;
  const minSocPercent = clampPercent(options.minSocPercent ?? ocpBenguerirProjectConfig.battery.minSocPercent);
  const projectConfig = {
    ...ocpBenguerirProjectConfig,
    battery: {
      ...ocpBenguerirProjectConfig.battery,
      minSocPercent,
      initialSocPercent: Math.max(ocpBenguerirProjectConfig.battery.initialSocPercent, minSocPercent),
    },
    emsStrategy: {
      ...ocpBenguerirProjectConfig.emsStrategy,
      socSecurityReservePercent: minSocPercent,
    },
  };
  const studyResult = runProjectStudy({
    projectConfig,
    profile: normalized.profile,
    selectedScenarios,
    simulationOptions: {
      expectedLength: 8760,
      allowShortProfile: options.allowShortProfileForTesting ?? false,
    },
    notes: "OCP Benguerir study generated from imported CSV profile.",
  });

  return {
    normalizedProfileSummary: {
      ...normalized.summary,
      warnings: normalized.validation.warnings.map((warning) => warning.message),
      detectedColumns: normalized.detectedColumns,
    },
    studyResult,
    scenarioComparisonCsv: exportScenarioComparisonToCsv(studyResult.scenarioComparison),
    studySummaryJson: exportStudySummaryToJson(studyResult),
  };
}
