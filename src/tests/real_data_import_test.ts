import { runOcpStudyFromCsv } from "../core/services/realDataStudyRunner.ts";
import { validateAndNormalizeOcpProfile } from "../core/services/profileNormalizationService.ts";
import { OCP_LOAD_MW } from "../core/constants/ocpDefaults.ts";

const sampleCsv = [
  "date,energie_pv",
  "2025-01-01T00:00:00,0",
  "2025-01-01T01:00:00,0",
  "2025-01-01T02:00:00,0",
  "2025-01-01T03:00:00,0",
  "2025-01-01T04:00:00,0",
  "2025-01-01T05:00:00,0",
  "2025-01-01T06:00:00,4",
  "2025-01-01T07:00:00,12",
  "2025-01-01T08:00:00,24",
  "2025-01-01T09:00:00,36",
  "2025-01-01T10:00:00,46",
  "2025-01-01T11:00:00,54",
  "2025-01-01T12:00:00,58",
  "2025-01-01T13:00:00,55",
  "2025-01-01T14:00:00,47",
  "2025-01-01T15:00:00,34",
  "2025-01-01T16:00:00,18",
  "2025-01-01T17:00:00,5",
  "2025-01-01T18:00:00,0",
  "2025-01-01T19:00:00,0",
  "2025-01-01T20:00:00,0",
  "2025-01-01T21:00:00,0",
  "2025-01-01T22:00:00,0",
  "2025-01-01T23:00:00,0",
].join("\n");

const normalized = validateAndNormalizeOcpProfile(sampleCsv, {
  constantLoadMW: OCP_LOAD_MW,
  allowShortProfileForTesting: true,
});
const study = runOcpStudyFromCsv(sampleCsv, {
  constantLoadMW: OCP_LOAD_MW,
  allowShortProfileForTesting: true,
});
const parsedJson = JSON.parse(study.studySummaryJson) as { metadata?: unknown };

assert(normalized.profile.length > 0, "Normalized rows must be greater than 0.");
assert(normalized.profile.every((point) => point.loadMWh === OCP_LOAD_MW), "Missing load column must generate 15 MWh load.");
assert(study.studyResult.recommendedScenario !== null, "Recommended scenario must exist.");
assert(parsedJson.metadata !== undefined, "JSON export must be valid.");

console.log("Real data import test passed.");
console.log("Detected columns:", normalized.detectedColumns);
console.log(`Normalized rows: ${normalized.profile.length}`);
console.log(`Total PV production: ${normalized.summary.totalPVProductionMWh.toFixed(2)} MWh`);
console.log(`Total load: ${normalized.summary.totalLoadMWh.toFixed(2)} MWh`);
console.log(`Recommended scenario: ${study.studyResult.recommendedScenario?.scenarioId ?? "none"}`);
console.log("Scenario comparison CSV preview:");
console.log(study.scenarioComparisonCsv.split("\n").slice(0, 3).join("\n"));

function assert(condition: boolean, message: string): void {
  if (!condition) {
    throw new Error(message);
  }
}
