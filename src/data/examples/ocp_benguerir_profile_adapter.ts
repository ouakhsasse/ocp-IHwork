import { parseCsvProfile, type CsvProfileOptions, type ParsedProfiles } from "../../core/services/dataLoader.ts";
import { OCP_LOAD_MW } from "../../core/constants/ocpDefaults.ts";

export interface OcpBenguerirColumnMap {
  timestamp: string;
  productionMWh: string;
  consumptionMWh?: string;
}

export const expectedOcpBenguerirCsvColumns = [
  "timestamp",
  "production_mwh",
  "consumption_mwh",
] as const;

export const ocpBenguerirProfilePreparationNotes = [
  "Export the real OCP PV production database from Excel as CSV.",
  "Map the PV production column to production_mwh in MWh per hour.",
  "Map the industrial demand column to consumption_mwh when available.",
  "If demand is missing, load the CSV with constantLoadMW: 15 to generate OCP's continuous 15 MWh hourly load.",
  "Keep exactly 8760 hourly rows for a normal year or 8784 rows for a leap year.",
];

export function adaptOcpBenguerirRowsToCsv(
  rows: Record<string, string | number>[],
  columnMap: OcpBenguerirColumnMap,
  constantLoadMW = OCP_LOAD_MW,
): string {
  const header = expectedOcpBenguerirCsvColumns.join(",");
  const csvRows = rows.map((row) => {
    const timestamp = String(row[columnMap.timestamp] ?? "");
    const production = Number(row[columnMap.productionMWh] ?? 0);
    const consumption = columnMap.consumptionMWh ? Number(row[columnMap.consumptionMWh] ?? constantLoadMW) : constantLoadMW;
    return [timestamp, production, consumption].join(",");
  });

  return [header, ...csvRows].join("\n");
}

export function loadOcpBenguerirAnnualCsv(csvText: string, options: CsvProfileOptions = {}): ParsedProfiles {
  return parseCsvProfile(csvText, {
    expectedLength: 8760,
    constantLoadMW: OCP_LOAD_MW,
    ...options,
  });
}
