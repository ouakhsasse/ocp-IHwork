import type { EnergyProfilePoint, ProfileValidationResult } from "../models/profile.ts";
import { OCP_LOAD_MW } from "../constants/ocpDefaults.ts";
import { buildHourlyTimestamps } from "../utils/time.ts";
import { validateHourlyProfile } from "./dataLoader.ts";

export interface DetectedProfileColumns {
  timestampColumn?: string;
  pvColumn?: string;
  loadColumn?: string;
  missingRequiredColumns: string[];
}

export interface ProfileNormalizationOptions {
  startDate?: string;
  constantLoadMW?: number;
  allowShortProfileForTesting?: boolean;
  simulationStepMinutes?: number;
}

export interface NormalizedOcpProfileResult {
  detectedColumns: DetectedProfileColumns;
  profile: EnergyProfilePoint[];
  validation: ProfileValidationResult;
  summary: {
    rowCount: number;
    totalPVProductionMWh: number;
    totalLoadMWh: number;
    generatedTimestamps: boolean;
    generatedConstantLoad: boolean;
  };
}

type RawProfileRow = Record<string, string>;

const TIMESTAMP_ALIASES = ["timestamp", "date", "datetime", "time", "heure", "date_time"];
const PV_ALIASES = ["production_mwh", "pv_production_mwh", "pv_mwh", "production", "energie_pv", "energy_pv", "pv", "pvsyst_mwh"];
const LOAD_ALIASES = ["consumption_mwh", "load_mwh", "demand_mwh", "consommation", "consommation_mwh", "charge_mwh", "load"];

export function detectProfileColumns(headers: string[]): DetectedProfileColumns {
  const normalizedHeaders = headers.map(normalizeHeader);
  const timestampColumn = findColumn(normalizedHeaders, TIMESTAMP_ALIASES);
  const pvColumn = findColumn(normalizedHeaders, PV_ALIASES);
  const loadColumn = findColumn(normalizedHeaders, LOAD_ALIASES);
  const missingRequiredColumns = pvColumn ? [] : ["PV production"];

  return {
    timestampColumn,
    pvColumn,
    loadColumn,
    missingRequiredColumns,
  };
}

export function normalizeRawProfileRows(
  rows: RawProfileRow[],
  options: ProfileNormalizationOptions = {},
): NormalizedOcpProfileResult {
  const normalizedRows = rows.map(normalizeRowKeys);
  const headers = Object.keys(normalizedRows[0] ?? {});
  const detectedColumns = detectProfileColumns(headers);
  const constantLoadMW = options.constantLoadMW ?? OCP_LOAD_MW;
  const simulationStepMinutes = options.simulationStepMinutes ?? 60;
  const stepHours = simulationStepMinutes / 60;
  const canGenerateTimestamps = !detectedColumns.timestampColumn && normalizedRows.length === 8760 && options.startDate;

  if (detectedColumns.missingRequiredColumns.length > 0) {
    throw new Error(`Missing required profile columns: ${detectedColumns.missingRequiredColumns.join(", ")}.`);
  }

  if (!detectedColumns.timestampColumn && !canGenerateTimestamps) {
    throw new Error("Missing timestamp column. Provide one, or provide startDate with exactly 8760 rows.");
  }

  const generatedTimestamps = detectedColumns.timestampColumn
    ? undefined
    : buildHourlyTimestamps(options.startDate ?? "", normalizedRows.length);
  const profile = normalizedRows.map((row, index) => ({
    timestamp: detectedColumns.timestampColumn ? row[detectedColumns.timestampColumn] : generatedTimestamps?.[index] ?? "",
    pvProductionMWh: Number(row[detectedColumns.pvColumn ?? ""]),
    loadMWh: detectedColumns.loadColumn ? Number(row[detectedColumns.loadColumn]) : constantLoadMW * stepHours,
  }));
  const validation = validateHourlyProfile(profile, {
    expectedLength: 8760,
    allowShortProfile: options.allowShortProfileForTesting,
    simulationStepMinutes,
  });

  if (!validation.isValid) {
    throw new Error(`Invalid normalized OCP profile: ${validation.errors.join(" ")}`);
  }

  return {
    detectedColumns,
    profile,
    validation,
    summary: {
      rowCount: profile.length,
      totalPVProductionMWh: sum(profile.map((point) => point.pvProductionMWh)),
      totalLoadMWh: sum(profile.map((point) => point.loadMWh)),
      generatedTimestamps: !detectedColumns.timestampColumn,
      generatedConstantLoad: !detectedColumns.loadColumn,
    },
  };
}

export function buildOcpConstantLoadProfile(profile: EnergyProfilePoint[], constantLoadMW = OCP_LOAD_MW): EnergyProfilePoint[] {
  return profile.map((point) => ({
    ...point,
    loadMWh: constantLoadMW,
  }));
}

export function validateAndNormalizeOcpProfile(
  csvText: string,
  options: ProfileNormalizationOptions = {},
): NormalizedOcpProfileResult {
  const rows = parseCsvRows(csvText);
  return normalizeRawProfileRows(rows, options);
}

function parseCsvRows(csvText: string): RawProfileRow[] {
  const lines = csvText.trim().split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (lines.length < 2) {
    throw new Error("CSV profile must include a header and at least one row.");
  }

  const delimiter = lines[0].includes(";") ? ";" : ",";
  const headers = lines[0].split(delimiter).map(normalizeHeader);

  return lines.slice(1).map((line) => {
    const values = line.split(delimiter).map((value) => value.trim().replace(/^"|"$/g, ""));
    return headers.reduce<RawProfileRow>((row, header, index) => {
      row[header] = values[index] ?? "";
      return row;
    }, {});
  });
}

function normalizeRowKeys(row: RawProfileRow): RawProfileRow {
  return Object.entries(row).reduce<RawProfileRow>((normalized, [key, value]) => {
    normalized[normalizeHeader(key)] = value;
    return normalized;
  }, {});
}

function normalizeHeader(header: string): string {
  return header.trim().replace(/^"|"$/g, "").toLowerCase();
}

function findColumn(headers: string[], aliases: string[]): string | undefined {
  return aliases.find((alias) => headers.includes(alias));
}

function sum(values: number[]): number {
  return values.reduce((total, value) => total + value, 0);
}
