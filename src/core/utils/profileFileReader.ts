import * as XLSX from "xlsx";
import { OCP_LOAD_MW } from "../constants/ocpDefaults.ts";

type ProfileRow = Record<string, unknown>;

export interface ProfileFileReaderOptions {
  defaultStartDate?: string;
  defaultConstantLoadMWh?: number;
  generateConstantLoadIfMissing?: boolean;
  annualRowsUseGeneratedTimestamps?: boolean;
}

export interface ProfileFileReadResult {
  csvText: string;
  detectedFormat: "CSV" | "Excel";
  rowCount: number;
  generatedHourlyTimestamps: boolean;
  generatedConstantLoad: boolean;
}

const TIMESTAMP_ALIASES = ["la date", "date", "timestamp", "datetime", "time", "heure", "date_time"];
const PV_ALIASES = [
  "production enmwh",
  "production en mwh",
  "production_mwh",
  "pv_production_mwh",
  "pv_mwh",
  "production",
  "energie_pv",
  "energy_pv",
  "pv",
  "pvsyst_mwh",
];
const LOAD_ALIASES = ["consumption_mwh", "load_mwh", "demand_mwh", "consommation", "consommation_mwh", "charge_mwh", "load"];

export async function readUploadedProfileFile(file: File, options: ProfileFileReaderOptions = {}): Promise<ProfileFileReadResult> {
  const extension = getFileExtension(file.name);

  if (extension === "csv") {
    const csvText = await file.text();
    return {
      csvText,
      detectedFormat: "CSV",
      rowCount: getCsvRowCount(csvText),
      generatedHourlyTimestamps: false,
      generatedConstantLoad: false,
    };
  }

  if (extension === "xlsx" || extension === "xls") {
    return parseExcelProfileFile(file, options);
  }

  throw new Error("Unsupported file type. Please upload a .csv, .xlsx, or .xls file.");
}

export async function parseCsvProfileFile(file: File): Promise<ProfileFileReadResult> {
  return readUploadedProfileFile(file);
}

export async function parseExcelProfileFile(file: File, options: ProfileFileReaderOptions = {}): Promise<ProfileFileReadResult> {
  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: true,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("Excel file does not contain any worksheets.");
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json<ProfileRow>(worksheet, {
    defval: "",
    raw: true,
  });

  return rowsToNormalizedCsv(rows, options);
}

export function rowsToNormalizedCsv(rows: ProfileRow[], options: ProfileFileReaderOptions = {}): ProfileFileReadResult {
  if (rows.length === 0) {
    throw new Error("Excel worksheet does not contain any data rows.");
  }

  const defaultStartDate = options.defaultStartDate ?? "2025-01-01T00:00:00";
  const defaultConstantLoadMWh = options.defaultConstantLoadMWh ?? OCP_LOAD_MW;
  const generateConstantLoadIfMissing = options.generateConstantLoadIfMissing ?? true;
  const annualRowsUseGeneratedTimestamps = options.annualRowsUseGeneratedTimestamps ?? true;
  const headers = Object.keys(rows[0]).map(normalizeExcelHeader);
  const timestampColumn = findColumn(headers, TIMESTAMP_ALIASES);
  const pvColumn = findColumn(headers, PV_ALIASES);
  const loadColumn = findColumn(headers, LOAD_ALIASES);
  const normalizedRows = rows.map(normalizeRowKeys);
  const generatedHourlyTimestamps = annualRowsUseGeneratedTimestamps && (normalizedRows.length === 8760 || normalizedRows.length === 8784);

  if (!pvColumn) {
    throw new Error("Could not detect PV production column. Expected one of: Production enMWH, production_mwh, pv_mwh, energie_pv, energy_pv, pv, pvsyst_mwh.");
  }

  if (!timestampColumn && !generatedHourlyTimestamps) {
    throw new Error("Could not detect timestamp column. Expected one of: La date, date, timestamp, datetime, time, heure, date_time.");
  }

  if (!loadColumn && !generateConstantLoadIfMissing) {
    throw new Error("Could not detect consumption/load column.");
  }

  const normalizedOutputRows = normalizedRows.map((row, index) => {
    const timestamp = generatedHourlyTimestamps
      ? generateHourlyTimestamp(index, defaultStartDate)
      : excelValueToIsoTimestamp(row[timestampColumn ?? ""], index);
    const pvProductionMWh = parseRequiredNumber(row[pvColumn], "PV production", index);
    const loadMWh = loadColumn ? parseRequiredNumber(row[loadColumn], "load", index) : defaultConstantLoadMWh;

    return {
      timestamp,
      pvProductionMWh,
      loadMWh,
    };
  });

  if (generatedHourlyTimestamps) {
    validateGeneratedAnnualTimestamps(normalizedOutputRows.map((row) => row.timestamp), defaultStartDate);
  }

  const outputRows = normalizedOutputRows.map((row) => [
    row.timestamp,
    formatCsvNumber(row.pvProductionMWh),
    formatCsvNumber(row.loadMWh),
  ].join(","));

  return {
    csvText: ["timestamp,production_mwh,consumption_mwh", ...outputRows].join("\n"),
    detectedFormat: "Excel",
    rowCount: normalizedOutputRows.length,
    generatedHourlyTimestamps,
    generatedConstantLoad: !loadColumn,
  };
}

export function normalizeExcelHeader(header: string): string {
  return header
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/\u00a0/g, " ");
}

export function findColumn(headers: string[], aliases: string[]): string | undefined {
  return aliases.find((alias) => headers.includes(alias));
}

export function generateHourlyTimestamp(rowIndex: number, startDate = "2025-01-01T00:00:00"): string {
  const [datePart, timePart = "00:00:00"] = startDate.split("T");
  const [year, month, day] = datePart.split("-").map(Number);
  const [hour = 0, minute = 0, second = 0] = timePart.split(":").map(Number);
  const utcMs = Date.UTC(year, month - 1, day, hour, minute, second) + rowIndex * 60 * 60 * 1000;
  const date = new Date(utcMs);
  const yyyy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(date.getUTCDate()).padStart(2, "0");
  const hh = String(date.getUTCHours()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}T${hh}:00:00`;
}

function getFileExtension(fileName: string): string {
  const lastDotIndex = fileName.lastIndexOf(".");
  return lastDotIndex === -1 ? "" : fileName.slice(lastDotIndex + 1).toLowerCase();
}

function getCsvRowCount(csvText: string): number {
  return Math.max(0, csvText.trim().split(/\r?\n/).length - 1);
}

function normalizeRowKeys(row: ProfileRow): ProfileRow {
  return Object.entries(row).reduce<ProfileRow>((normalized, [key, value]) => {
    normalized[normalizeExcelHeader(key)] = value;
    return normalized;
  }, {});
}

function excelValueToIsoTimestamp(value: unknown, rowIndex: number): string {
  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialDateToIso(value);
  }

  if (value instanceof Date) {
    return dateToCleanIso(roundDateToNearestHour(value));
  }

  if (typeof value === "string" && value.trim().length > 0) {
    const parsedDate = parseExcelDateString(value);
    if (parsedDate) {
      return dateToCleanIso(roundDateToNearestHour(parsedDate));
    }
  }

  throw new Error(`Invalid or missing timestamp at Excel row ${rowIndex + 2}.`);
}

function excelSerialDateToIso(serial: number): string {
  const parsed = XLSX.SSF.parse_date_code(serial);
  if (!parsed) {
    throw new Error(`Invalid Excel serial date: ${serial}.`);
  }

  const date = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.round(parsed.S)));
  return dateToCleanIso(roundDateToNearestHour(date));
}

function parseExcelDateString(value: string): Date | null {
  const trimmed = value.trim();
  const directDate = new Date(trimmed);
  if (Number.isFinite(directDate.getTime())) {
    return directDate;
  }

  const match = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s+(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  const hour = Number(match[4] ?? 0);
  const minute = Number(match[5] ?? 0);
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, 0));
  return Number.isFinite(date.getTime()) ? date : null;
}

function validateGeneratedAnnualTimestamps(timestamps: string[], startDate: string): void {
  const seen = new Set<string>();
  let previousUtcMs = -Infinity;

  timestamps.forEach((timestamp, index) => {
    if (seen.has(timestamp)) {
      throw new Error(`Generated timestamp validation failed: row ${index + 1} duplicates ${timestamp}.`);
    }
    seen.add(timestamp);

    const currentUtcMs = cleanIsoTimestampToUtcMs(timestamp);
    if (currentUtcMs <= previousUtcMs) {
      throw new Error(`Generated timestamp validation failed: row ${index + 1} is not strictly ordered.`);
    }
    previousUtcMs = currentUtcMs;
  });

  if (timestamps[0] !== startDate) {
    throw new Error(`Generated timestamp validation failed: first timestamp is ${timestamps[0]}.`);
  }
}

function cleanIsoTimestampToUtcMs(timestamp: string): number {
  const match = timestamp.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
  if (!match) {
    throw new Error(`Generated timestamp validation failed: invalid format ${timestamp}.`);
  }

  return Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6]),
  );
}

function roundDateToNearestHour(date: Date): Date {
  const hourMs = 60 * 60 * 1000;
  return new Date(Math.round(date.getTime() / hourMs) * hourMs);
}

function dateToCleanIso(date: Date): string {
  return date.toISOString().slice(0, 19);
}

function parseRequiredNumber(value: unknown, label: string, rowIndex: number): number {
  const parsed = typeof value === "number" ? value : Number(String(value).trim().replace(",", "."));
  if (!Number.isFinite(parsed)) {
    throw new Error(`Invalid numeric value for ${label} at Excel row ${rowIndex + 2}.`);
  }
  if (parsed < 0) {
    throw new Error(`Negative ${label} value at Excel row ${rowIndex + 2}.`);
  }
  return parsed;
}

function formatCsvNumber(value: number): string {
  return String(Math.round(value * 1000000) / 1000000);
}
