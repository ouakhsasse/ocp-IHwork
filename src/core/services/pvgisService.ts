import type { HourlyPvPoint } from "./scadaSimulationService.ts";
import { buildFallbackHourlyPvProfile } from "./scadaSimulationService.ts";

export interface PvgisRequestOptions {
  latitude?: number;
  longitude?: number;
  peakPowerKw?: number;
  lossPercent?: number;
  angle?: number;
  aspect?: number;
  startYear?: number;
  endYear?: number;
  apiBaseUrl?: string;
}

export interface PvgisProfileResult {
  source: "pvgis" | "fallback";
  hourlyProfile: HourlyPvPoint[];
  message: string;
}

interface PvgisHourlyOutput {
  time: string;
  P?: number;
}

interface PvgisSeriesResponse {
  outputs?: {
    hourly?: PvgisHourlyOutput[];
  };
}

export async function fetchPvgisHourlyPvProfile(options: PvgisRequestOptions = {}): Promise<PvgisProfileResult> {
  if (isBrowserRuntime() && !options.apiBaseUrl) {
    return {
      source: "fallback",
      hourlyProfile: buildFallbackHourlyPvProfile(),
      message: "PVGIS browser proxy is not configured, using local fallback profile.",
    };
  }

  const params = new URLSearchParams({
    lat: String(options.latitude ?? 32.2362),
    lon: String(options.longitude ?? -7.9536),
    peakpower: String(options.peakPowerKw ?? 67000),
    loss: String(options.lossPercent ?? 14),
    angle: String(options.angle ?? 30),
    aspect: String(options.aspect ?? 0),
    startyear: String(options.startYear ?? 2020),
    endyear: String(options.endYear ?? 2020),
    outputformat: "json",
  });
  const baseUrl = options.apiBaseUrl ?? "https://re.jrc.ec.europa.eu/api/v5_3/";
  const url = `${baseUrl.replace(/\/?$/, "/")}seriescalc?${params.toString()}`;

  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`PVGIS returned HTTP ${response.status}`);
    }

    const payload = await response.json() as PvgisSeriesResponse;
    const hourly = payload.outputs?.hourly ?? [];
    const profile = normalizePvgisHourlyPower(hourly);

    if (profile.length !== 24) {
      throw new Error("PVGIS response did not contain enough hourly PV power data.");
    }

    return {
      source: "pvgis",
      hourlyProfile: profile,
      message: "PVGIS hourly profile loaded.",
    };
  } catch (error) {
    return {
      source: "fallback",
      hourlyProfile: buildFallbackHourlyPvProfile(),
      message: `PVGIS unavailable, using local fallback profile. ${formatError(error)}`,
    };
  }
}

function isBrowserRuntime(): boolean {
  return typeof window !== "undefined" && typeof window.document !== "undefined";
}

export function normalizePvgisHourlyPower(hourly: PvgisHourlyOutput[]): HourlyPvPoint[] {
  const buckets = new Map<number, { total: number; count: number }>();

  hourly.forEach((point) => {
    const hour = Number(point.time.slice(-5, -3));
    const powerKw = Number(point.P ?? 0);
    if (!Number.isFinite(hour) || !Number.isFinite(powerKw)) return;
    const bucket = buckets.get(hour) ?? { total: 0, count: 0 };
    bucket.total += Math.max(0, powerKw / 1000);
    bucket.count += 1;
    buckets.set(hour, bucket);
  });

  if (buckets.size === 0) return [];

  return Array.from({ length: 24 }, (_, hour) => {
    const bucket = buckets.get(hour);
    return {
      hour,
      pvMW: bucket ? Math.round((bucket.total / bucket.count) * 1000) / 1000 : 0,
    };
  });
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
