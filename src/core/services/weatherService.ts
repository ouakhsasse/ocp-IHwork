import { shouldAllowGridToBessOffPeak } from "./scadaSimulationService.ts";

export interface WeatherForecastDay {
  date: string;
  shortwaveRadiationSumMjM2: number;
  cloudCoverMeanPercent: number;
  temperatureMaxC: number;
  sunrise: string;
  sunset: string;
}

export interface WeatherForecastResult {
  source: "open-meteo" | "fallback";
  today: WeatherForecastDay;
  tomorrow: WeatherForecastDay;
  allowGridToBessTonight: boolean;
  message: string;
}

interface OpenMeteoDailyResponse {
  daily?: {
    time?: string[];
    shortwave_radiation_sum?: number[];
    cloud_cover_mean?: number[];
    temperature_2m_max?: number[];
    sunrise?: string[];
    sunset?: string[];
  };
}

export async function fetchBenguerirWeatherForecast(apiBaseUrl = "https://api.open-meteo.com/v1/forecast"): Promise<WeatherForecastResult> {
  const params = new URLSearchParams({
    latitude: "32.2362",
    longitude: "-7.9536",
    daily: "shortwave_radiation_sum,cloud_cover_mean,temperature_2m_max,sunrise,sunset",
    timezone: "Africa/Casablanca",
    forecast_days: "2",
  });

  try {
    const response = await fetch(`${apiBaseUrl}?${params.toString()}`);
    if (!response.ok) {
      throw new Error(`Open-Meteo returned HTTP ${response.status}`);
    }

    const payload = await response.json() as OpenMeteoDailyResponse;
    const today = toForecastDay(payload, 0);
    const tomorrow = toForecastDay(payload, 1);

    return {
      source: "open-meteo",
      today,
      tomorrow,
      allowGridToBessTonight: shouldAllowGridToBessOffPeak(tomorrow.shortwaveRadiationSumMjM2, tomorrow.cloudCoverMeanPercent),
      message: "Open-Meteo forecast loaded.",
    };
  } catch (error) {
    const fallback = buildFallbackWeatherForecast();
    return {
      ...fallback,
      message: `Open-Meteo unavailable, using local weather fallback. ${formatError(error)}`,
    };
  }
}

export function buildFallbackWeatherForecast(): WeatherForecastResult {
  const today = {
    date: "2026-06-21",
    shortwaveRadiationSumMjM2: 23,
    cloudCoverMeanPercent: 28,
    temperatureMaxC: 33,
    sunrise: "2026-06-21T06:26",
    sunset: "2026-06-21T20:39",
  };
  const tomorrow = {
    date: "2026-06-22",
    shortwaveRadiationSumMjM2: 18,
    cloudCoverMeanPercent: 42,
    temperatureMaxC: 34,
    sunrise: "2026-06-22T06:26",
    sunset: "2026-06-22T20:39",
  };

  return {
    source: "fallback",
    today,
    tomorrow,
    allowGridToBessTonight: shouldAllowGridToBessOffPeak(tomorrow.shortwaveRadiationSumMjM2, tomorrow.cloudCoverMeanPercent),
    message: "Using local weather fallback.",
  };
}

function toForecastDay(payload: OpenMeteoDailyResponse, index: number): WeatherForecastDay {
  const daily = payload.daily;
  if (!daily?.time?.[index]) {
    throw new Error("Open-Meteo response is missing daily forecast data.");
  }

  return {
    date: daily.time[index],
    shortwaveRadiationSumMjM2: Number(daily.shortwave_radiation_sum?.[index] ?? 0),
    cloudCoverMeanPercent: Number(daily.cloud_cover_mean?.[index] ?? 0),
    temperatureMaxC: Number(daily.temperature_2m_max?.[index] ?? 0),
    sunrise: daily.sunrise?.[index] ?? `${daily.time[index]}T06:30`,
    sunset: daily.sunset?.[index] ?? `${daily.time[index]}T20:30`,
  };
}

function formatError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
