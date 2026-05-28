import { fetchPvgisHourlyPvProfile } from "../core/services/pvgisService.ts";
import { fetchBenguerirWeatherForecast } from "../core/services/weatherService.ts";
import * as XLSX from "xlsx";
import {
  buildFallbackHourlyPvProfile,
  classifyEmsModeByHour,
  classifyTariffByHour,
  interpolateHourlyProfileToMinutes,
  simulateScadaDay,
} from "../core/services/scadaSimulationService.ts";
import {
  EXCEL_SCADA_EXPECTED_HOURS,
  EXCEL_SCADA_EXPECTED_MINUTES,
  EXCEL_SCADA_SHEET_NAME,
  exportExcelHourlySummaryCsv,
  exportExcelMinuteLogCsv,
  exportExcelSummaryJson,
  getChartCursorHourPosition,
  getExcelMinutePointer,
  normalizeExcelRow,
  parseExcelScadaWorkbook,
  simulateExcelScadaFullYear,
} from "../core/services/excelScadaSimulationService.ts";

assert(classifyTariffByHour(0).label === "Creuses", "00:00 should be Heures creuses.");
assert(classifyTariffByHour(6).label === "Creuses", "06:00 should be Heures creuses.");
assert(classifyTariffByHour(7).label === "Pleines", "07:00 should be Heures pleines.");
assert(classifyTariffByHour(16).label === "Pleines", "16:00 should be Heures pleines.");
assert(classifyTariffByHour(17).label === "Pointe", "17:00 should be Heures de pointe.");
assert(classifyTariffByHour(21).label === "Pointe", "21:00 should be Heures de pointe.");
assert(classifyTariffByHour(22).label === "Pleines", "22:00 should be Heures pleines.");
assert(classifyTariffByHour(23).label === "Creuses", "23:00 should be Heures creuses.");

assert(classifyEmsModeByHour(8) === "A", "08:00 should use EMS Mode A.");
assert(classifyEmsModeByHour(18) === "B", "18:00 should use EMS Mode B.");
assert(classifyEmsModeByHour(1) === "C", "01:00 should use EMS Mode C.");
assert(classifyEmsModeByHour(22) === "C", "22:00 should use EMS Mode C.");

const minuteProfile = interpolateHourlyProfileToMinutes(buildFallbackHourlyPvProfile());
assert(minuteProfile.length === 1440, "Minute interpolation should create 1440 points.");
assert(minuteProfile[0].minute === 0 && minuteProfile[1439].minute === 1439, "Minute profile should cover the full day.");

const originalFetch = globalThis.fetch;
globalThis.fetch = async () => new Response("PVGIS unavailable", { status: 503 });
const pvgisResult = await fetchPvgisHourlyPvProfile();
assert(pvgisResult.source === "fallback", "PVGIS failure should trigger fallback profile.");
assert(pvgisResult.hourlyProfile.length === 24, "Fallback PVGIS profile should contain 24 hourly values.");

globalThis.fetch = async () => new Response("Open-Meteo unavailable", { status: 503 });
const weatherResult = await fetchBenguerirWeatherForecast();
globalThis.fetch = originalFetch;
assert(weatherResult.source === "fallback", "Open-Meteo failure should trigger fallback weather data.");

const states = simulateScadaDay(buildFallbackHourlyPvProfile(), {
  allowGridToBess: true,
  nextDayIrradiationMjM2: 10,
  nextDayCloudCoverPercent: 70,
});

assert(states.length === 1440, "SCADA day should contain 1440 minute states.");
assert(states.every((state) => state.socPercent >= 0 && state.socPercent <= 100), "SoC should remain between 0 and 100%.");
assert(states.every((state) => state.loadMW === 15), "OCP load should remain constant at 15 MW.");
assert(
  states.every((state) => state.tariff !== "Pointe" || state.bessPowerMW <= 1e-9),
  "BESS should never charge during peak tariff.",
);
assert(states.some((state) => state.supportMode === "D" && state.rampEvent), "Mode D support should appear during PV ramp events.");
assert(states.filter((state) => state.minute < 300 || state.minute >= 1080).every((state) => state.pvMW === 0), "Fallback PV should be zero during deep night.");
assert(states.every((state) => Math.abs((state.pvToLoadMW + state.bessToLoadMW + state.gridToLoadMW) - state.loadMW) < 1e-6), "OCP load supply mix should sum to load.");
assert(states.every((state) => Math.abs((state.pvToLoadMW + state.pvToBessMW + state.pvToWheelingMW) - state.pvMW) < 1e-6), "PV allocation should sum to PV production.");
assert(states.every((state) => Math.abs((state.gridToLoadMW + state.gridToBessMW) - state.gridImportMW) < 1e-6), "Grid import breakdown should sum correctly.");
assert(states.some((state) => state.bessPowerMW > 0), "BESS charging should appear as positive chart data.");
assert(states.some((state) => state.bessPowerMW < 0), "BESS discharging should appear as negative chart data.");

const highPeakPvProfile = buildFallbackHourlyPvProfile().map((point) => (
  point.hour >= 17 && point.hour < 22 ? { ...point, pvMW: 45 } : point
));
const peakStressStates = simulateScadaDay(highPeakPvProfile, {
  allowGridToBess: true,
  nextDayIrradiationMjM2: 10,
  nextDayCloudCoverPercent: 70,
});
assert(
  peakStressStates.every((state) => state.tariff !== "Pointe" || state.pvToBessMW + state.gridToBessMW <= 1e-9),
  "BESS should not charge during peak tariff even under high PV ramp conditions.",
);

const excelRecords = parseExcelScadaWorkbook(buildSyntheticExcelWorkbook());
assert(excelRecords.length === EXCEL_SCADA_EXPECTED_HOURS, "Excel data should produce 8760 hourly records.");
assert(excelRecords[0].productionMWh === 0, "First Excel row should have PV = 0.");
assert(excelRecords[0].consumptionMWh === 15, "First Excel row should have load = 15.");
const alternateSheetRecords = parseExcelScadaWorkbook(buildSyntheticExcelWorkbook("alternate sheet name", true));
assert(alternateSheetRecords.length === EXCEL_SCADA_EXPECTED_HOURS, "Excel parser should fall back to the first sheet when the exact sheet is absent.");
const normalizedVariantRow = normalizeExcelRow({
  Date: "2025-01-01T00:00:00",
  "Production MWh": "0",
  "Consumption MWh": "15",
});
assert(normalizedVariantRow.production_mwh === "0" && normalizedVariantRow.consumption_mwh === "15", "Excel row normalization should support common column variants.");

const excelResult = simulateExcelScadaFullYear(excelRecords, { allowNightGridCharging: false, initialSocPercent: 0 });
assert(excelResult.minuteStates.length === EXCEL_SCADA_EXPECTED_MINUTES, "Hourly Excel data should expand to 525600 minute records.");
assert(excelResult.minuteStates[0].socPercent === 0, "Initial SoC should be 0%.");
assert(excelResult.minuteStates[0].socMWh === 0, "First timestamp should start with BESS discharged.");
assert(excelResult.minuteStates[1].timestamp === "2025-01-01T00:01:00", "Next minute should advance timestamp by exactly 1 minute.");
assert(excelResult.minuteStates[61].minuteIndexInHour === 1, "At 01:01, minuteInsideHour should be 1.");
assert(excelResult.minuteStates[119].timestamp === "2025-01-01T01:59:00", "01:59 should map to minute 119.");
assert(excelResult.minuteStates[120].timestamp === "2025-01-01T02:00:00", "Next minute after 01:59 should move to 02:00.");
assert(getExcelMinutePointer(excelRecords, 61).hourIndex === 1, "currentMinuteIndex should map correctly to Excel hour row.");
assert(Math.abs(excelResult.minuteStates[0].loadMWh - 0.25) < 1e-9, "15 MW load should equal 0.25 MWh per minute.");
assert(excelResult.minuteStates.every((state) => state.socPercent >= 0 && state.socPercent <= 100), "Excel SoC should remain between 0% and 100%.");
assert(excelResult.minuteStates.every((state) => state.tariff !== "Pointe" || state.bessChargeMWh <= 1e-9), "Excel BESS should never charge during Pointe.");
assert(excelResult.minuteStates.every((state) => state.timestamp.slice(11, 13) !== "22" || state.gridToBessMW <= 1e-9), "Excel BESS should not charge from grid at 22h Pleines.");
assert(excelResult.minuteStates.some((state) => state.emsMode === "A" && state.pvToBessMW > 0), "Excel BESS should charge from PV surplus during Mode A.");
assert(excelResult.minuteStates.some((state) => state.emsMode === "B" && state.bessToLoadMW > 0), "Excel BESS should discharge during Mode B when SoC is available.");
assert(excelResult.minuteStates.every((state) => Math.abs(state.energyBalanceError) <= 1e-8), "Excel energy balance should be valid for all minutes.");
assert(excelResult.minuteStates.every((state) => state.energyBalanceOk), "Excel chart/log balance flag should be OK for all synthetic minutes.");
assert(excelResult.minuteStates.every((state) => Math.abs((state.pvToLoadMW + state.bessToLoadMW + state.gridToLoadMW) - state.loadMW) < 1e-6), "Excel OCP load supply mix should sum to load.");
assert(excelResult.minuteStates.every((state) => Math.abs((state.pvToLoadMW + state.pvToBessMW + state.pvToWheelingMW) - state.pvMW) < 1e-6), "Excel PV allocation should sum to PV production.");
assert(excelResult.minuteStates.every((state) => Math.abs((state.gridToLoadMW + state.gridToBessMW) - state.gridImportMW) < 1e-6), "Excel grid import breakdown should sum correctly.");
assert(excelResult.minuteStates.some((state) => state.bessPowerMW > 0), "Excel chart data should include positive BESS charging.");
assert(excelResult.minuteStates.some((state) => state.bessPowerMW < 0), "Excel chart data should include negative BESS discharging.");
assert(excelResult.minuteStates[1].cumulativeGainDh !== excelResult.minuteStates[0].cumulativeGainDh || excelResult.minuteStates[1].cumulativeEmsCostDh > 0, "Economic cumulative data should update after each minute.");
assert(getChartCursorHourPosition(60) === 1, "Chart cursor should map 01:00 to x = 1.00.");
assert(Math.abs(getChartCursorHourPosition(90) - 1.5) < 1e-9, "Chart cursor should map 01:30 to x = 1.50.");
assert(excelResult.minuteStates.every((state) => [
  state.pvToLoadMW,
  state.bessToLoadMW,
  state.gridToLoadMW,
  state.pvToBessMW,
  state.pvToWheelingMW,
  state.gridToBessMW,
  state.gridImportMW,
  state.cumulativeBaselineCostDh,
  state.cumulativeEmsCostDh,
  state.cumulativeGainDh,
].every(Number.isFinite)), "Chart datasets should not contain NaN or undefined values.");
assert(excelResult.completed, "Full Excel simulation should reach completion.");
assert(exportExcelMinuteLogCsv(excelResult.minuteStates).split("\n").length === EXCEL_SCADA_EXPECTED_MINUTES + 1, "Minute CSV should export every minute plus header.");
assert(exportExcelHourlySummaryCsv(excelResult.hourlySummaries).split("\n").length === EXCEL_SCADA_EXPECTED_HOURS + 1, "Hourly CSV should export every hour plus header.");
const excelJson = JSON.parse(exportExcelSummaryJson(excelResult));
assert(excelJson.minuteCount === EXCEL_SCADA_EXPECTED_MINUTES, "Full-year JSON should include minute count.");
assert(excelJson.hourlySummaries.length === EXCEL_SCADA_EXPECTED_HOURS, "Full-year JSON should include hourly data.");

console.log("SCADA supervision tests passed.");

function assert(condition: boolean, message: string): void {
  if (!condition) {
    throw new Error(message);
  }
}

function buildSyntheticExcelWorkbook(sheetName = EXCEL_SCADA_SHEET_NAME, useColumnVariants = false): Uint8Array {
  const rows = Array.from({ length: EXCEL_SCADA_EXPECTED_HOURS }, (_, index) => {
    const date = new Date(Date.UTC(2025, 0, 1, index));
    const hour = date.getUTCHours();
    const productionByHour = [
      0, 0, 0, 0, 0, 0, 2, 9, 18, 28, 38, 45,
      50, 48, 39, 28, 14, 3, 0, 0, 0, 0, 0, 0,
    ];
    const standardRow = {
      "La date": `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}T${pad(date.getUTCHours())}:00:00`,
      production_mwh: productionByHour[hour],
      consumption_mwh: 15,
    };
    return useColumnVariants
      ? {
          Date: standardRow["La date"],
          "Production MWh": standardRow.production_mwh,
          "load_mwh": standardRow.consumption_mwh,
        }
      : standardRow;
  });
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) as Uint8Array;
}

function pad(value: number): string {
  return String(value).padStart(2, "0");
}
