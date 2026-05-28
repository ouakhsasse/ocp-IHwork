import Chart from "chart.js/auto";
import type { ChartConfiguration, Plugin } from "chart.js";
import type { PvgisProfileResult } from "./core/services/pvgisService.ts";
import type { ScadaMinuteState } from "./core/services/scadaSimulationService.ts";
import type { WeatherForecastResult } from "./core/services/weatherService.ts";
import type { ExcelMinuteScadaState, ExcelScadaSimulationResult } from "./core/services/excelScadaSimulationService.ts";
import { fetchPvgisHourlyPvProfile } from "./core/services/pvgisService.ts";
import { simulateScadaDay } from "./core/services/scadaSimulationService.ts";
import { fetchBenguerirWeatherForecast } from "./core/services/weatherService.ts";
import {
  EXCEL_SCADA_EXPECTED_HOURS,
  EXCEL_SCADA_EXPECTED_MINUTES,
  exportExcelHourlySummaryCsv,
  exportExcelMinuteLogCsv,
  exportExcelSummaryJson,
  getChartCursorHourPosition,
  getExcelMinutePointer,
  parseExcelScadaWorkbook,
  simulateExcelScadaFullYear,
} from "./core/services/excelScadaSimulationService.ts";

type DisplayMinuteState = ScadaMinuteState | ExcelMinuteScadaState;

let minuteStates: DisplayMinuteState[] = simulateScadaDay();
let currentMinute = 0;
let timerId: number | null = null;
let activeCharts: ScadaLineChart[] = [];
let latestPvgis: PvgisProfileResult | null = null;
let latestWeather: WeatherForecastResult | null = null;
let latestExcelResult: ExcelScadaSimulationResult | null = null;
let excelModeActive = false;
let chartWindowStart = 0;
let lastLoggedMinute: number | null = null;
interface MinuteLogEntry {
  html: string;
  mode: string;
  hasAlarm: boolean;
  hasBalanceError: boolean;
  isEvent: boolean;
}
const visibleMinuteLogRows: MinuteLogEntry[] = [];
let activeLogFilter = "all";
let chartWindowMode: "day" | "1h" | "6h" | "full-day" | "year" = "day";

const simulationDate = "2026-06-21";
type ScadaLineChart = Chart<"line", Array<number | null>, string>;
interface ChartView {
  states: DisplayMinuteState[];
  labels: string[];
  windowStartMinute: number;
}

document.addEventListener("DOMContentLoaded", () => {
  initScadaSupervision().catch((error) => {
    const message = error instanceof Error ? error.message : String(error);
    setBadge("emsStatusBadge", "WARNING", "warning");
    setText("weatherSummary", "SCADA initialization warning");
    setText("sunSummary", message);
  });
});

async function initScadaSupervision(): Promise<void> {
  bindControls();
  renderCharts();
  renderMinute(0);

  const [pvgis, weather] = await Promise.all([
    fetchPvgisHourlyPvProfile(),
    fetchBenguerirWeatherForecast(),
  ]);
  latestPvgis = pvgis;
  latestWeather = weather;
  minuteStates = simulateScadaDay(pvgis.hourlyProfile, {
    allowGridToBess: weather.allowGridToBessTonight,
    nextDayIrradiationMjM2: weather.tomorrow.shortwaveRadiationSumMjM2,
    nextDayCloudCoverPercent: weather.tomorrow.cloudCoverMeanPercent,
  });

  setBadge("pvgisStatus", pvgis.source === "pvgis" ? "PVGIS ONLINE" : "PVGIS FALLBACK MODE", pvgis.source === "pvgis" ? "online" : "warning");
  setBadge("weatherStatus", weather.source === "open-meteo" ? "Open-Meteo ONLINE" : "Open-Meteo FALLBACK MODE", weather.source === "open-meteo" ? "online" : "warning");
  updateWeatherSummary(weather);
  renderCharts();
  renderMinute(currentMinute);
}

function bindControls(): void {
  getButton("startButton").addEventListener("click", startSimulation);
  getButton("pauseButton").addEventListener("click", () => {
    pauseSimulation();
  });
  getButton("nextMinuteButton").addEventListener("click", () => {
    pauseSimulation();
    stepSimulation(1);
  });
  getButton("nextHourButton").addEventListener("click", () => {
    pauseSimulation();
    stepSimulation(60);
  });
  getButton("resetButton").addEventListener("click", () => {
    pauseSimulation();
    currentMinute = 0;
    renderMinute(currentMinute);
    setBadge("emsStatusBadge", "EMS READY", "ready");
  });
  getButton("exportCsvButton").addEventListener("click", exportCsv);
  getButton("exportJsonButton").addEventListener("click", exportJson);
  getButton("exportHourlyButton").addEventListener("click", exportHourlySummary);
  getButton("loadExcelButton").addEventListener("click", () => {
    void loadExcelData();
  });
  getElement<HTMLInputElement>("excelFileInput").addEventListener("change", (event) => {
    const input = event.currentTarget;
    if (input instanceof HTMLInputElement && input.files?.[0]) {
      void loadExcelData(input.files[0]);
    }
  });
  getButton("startExcelButton").addEventListener("click", startExcelMode);
  getButton("nextDayButton").addEventListener("click", () => {
    pauseSimulation();
    stepSimulation(1440);
  });
  getButton("runCurrentDayButton").addEventListener("click", runCurrentDay);
  getButton("runFullExcelButton").addEventListener("click", runFullExcel);
  getButton("resetExcelButton").addEventListener("click", resetExcelMode);
  getButton("toggleChartsButton").addEventListener("click", toggleCharts);
  getButton("clearVisibleLogButton").addEventListener("click", () => {
    visibleMinuteLogRows.length = 0;
    lastLoggedMinute = null;
    renderMinuteLog();
  });
  getButton("exportMinuteLogButton").addEventListener("click", exportCsv);
  getButton("compactLogButton").addEventListener("click", () => {
    const log = document.querySelector(".minute-log");
    log?.classList.toggle("compact-log");
    setText("compactLogButton", log?.classList.contains("compact-log") ? "Show all columns" : "Compact columns");
  });
  document.querySelectorAll<HTMLButtonElement>(".chart-window").forEach((button) => {
    button.addEventListener("click", () => {
      chartWindowMode = (button.dataset.window as typeof chartWindowMode) ?? "day";
      document.querySelectorAll<HTMLButtonElement>(".chart-window").forEach((item) => item.classList.toggle("active", item === button));
      renderCharts();
    });
  });
  document.querySelectorAll<HTMLButtonElement>(".log-filter").forEach((button) => {
    button.addEventListener("click", () => {
      activeLogFilter = button.dataset.filter ?? "all";
      document.querySelectorAll<HTMLButtonElement>(".log-filter").forEach((item) => item.classList.toggle("active", item === button));
      renderMinuteLog();
    });
  });
  document.querySelectorAll<HTMLElement>("#alarmList, #activeAlarmList").forEach((alarmContainer) => alarmContainer.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement) || !target.classList.contains("ack")) return;
    target.textContent = "ACKED";
    target.disabled = true;
    target.closest(".alarm")?.classList.add("acknowledged");
  }));
}

function startSimulation(): void {
  const speed = getElement<HTMLSelectElement>("speedSelect").value;
  setBadge("emsStatusBadge", "RUNNING", "running");

  if (speed === "instant-full") {
    pauseSimulation(false);
    currentMinute = getLastMinuteIndex();
    renderMinute(currentMinute);
    setBadge("emsStatusBadge", "EMS READY", "ready");
    return;
  }

  if (speed === "instant") {
    pauseSimulation(false);
    runCurrentDay();
    setBadge("emsStatusBadge", "EMS READY", "ready");
    return;
  }

  pauseSimulation(false);
  const minutesPerTick = Number(speed);
  timerId = window.setInterval(() => {
    stepSimulation(minutesPerTick);
  }, 1000);
}

function pauseSimulation(updateBadge = true): void {
  if (timerId !== null) {
    window.clearInterval(timerId);
    timerId = null;
  }
  if (updateBadge) {
    setBadge("emsStatusBadge", "PAUSED", "paused");
  }
}

function stepSimulation(minutes: number): void {
  const previousMinute = currentMinute;
  currentMinute = Math.min(getLastMinuteIndex(), currentMinute + minutes);
  renderMinute(currentMinute);
  appendTransitionMessages(previousMinute, currentMinute);
  if (currentMinute >= getLastMinuteIndex()) {
    pauseSimulation();
    if (excelModeActive) setBadge("emsStatusBadge", "FULL EXCEL SCADA TEST COMPLETED", "ready");
  }
}

function renderMinute(minute: number): void {
  const state = minuteStates[minute];
  if (!state) return;
  setText("simDateTime", isExcelState(state) ? state.timestamp.replace("T", " ") : `${simulationDate} ${state.timeLabel}`);
  setText("plantModeTag", `MODE ${state.emsMode}${state.supportMode ? " + D" : ""}`);
  setText("pvPower", formatNumber(state.pvMW, 2));
  setText("loadPower", formatNumber(state.loadMW, 2));
  setText("bessPower", formatSigned(state.bessPowerMW, 2));
  setText("socPercent", formatNumber(state.socPercent, 1));
  setText("gridImport", formatNumber(state.gridImportMW, 2));
  setText("wheelingExport", formatNumber(state.wheelingMW, 2));
  setText("currentTariff", state.tariff);
  setText("currentMode", state.emsMode);
  setText("netCost", formatNumber(state.netCostDhPerHour, 0));
  setText("dailyGain", formatNumber(state.cumulativeGainDh, 0));
  setText("activeCommand", state.command);
  setText("nodePv", `${formatNumber(state.pvMW, 2)} MW`);
  setText("nodeBess", `${formatSigned(state.bessPowerMW, 2)} MW`);
  setText("nodeBessSoc", `${formatNumber(state.socMWh, 1)} MWh | ${formatNumber(state.socPercent, 1)}%`);
  setText("nodeGrid", `${formatNumber(state.gridImportMW, 2)} MW`);
  setText("nodeWheeling", `${formatNumber(state.wheelingMW, 2)} MW`);
  setText("labelPvLoad", `${formatNumber(state.pvToLoadMW, 2)} MW`);
  setText("labelPvBess", `${formatNumber(state.pvToBessMW, 2)} MW`);
  setText("labelWheel", `${formatNumber(state.pvToWheelingMW, 2)} MW`);
  setText("labelBessLoad", `${formatNumber(state.bessToLoadMW, 2)} MW`);
  setText("labelGridLoad", `${formatNumber(state.gridToLoadMW, 2)} MW`);
  setText("labelGridBess", `${formatNumber(state.gridToBessMW, 2)} MW`);
  getElement<HTMLElement>("timelineCursor").style.setProperty("--cursor", `${((minute % 1440) / 1439) * 100}%`);

  updateFlow("flowPvInv", state.pvMW);
  updateFlow("flowInvLoad", state.pvToLoadMW);
  updateFlow("flowPvBess", state.pvToBessMW);
  updateFlow("flowInvWheel", state.pvToWheelingMW);
  updateFlow("flowBessLoad", state.bessToLoadMW);
  updateFlow("flowGridLoad", state.gridToLoadMW);
  updateFlow("flowGridBess", state.gridToBessMW);
  updatePeriodOverlay(state);
  renderAlarms(state);
  renderExcelProgress(state);
  renderLiveMinuteState(state);
  renderEnergyThisMinute(state);
  renderEnergyInterpretation(state);
  renderCurrentMinuteSummary(state);
  appendMinuteLogRow(state);
  setText("chartCursorLabel", `Current minute: ${state.timeLabel} | x=${formatNumber(getChartCursorHourPosition(state.minute), 2)}h`);
  updateChartsLive();
}

function updateFlow(id: string, valueMW: number): void {
  const element = getElement<SVGPathElement>(id);
  element.classList.toggle("active", valueMW > 0.05);
  element.style.strokeWidth = String(2 + Math.min(7, valueMW / 8));
  element.style.animationDuration = `${Math.max(0.35, 1.4 - Math.min(1, valueMW / 35))}s`;
}

function updatePeriodOverlay(state: ScadaMinuteState): void {
  const overlay = getElement<HTMLElement>("periodOverlay");
  overlay.className = "period-overlay";
  if (state.tariff === "Pointe") {
    overlay.classList.add("peak");
  } else if (state.emsMode === "C" || state.pvMW <= 0.05) {
    overlay.classList.add("night");
  }
}

function renderAlarms(state: DisplayMinuteState): void {
  const alarms = [
    { name: "Low SoC", severity: "WARNING", active: state.socPercent <= 18 },
    { name: "High SoC", severity: "INFO", active: state.socPercent >= 92 },
    { name: "Peak tariff period active - discharge priority enabled", severity: "INFO", active: state.tariff === "Pointe" },
    { name: "PV ramp detected - smoothing support active", severity: "WARNING", active: state.rampEvent },
    { name: "External API unavailable - local fallback profile in service", severity: "WARNING", active: latestPvgis?.source === "fallback" || latestWeather?.source === "fallback" },
    { name: "Grid import above supervision threshold", severity: "CRITICAL", active: state.gridImportMW >= 20 },
    { name: "ENERGY BALANCE ERROR", severity: "CRITICAL", active: isExcelState(state) && Math.abs(state.energyBalanceError) > 1e-8 },
  ].filter((alarm) => alarm.active);

  const alarmList = getElement<HTMLElement>("alarmList");
  const alarmHtml = alarms.length > 0
    ? alarms.map((alarm) => renderAlarm(alarm.name, alarm.severity, state.timeLabel)).join("")
    : `<div class="alarm"><div class="alarm-head"><span>All monitored conditions normal</span><span>INFO</span></div><div class="alarm-time">${simulationDate} ${state.timeLabel}</div></div>`;
  alarmList.innerHTML = alarmHtml;
  setElementHtml("activeAlarmList", alarmHtml);
  setText("alarmCount", `${alarms.length} ACTIVE`);
  setText("alarmHistoryCount", `${alarms.length} ACTIVE`);
}

function renderAlarm(name: string, severity: string, timeLabel: string): string {
  const severityClass = severity === "CRITICAL" ? "critical" : severity === "WARNING" ? "warning" : "";
  return `
    <div class="alarm ${severityClass}">
      <div class="alarm-head"><span>${escapeHtml(name)}</span><span>${escapeHtml(severity)}</span></div>
      <div class="alarm-time">${simulationDate} ${timeLabel}</div>
      <button class="ack" type="button">ACK</button>
    </div>
  `;
}

function renderCharts(): void {
  activeCharts.forEach((chart) => chart.destroy());
  activeCharts = [];
  const view = getChartView();
  chartWindowStart = view.windowStartMinute;

  createChart("loadSupplyChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        areaDataset("PV to Load", chartValues(view, (state) => state.pvToLoadMW), "#f59e0b", "loadSupply"),
        areaDataset("BESS to Load", chartValues(view, (state) => state.bessToLoadMW), "#818cf8", "loadSupply"),
        areaDataset("Grid to Load", chartValues(view, (state) => state.gridToLoadMW), "#ef4444", "loadSupply"),
      ],
    },
    options: chartOptions("MW", true),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  createChart("bessChargeChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        areaDataset("BESS charging", chartValues(view, (state) => Math.max(0, state.bessPowerMW)), "#22c55e"),
        areaDataset("BESS discharging", chartValues(view, (state) => Math.min(0, state.bessPowerMW)), "#ef4444"),
      ],
    },
    options: chartOptions("MW"),
    plugins: [minuteCursorPlugin, zeroLinePlugin, transitionMarkerPlugin],
  });

  createChart("pvAllocationChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        areaDataset("PV to Load", chartValues(view, (state) => state.pvToLoadMW), "#f59e0b", "pvAllocation"),
        areaDataset("PV to BESS", chartValues(view, (state) => state.pvToBessMW), "#818cf8", "pvAllocation"),
        areaDataset("PV to Wheeling", chartValues(view, (state) => state.pvToWheelingMW), "#22c55e", "pvAllocation"),
        areaDataset("Curtailment", chartValues(view, () => 0), "#94a3b8", "pvAllocation"),
      ],
    },
    options: chartOptions("MW", true),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  createChart("gridBreakdownChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        areaDataset("Grid to Load", chartValues(view, (state) => state.gridToLoadMW), "#ef4444"),
        areaDataset("Grid to BESS", chartValues(view, (state) => state.gridToBessMW), "#f59e0b"),
        lineDataset("Total Grid Import", chartValues(view, (state) => state.gridImportMW), "#e5e7eb"),
      ],
    },
    options: chartOptions("MW"),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  createChart("socChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        lineDataset("SoC %", chartValues(view, (state) => state.socPercent), "#8b5cf6"),
      ],
    },
    options: chartOptions("%"),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  createChart("economicChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        lineDataset("Baseline cost", chartValues(view, cumulativeBaselineCostDh), "#94a3b8"),
        lineDataset("EMS cost", chartValues(view, cumulativeEmsCostDh), "#f59e0b"),
        lineDataset("Cumulative gain", chartValues(view, (state) => state.cumulativeGainDh), "#22c55e"),
      ],
    },
    options: chartOptions("DH"),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  createChart("minuteCostChart", {
    type: "line",
    data: {
      labels: view.labels,
      datasets: [
        lineDataset("Baseline cost DH/min", chartValues(view, baselineCostDhMin), "#94a3b8"),
        lineDataset("EMS cost DH/min", chartValues(view, emsCostDhMin), "#f59e0b"),
        areaDataset("Avoided cost / gain DH/min", chartValues(view, gainDhMin), "#22c55e"),
      ],
    },
    options: chartOptions("DH/min"),
    plugins: [minuteCursorPlugin, transitionMarkerPlugin],
  });

  updateLatestChartBadges();
}

function updateChartsLive(): void {
  if (activeCharts.length === 0) {
    renderCharts();
    return;
  }
  const view = getChartView();
  if (view.windowStartMinute !== chartWindowStart || activeCharts[0]?.data.labels?.length !== view.labels.length) {
    renderCharts();
    return;
  }

  const chartData = [
    [
      chartValues(view, (state) => state.pvToLoadMW),
      chartValues(view, (state) => state.bessToLoadMW),
      chartValues(view, (state) => state.gridToLoadMW),
    ],
    [
      chartValues(view, (state) => Math.max(0, state.bessPowerMW)),
      chartValues(view, (state) => Math.min(0, state.bessPowerMW)),
    ],
    [
      chartValues(view, (state) => state.pvToLoadMW),
      chartValues(view, (state) => state.pvToBessMW),
      chartValues(view, (state) => state.pvToWheelingMW),
      chartValues(view, () => 0),
    ],
    [
      chartValues(view, (state) => state.gridToLoadMW),
      chartValues(view, (state) => state.gridToBessMW),
      chartValues(view, (state) => state.gridImportMW),
    ],
    [
      chartValues(view, (state) => state.socPercent),
    ],
    [
      chartValues(view, cumulativeBaselineCostDh),
      chartValues(view, cumulativeEmsCostDh),
      chartValues(view, (state) => state.cumulativeGainDh),
    ],
    [
      chartValues(view, baselineCostDhMin),
      chartValues(view, emsCostDhMin),
      chartValues(view, gainDhMin),
    ],
  ];

  activeCharts.forEach((chart, chartIndex) => {
    chart.data.labels = view.labels;
    chartData[chartIndex]?.forEach((datasetData, datasetIndex) => {
      if (chart.data.datasets[datasetIndex]) chart.data.datasets[datasetIndex].data = datasetData;
    });
    chart.update("none");
  });
  updateLatestChartBadges();
}

const minuteCursorPlugin: Plugin<"line"> = {
  id: "minuteCursor",
  afterDraw(chart) {
    const xScale = chart.scales.x;
    const area = chart.chartArea;
    if (!xScale || !area) return;
    const chartIndex = getCurrentChartIndex(Number(chart.data.labels?.length ?? 0));
    const x = xScale.getPixelForValue(chartIndex);
    const ctx = chart.ctx;
    ctx.save();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(x, area.top);
    ctx.lineTo(x, area.bottom);
    ctx.stroke();
    ctx.restore();
  },
};

const zeroLinePlugin: Plugin<"line"> = {
  id: "zeroLine",
  afterDraw(chart) {
    const yScale = chart.scales.y;
    const area = chart.chartArea;
    if (!yScale || !area) return;
    const y = yScale.getPixelForValue(0);
    const ctx = chart.ctx;
    ctx.save();
    ctx.strokeStyle = "rgba(226, 232, 240, 0.85)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(area.left, y);
    ctx.lineTo(area.right, y);
    ctx.stroke();
    ctx.restore();
  },
};

const transitionMarkerPlugin: Plugin<"line"> = {
  id: "transitionMarkers",
  afterDraw(chart) {
    const xScale = chart.scales.x;
    const area = chart.chartArea;
    if (!xScale || !area) return;
    const markers = [
      { hour: 7, label: "Mode A", detail: "PV autoconsumption / charge BESS", color: "#22c55e" },
      { hour: 17, label: "Mode B", detail: "Peak shaving / discharge BESS", color: "#ef4444" },
      { hour: 22, label: "Mode C", detail: "Night preparation", color: "#818cf8" },
      { hour: 23, label: "Creuses", detail: "off-peak period", color: "#94a3b8" },
    ];
    const ctx = chart.ctx;
    ctx.save();
    ctx.font = "9px Consolas, monospace";
    markers.forEach((marker) => {
      const index = getMarkerIndex(marker.hour);
      if (index < 0 || index >= Number(chart.data.labels?.length ?? 0)) return;
      const x = xScale.getPixelForValue(index);
      if (x < area.left || x > area.right) return;
      ctx.strokeStyle = marker.color;
      ctx.fillStyle = marker.color;
      ctx.globalAlpha = 0.78;
      ctx.lineWidth = 1;
      ctx.setLineDash([2, 3]);
      ctx.beginPath();
      ctx.moveTo(x, area.top);
      ctx.lineTo(x, area.bottom);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.save();
      ctx.translate(x + 3, area.top + 52);
      ctx.rotate(-Math.PI / 2);
      ctx.fillText(`${marker.label}: ${marker.detail}`, 0, 0);
      ctx.restore();
    });
    ctx.restore();
  },
};

function createChart(canvasId: string, config: ChartConfiguration<"line", Array<number | null>, string>): void {
  activeCharts.push(new Chart(getElement<HTMLCanvasElement>(canvasId), config));
}

function lineDataset(label: string, data: Array<number | null>, color: string) {
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: transparentize(color, 0.12),
    borderWidth: 1.8,
    pointRadius: 0,
    tension: 0.18,
    spanGaps: false,
  };
}

function areaDataset(label: string, data: Array<number | null>, color: string, stack?: string) {
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: transparentize(color, 0.36),
    borderWidth: 1.4,
    pointRadius: 0,
    tension: 0.12,
    fill: true,
    stack,
  };
}

function chartOptions(yTitle: string, stacked = false) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false as const,
    interaction: { intersect: false, mode: "index" as const },
    plugins: {
      legend: {
        display: true,
        position: "bottom" as const,
        labels: { color: "#9fb2c6", boxWidth: 10, font: { size: 10 } },
      },
      tooltip: {
        callbacks: {
          label: (context: { dataset: { label?: string }; raw: unknown }) => `${context.dataset.label ?? ""}: ${formatNumber(Number(context.raw), yTitle === "DH" ? 0 : 2)} ${yTitle}`,
        },
      },
    },
    scales: {
      x: {
        grid: { color: "rgba(148, 163, 184, 0.12)" },
        ticks: { color: "#6f849b", maxRotation: 0, autoSkip: true, font: { size: 9 } },
      },
      y: {
        stacked,
        grid: { color: "rgba(148, 163, 184, 0.12)" },
        ticks: { color: "#6f849b", font: { size: 9 } },
        title: { display: false, text: yTitle },
      },
    },
  };
}

function getChartView(): ChartView {
  if (chartWindowMode === "year" && excelModeActive) {
    const states = minuteStates.filter((_, index) => index % 60 === 0 && index <= currentMinute);
    return {
      states,
      labels: states.map((_, index) => index % 24 === 0 ? `D${Math.floor(index / 24) + 1}` : ""),
      windowStartMinute: states[0]?.minute ?? 0,
    };
  }

  const dayStart = Math.floor(currentMinute / 1440) * 1440;
  const dayEnd = Math.min(getLastMinuteIndex(), dayStart + 1439);
  let start = dayStart;
  let end = dayEnd;
  if (chartWindowMode === "1h") start = Math.max(dayStart, currentMinute - 59);
  if (chartWindowMode === "6h") start = Math.max(dayStart, currentMinute - 359);
  if (chartWindowMode === "day" || chartWindowMode === "full-day") {
    start = dayStart;
    end = dayEnd;
  } else {
    end = currentMinute;
  }
  const states = minuteStates.slice(start, end + 1);
  return {
    states,
    labels: states.map((state, index) => index % 60 === 0 || states.length <= 90 ? state.timeLabel : ""),
    windowStartMinute: start,
  };
}

function chartValues(view: ChartView, getValue: (state: DisplayMinuteState) => number): Array<number | null> {
  return view.states.map((state) => state.minute <= currentMinute ? finiteOrZero(getValue(state)) : null);
}

function getMarkerIndex(hour: number): number {
  if (chartWindowMode === "year") return -1;
  const markerMinute = Math.floor(currentMinute / 1440) * 1440 + hour * 60;
  return markerMinute - chartWindowStart;
}

function getCurrentChartIndex(labelCount: number): number {
  if (chartWindowMode === "year") return Math.max(0, labelCount - 1);
  return currentMinute - chartWindowStart;
}

function cumulativeBaselineCostDh(state: DisplayMinuteState): number {
  if (isExcelState(state)) return state.cumulativeBaselineCostDh;
  return cumulativeCostUntil(state.minute, "baseline");
}

function baselineCostDhMin(state: DisplayMinuteState): number {
  if (isExcelState(state)) return state.baselineCostDh;
  return (state.loadMW / 60) * 1000 * state.tariffDhPerKWh;
}

function emsCostDhMin(state: DisplayMinuteState): number {
  if (isExcelState(state)) return state.costDh;
  return (state.gridImportMW / 60) * 1000 * state.tariffDhPerKWh - (state.wheelingMW / 60) * 1000 * 0.7131;
}

function gainDhMin(state: DisplayMinuteState): number {
  return baselineCostDhMin(state) - emsCostDhMin(state);
}

function cumulativeEmsCostDh(state: DisplayMinuteState): number {
  if (isExcelState(state)) return state.cumulativeEmsCostDh;
  return cumulativeCostUntil(state.minute, "ems");
}

function cumulativeCostUntil(minute: number, kind: "baseline" | "ems"): number {
  let total = 0;
  for (let index = 0; index <= minute && index < minuteStates.length; index += 1) {
    const state = minuteStates[index];
    const baseline = (state.loadMW / 60) * 1000 * state.tariffDhPerKWh;
    const ems = (state.gridImportMW / 60) * 1000 * state.tariffDhPerKWh;
    total += kind === "baseline" ? baseline : ems;
  }
  return Math.round(total * 1000) / 1000;
}

function updateLatestChartBadges(): void {
  const state = minuteStates[currentMinute];
  if (!state) return;
  setText("supplyMixLatest", `PV ${formatNumber(state.pvToLoadMW, 1)} | BESS ${formatNumber(state.bessToLoadMW, 1)} | Grid ${formatNumber(state.gridToLoadMW, 1)} MW`);
  setText("bessPowerLatest", `${state.bessPowerMW >= 0 ? "Charge" : "Discharge"} ${formatSigned(state.bessPowerMW, 1)} MW`);
  setText("pvAllocationLatest", `Load ${formatNumber(state.pvToLoadMW, 1)} | BESS ${formatNumber(state.pvToBessMW, 1)} | Wheel ${formatNumber(state.pvToWheelingMW, 1)} MW`);
  setText("gridBreakdownLatest", `Load ${formatNumber(state.gridToLoadMW, 1)} | BESS ${formatNumber(state.gridToBessMW, 1)} | Total ${formatNumber(state.gridImportMW, 1)} MW`);
  setText("socLatest", `${formatNumber(state.socPercent, 1)}%`);
  setText("economicLatest", `Gain ${formatNumber(state.cumulativeGainDh, 0)} DH`);
  setText("minuteCostLatest", `Baseline ${formatNumber(baselineCostDhMin(state), 1)} | EMS ${formatNumber(emsCostDhMin(state), 1)} | Gain ${formatNumber(gainDhMin(state), 1)} DH/min`);
}

function renderEnergyInterpretation(state: DisplayMinuteState): void {
  const balanceOk = !isExcelState(state) || state.energyBalanceOk;
  let text = "OCP load is currently supplied by ONEE grid because PV = 0 MW and BESS is in standby.";
  if (!balanceOk) {
    text = "ENERGY BALANCE ERROR detected. Verify PV, BESS, grid, load, and wheeling flows for this minute.";
  } else if (state.bessToLoadMW > 0.05 && state.tariff === "Pointe") {
    text = `BESS is discharging at ${formatNumber(state.bessToLoadMW, 1)} MW during peak tariff to reduce ONEE grid import.`;
  } else if (state.pvToBessMW > 0.05) {
    text = `PV surplus is charging the BESS at ${formatNumber(state.pvToBessMW, 1)} MW after supplying the OCP load.`;
  } else if (state.gridToBessMW > 0.05) {
    text = `Off-peak ONEE energy is charging the BESS at ${formatNumber(state.gridToBessMW, 1)} MW because Mode C night charging is enabled.`;
  } else if (state.pvToWheelingMW > 0.05) {
    text = `BESS is full or charge-limited, so PV surplus is exported through wheeling at ${formatNumber(state.pvToWheelingMW, 1)} MW.`;
  } else if (state.pvToLoadMW > 0.05 && state.gridToLoadMW > 0.05) {
    text = `OCP load is supplied by PV (${formatNumber(state.pvToLoadMW, 1)} MW) and ONEE grid (${formatNumber(state.gridToLoadMW, 1)} MW).`;
  } else if (state.pvToLoadMW >= state.loadMW - 0.05) {
    text = "OCP load is fully supplied by PV production for this minute.";
  } else if (state.gridToBessMW <= 0.05 && (state.tariff === "Pleines" || state.tariff === "Pointe")) {
    text = `Grid-to-BESS charging is locked because current tariff is ${state.tariff}.`;
  }
  setText("energyInterpretation", text);
}

function transparentize(hex: string, alpha: number): string {
  const normalized = hex.replace("#", "");
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function finiteOrZero(value: number): number {
  return Number.isFinite(value) ? value : 0;
}

function updateWeatherSummary(weather: WeatherForecastResult): void {
  const tomorrow = weather.tomorrow;
  setText("weatherSummary", `J+1 ${formatNumber(tomorrow.shortwaveRadiationSumMjM2, 1)} MJ/m2 | Clouds ${formatNumber(tomorrow.cloudCoverMeanPercent, 0)}% | Tmax ${formatNumber(tomorrow.temperatureMaxC, 1)} C`);
  setText("sunSummary", `Sunrise ${formatClock(weather.today.sunrise)} | Sunset ${formatClock(weather.today.sunset)} | Mode C grid charge ${weather.allowGridToBessTonight ? "ENABLED" : "LOCKED"}`);
}

async function loadExcelData(selectedFile?: File): Promise<void> {
  const input = getElement<HTMLInputElement>("excelFileInput");
  const file = selectedFile ?? input.files?.[0] ?? null;
  if (!file) {
    setText("excelStatus", "Select data base 2025(1)mod.xlsx first.");
    return;
  }

  try {
    console.log("Selected Excel file:", file.name);
    setText("excelStatus", "Loading Excel workbook...");
    const buffer = await file.arrayBuffer();
    const hourlyRecords = parseExcelScadaWorkbook(buffer, file.name);
    latestExcelResult = simulateExcelScadaFullYear(hourlyRecords, {
      allowNightGridCharging: getElement<HTMLInputElement>("allowNightGridCharging").checked,
      initialSocPercent: 0,
    }, file.name);
    console.log("Minute records:", latestExcelResult.minuteStates.length);
    setText("excelStatus", `EXCEL DATA LOADED — ${formatNumber(hourlyRecords.length, 0)} HOURS / ${formatNumber(latestExcelResult.minuteStates.length, 0)} MINUTES`);
    setBadge("pvgisStatus", "PVGIS FALLBACK: EXCEL PROFILE", "warning");
    startExcelMode();
  } catch (error) {
    latestExcelResult = null;
    excelModeActive = false;
    setText("excelStatus", formatExcelLoadError(error));
    setBadge("emsStatusBadge", "WARNING", "warning");
  }
}

function startExcelMode(): void {
  if (!latestExcelResult) {
    setText("excelStatus", "Load a valid 8760-hour Excel workbook first.");
    return;
  }

  pauseSimulation(false);
  excelModeActive = true;
  minuteStates = latestExcelResult.minuteStates;
  currentMinute = 0;
  chartWindowStart = 0;
  lastLoggedMinute = null;
  visibleMinuteLogRows.length = 0;
  renderMinuteLog();
  renderCharts();
  renderMinute(currentMinute);
  renderExcelKpis();
  setBadge("emsStatusBadge", "EMS READY", "ready");
}

function resetExcelMode(): void {
  if (!latestExcelResult) {
    setText("excelStatus", "Load Excel data before resetting full-year mode.");
    return;
  }
  latestExcelResult = simulateExcelScadaFullYear(latestExcelResult.hourlyRecords, {
    allowNightGridCharging: getElement<HTMLInputElement>("allowNightGridCharging").checked,
    initialSocPercent: 0,
  }, latestExcelResult.sourceName);
  startExcelMode();
}

function runCurrentDay(): void {
  const nextDayEnd = Math.min(getLastMinuteIndex(), Math.floor(currentMinute / 1440) * 1440 + 1439);
  currentMinute = nextDayEnd;
  renderMinute(currentMinute);
}

function runFullExcel(): void {
  if (!excelModeActive) startExcelMode();
  if (!excelModeActive) return;
  currentMinute = getLastMinuteIndex();
  renderMinute(currentMinute);
  setBadge("emsStatusBadge", "FULL EXCEL SCADA TEST COMPLETED", "ready");
}

function renderExcelProgress(state: DisplayMinuteState): void {
  if (!excelModeActive || !isExcelState(state)) {
    setText("excelProgress", "Waiting for 8760-hour workbook");
    return;
  }

  const completion = ((state.minute + 1) / EXCEL_SCADA_EXPECTED_MINUTES) * 100;
  const remainingMinutes = EXCEL_SCADA_EXPECTED_MINUTES - state.minute - 1;
  setText("excelProgress", [
    `Excel row ${formatNumber(state.excelRowIndex + 1, 0)} / ${formatNumber(EXCEL_SCADA_EXPECTED_HOURS, 0)}`,
    `Minute step ${formatNumber(state.minute + 1, 0)} / ${formatNumber(EXCEL_SCADA_EXPECTED_MINUTES, 0)}`,
    `Day ${formatNumber(state.dayNumber, 0)} / 365`,
    `Completion ${formatNumber(completion, 2)}%`,
    `Remaining ${formatNumber(remainingMinutes, 0)} simulated minutes`,
  ].join("\n"));
}

function renderLiveMinuteState(state: DisplayMinuteState): void {
  const balanceOk = !isExcelState(state) || Math.abs(state.energyBalanceError) <= 1e-8;
  if (isExcelState(state) && latestExcelResult) {
    const pointer = getExcelMinutePointer(latestExcelResult.hourlyRecords, state.minute);
    setElementHtml("liveMinuteState", [
      ["Current minute", pointer.timestamp.replace("T", " ").slice(0, 16)],
      ["Excel row", `${formatNumber(pointer.hourIndex + 1, 0)} / ${formatNumber(EXCEL_SCADA_EXPECTED_HOURS, 0)}`],
      ["Minute inside hour", `${formatNumber(pointer.minuteInsideHour, 0)} / 59`],
      ["Global minute", `${formatNumber(pointer.currentMinuteIndex + 1, 0)} / ${formatNumber(pointer.totalMinutes, 0)}`],
      ["Day", `${formatNumber(pointer.dayNumber, 0)} / 365`],
      ["Progress", `${formatNumber(pointer.progressPercent, 2)}%`],
      ["Tariff", state.tariff],
      ["EMS mode", state.emsMode],
      ["Command", state.command],
      ["Balance", balanceOk ? "OK" : "ERROR"],
    ].map(([label, value]) => `<span>${escapeHtml(label)}</span><span>${escapeHtml(value)}</span>`).join(""));
    return;
  }

  setElementHtml("liveMinuteState", [
    ["Current minute", `${simulationDate} ${state.timeLabel}`],
    ["Excel row", "--"],
    ["Minute inside hour", `${formatNumber(state.minute % 60, 0)} / 59`],
    ["Global minute", `${formatNumber(state.minute + 1, 0)} / 1,440`],
    ["Day", "1 / 1"],
    ["Progress", `${formatNumber(((state.minute + 1) / 1440) * 100, 2)}%`],
    ["Tariff", state.tariff],
    ["EMS mode", state.emsMode],
    ["Command", state.command],
    ["Balance", "OK"],
  ].map(([label, value]) => `<span>${escapeHtml(label)}</span><span>${escapeHtml(value)}</span>`).join(""));
}

function renderEnergyThisMinute(state: DisplayMinuteState): void {
  const pvMWh = isExcelState(state) ? state.pvMWh : state.pvMW / 60;
  const loadMWh = isExcelState(state) ? state.loadMWh : state.loadMW / 60;
  const gridMWh = isExcelState(state) ? state.gridImportMWh : state.gridImportMW / 60;
  const bessMWh = isExcelState(state) ? state.bessChargeMWh - state.bessDischargeMWh : state.bessPowerMW / 60;
  const wheelingMWh = isExcelState(state) ? state.wheelingMWh : state.wheelingMW / 60;
  const costDh = isExcelState(state) ? state.costDh : state.netCostDhPerHour / 60;
  const gainDh = isExcelState(state) ? state.baselineCostDh - state.costDh : 0;
  setElementHtml("energyThisMinute", [
    ["PV", `${formatNumber(pvMWh, 3)} MWh`],
    ["Load", `${formatNumber(loadMWh, 3)} MWh`],
    ["Grid", `${formatNumber(gridMWh, 3)} MWh`],
    ["BESS", `${formatSigned(bessMWh, 3)} MWh`],
    ["Wheeling", `${formatNumber(wheelingMWh, 3)} MWh`],
    ["Cost", `${formatNumber(costDh, 3)} DH/min`],
    ["Gain", `${formatNumber(gainDh, 3)} DH/min`],
  ].map(([label, value]) => `<span>${escapeHtml(label)}</span><span>${escapeHtml(value)}</span>`).join(""));
}

function renderCurrentMinuteSummary(state: DisplayMinuteState): void {
  const timestamp = isExcelState(state) ? state.timestamp.replace("T", " ").slice(0, 16) : `${simulationDate} ${state.timeLabel}`;
  const excelRow = isExcelState(state) ? `${formatNumber(state.excelRowIndex + 1, 0)} / ${formatNumber(EXCEL_SCADA_EXPECTED_HOURS, 0)}` : "--";
  const balance = !isExcelState(state) || Math.abs(state.energyBalanceError) <= 1e-8 ? "OK" : "ERROR";
  const items = [
    ["Timestamp", timestamp],
    ["Minute index", formatNumber(state.minute + 1, 0)],
    ["Excel row", excelRow],
    ["PV MW", formatNumber(state.pvMW, 2)],
    ["Load MW", formatNumber(state.loadMW, 2)],
    ["BESS MW", formatSigned(state.bessPowerMW, 2)],
    ["SoC %", formatNumber(state.socPercent, 1)],
    ["Grid MW", formatNumber(state.gridImportMW, 2)],
    ["Wheeling MW", formatNumber(state.wheelingMW, 2)],
    ["Tariff", state.tariff],
    ["EMS mode", `${state.emsMode}${state.supportMode ? " + D" : ""}`],
    ["Command", state.command],
    ["Balance", balance],
  ];
  setElementHtml("currentMinuteSummary", items.map(([label, value]) => `
    <div class="summary-item">
      <div class="summary-label">${escapeHtml(label)}</div>
      <div class="summary-value">${escapeHtml(value)}</div>
    </div>
  `).join(""));
}

function toggleCharts(): void {
  const section = getElement<HTMLElement>("chartsSection");
  const collapsed = section.classList.toggle("charts-collapsed");
  setText("toggleChartsButton", collapsed ? "Show charts" : "Hide charts");
  if (!collapsed) {
    activeCharts.forEach((chart) => chart.resize());
  }
}

function tariffBadge(tariff: string): string {
  const className = tariff === "Pointe" ? "pointe" : tariff === "Pleines" ? "pleines" : "creuses";
  return `<span class="log-badge ${className}">${escapeHtml(tariff)}</span>`;
}

function modeBadge(mode: string, supportMode: "D" | null): string {
  return `<span class="log-badge mode">${escapeHtml(`${mode}${supportMode ? " + D" : ""}`)}</span>`;
}

function appendMinuteLogRow(state: DisplayMinuteState): void {
  if (lastLoggedMinute === state.minute) return;
  lastLoggedMinute = state.minute;
  const minuteNumber = state.minute + 1;
  const excelRow = isExcelState(state) ? state.excelRowIndex + 1 : Math.floor(state.minute / 60) + 1;
  const costDh = isExcelState(state) ? state.costDh : state.netCostDhPerHour / 60;
  const gainDh = isExcelState(state) ? state.baselineCostDh - state.costDh : 0;
  const balance = !isExcelState(state) || Math.abs(state.energyBalanceError) <= 1e-8 ? "OK" : "ERROR";
  const hasAlarm = state.rampEvent || state.gridImportMW >= 20 || state.socPercent <= 18 || state.socPercent >= 92 || balance === "ERROR";
  visibleMinuteLogRows.push({
    mode: state.emsMode,
    hasAlarm,
    hasBalanceError: balance === "ERROR",
    isEvent: false,
    html: `
    <tr class="latest-row">
      <td>${escapeHtml(isExcelState(state) ? state.timestamp.replace("T", " ").slice(0, 16) : `${simulationDate} ${state.timeLabel}`)}</td>
      <td>${formatNumber(minuteNumber, 0)}</td>
      <td>${formatNumber(excelRow, 0)}</td>
      <td>${formatNumber(state.pvMW, 2)}</td>
      <td>${formatNumber(state.loadMW, 2)}</td>
      <td>${formatSigned(state.bessPowerMW, 2)}</td>
      <td>${formatNumber(state.socPercent, 1)}</td>
      <td>${formatNumber(state.gridImportMW, 2)}</td>
      <td>${formatNumber(state.wheelingMW, 2)}</td>
      <td>${tariffBadge(state.tariff)}</td>
      <td>${modeBadge(state.emsMode, state.supportMode)}</td>
      <td>${escapeHtml(state.command)}</td>
      <td>${formatNumber(costDh, 3)}</td>
      <td>${formatNumber(gainDh, 3)}</td>
      <td><span class="${balance === "OK" ? "balance-ok" : "balance-error"}">${balance}</span></td>
    </tr>
  `,
  });
  while (visibleMinuteLogRows.length > 120) visibleMinuteLogRows.shift();
  renderMinuteLog();
}

function appendTransitionMessages(previousMinute: number, nextMinute: number): void {
  const start = previousMinute + 1;
  for (let minute = start; minute <= nextMinute; minute += 1) {
    const state = minuteStates[minute];
    if (!state || minute % 60 !== 0) continue;
    const time = state.timeLabel;
    const message = getTransitionMessage(time);
    if (!message) continue;
    visibleMinuteLogRows.push({
      html: `<tr class="event-row"><td colspan="15">${escapeHtml(isExcelState(state) ? state.timestamp.replace("T", " ").slice(0, 16) : `${simulationDate} ${time}`)} - ${escapeHtml(message)}</td></tr>`,
      mode: state.emsMode,
      hasAlarm: true,
      hasBalanceError: false,
      isEvent: true,
    });
    while (visibleMinuteLogRows.length > 120) visibleMinuteLogRows.shift();
  }
  renderMinuteLog();
}

function getTransitionMessage(timeLabel: string): string | null {
  if (timeLabel === "07:00") return "Mode A started - PV autoconsumption and BESS charging enabled";
  if (timeLabel === "17:00") return "Mode B started - peak shaving by BESS discharge";
  if (timeLabel === "22:00") return "Mode C started - night preparation mode";
  if (timeLabel === "23:00") return "Creuses started - off-peak grid charging allowed only if enabled and forecast bad";
  return null;
}

function renderMinuteLog(): void {
  const filteredRows = visibleMinuteLogRows.filter((row) => {
    if (activeLogFilter === "all") return true;
    if (activeLogFilter === "alarms") return row.hasAlarm;
    if (activeLogFilter === "balance") return row.hasBalanceError;
    return row.mode === activeLogFilter;
  });
  setElementHtml("minuteLogBody", filteredRows.map((row, index) => row.html.replace("latest-row", index === filteredRows.length - 1 ? "latest-row" : "")).join(""));
  const log = document.querySelector(".minute-log");
  if (log && getElement<HTMLInputElement>("autoScrollLog").checked) log.scrollTop = log.scrollHeight;
}

function renderExcelKpis(): void {
  if (!latestExcelResult) return;
  const kpis = latestExcelResult.kpis;
  setText("excelKpis", [
    `PV ${formatNumber(kpis.totalPvProductionMWh, 1)} MWh`,
    `Load ${formatNumber(kpis.totalOcpConsumptionMWh, 1)} MWh`,
    `Grid ${formatNumber(kpis.totalGridImportMWh, 1)} MWh`,
    `BESS charge ${formatNumber(kpis.totalBessChargeMWh, 1)} MWh`,
    `BESS discharge ${formatNumber(kpis.totalBessDischargeMWh, 1)} MWh`,
    `Wheeling ${formatNumber(kpis.totalWheelingMWh, 1)} MWh`,
    `Final SoC ${formatNumber(kpis.finalSocPercent, 1)}%`,
    `Cycles ${formatNumber(kpis.equivalentBessCycles, 2)}`,
    `EMS cost ${formatNumber(kpis.totalCostWithEmsDh, 0)} DH`,
    `Baseline ${formatNumber(kpis.baselineCostWithoutEmsDh, 0)} DH`,
    `Gain ${formatNumber(kpis.totalEconomicGainDh, 0)} DH`,
    `Alarms ${formatNumber(kpis.alarmCount, 0)} | Balance errors ${formatNumber(kpis.energyBalanceErrorCount, 0)}`,
  ].join("\n"));
}

function getLastMinuteIndex(): number {
  return Math.max(0, minuteStates.length - 1);
}

function isExcelState(state: DisplayMinuteState): state is ExcelMinuteScadaState {
  return "excelRowIndex" in state;
}

function formatExcelLoadError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (message.toLowerCase().includes("xlsx")) return "XLSX library not loaded or workbook parsing failed.";
  return message || "File reading failed.";
}

function exportCsv(): void {
  if (excelModeActive && latestExcelResult) {
    downloadText(
      `ocp_excel_scada_minute_log_${latestExcelResult.minuteStates[0].timestamp.slice(0, 10)}.csv`,
      exportExcelMinuteLogCsv(latestExcelResult.minuteStates),
      "text/csv",
    );
    return;
  }

  const header = [
    "minute",
    "time",
    "pv_mw",
    "load_mw",
    "bess_power_mw",
    "soc_percent",
    "grid_import_mw",
    "wheeling_mw",
    "tariff",
    "ems_mode",
    "command",
    "cumulative_gain_dh",
  ];
  const rows = minuteStates.map((state) => [
    state.minute,
    state.timeLabel,
    state.pvMW,
    state.loadMW,
    state.bessPowerMW,
    state.socPercent,
    state.gridImportMW,
    state.wheelingMW,
    state.tariff,
    state.emsMode,
    state.command,
    state.cumulativeGainDh,
  ]);
  downloadText(`ocp_scada_supervision_${simulationDate}.csv`, [header, ...rows].map((row) => row.join(",")).join("\n"), "text/csv");
}

function exportJson(): void {
  if (excelModeActive && latestExcelResult) {
    downloadText(
      `ocp_excel_scada_full_year_${latestExcelResult.minuteStates[0].timestamp.slice(0, 10)}.json`,
      exportExcelSummaryJson(latestExcelResult),
      "application/json",
    );
    return;
  }

  downloadText(
    `ocp_scada_supervision_${simulationDate}.json`,
    JSON.stringify({
      project: "OCP Benguerir PV-BESS SCADA",
      simulationDate,
      pvgis: latestPvgis,
      weather: latestWeather,
      minuteStates,
    }, null, 2),
    "application/json",
  );
}

function exportHourlySummary(): void {
  if (!excelModeActive || !latestExcelResult) {
    setText("excelStatus", "Hourly export is available after loading Excel full-year mode.");
    return;
  }

  downloadText(
    `ocp_excel_scada_hourly_summary_${latestExcelResult.minuteStates[0].timestamp.slice(0, 10)}.csv`,
    exportExcelHourlySummaryCsv(latestExcelResult.hourlySummaries),
    "text/csv",
  );
}

function downloadText(filename: string, content: string, mimeType: string): void {
  const blob = new Blob([content], { type: `${mimeType};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function setBadge(id: string, text: string, statusClass: "ready" | "running" | "paused" | "warning" | "fault" | "online" | "offline"): void {
  const element = getElement<HTMLElement>(id);
  element.textContent = text;
  element.className = `badge ${statusClass}`;
}

function getButton(id: string): HTMLButtonElement {
  return getElement<HTMLButtonElement>(id);
}

function getElement<T extends HTMLElement | SVGPathElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing SCADA element: ${id}`);
  }
  return element as T;
}

function setText(id: string, value: string): void {
  getElement<HTMLElement>(id).textContent = value;
}

function setElementHtml(id: string, value: string): void {
  getElement<HTMLElement>(id).innerHTML = value;
}

function formatClock(value: string): string {
  return value.slice(11, 16);
}

function formatSigned(value: number, decimals: number): string {
  const sign = value > 0 ? "+" : "";
  return `${sign}${formatNumber(value, decimals)}`;
}

function formatNumber(value: number, decimals: number): string {
  if (!Number.isFinite(value)) return "N/A";
  return value.toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
