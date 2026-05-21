import type { OcpCsvStudyResult } from "./core/services/realDataStudyRunner.ts";
import type { ScenarioComparisonRow, ScenarioStudyResult } from "./core/models/study.ts";
import type { DispatchStepResult } from "./core/models/scenario.ts";
import type { MonthlyEnergySummary } from "./core/models/profile.ts";
import type { ChartConfiguration, ChartTypeRegistry } from "chart.js";
import { DEFAULT_MIN_SOC_PERCENT, OCP_LOAD_MW } from "./core/constants/ocpDefaults.ts";
import { runOcpStudyFromCsv } from "./core/services/realDataStudyRunner.ts";
import { readUploadedProfileFile } from "./core/utils/profileFileReader.ts";
import { clampPercent } from "./core/utils/units.ts";
import Chart from "chart.js/auto";

let lastStudyResult: OcpCsvStudyResult | null = null;
let selectedProfileFile: File | null = null;
let selectedProfileCsvText: string | null = null;
let selectedProfileSource: InputSource | null = null;
let activeCharts: Array<Chart<keyof ChartTypeRegistry, unknown, unknown>> = [];
let pageInitialized = false;

type InputSource = "file" | "textarea";

let csvFileInput: HTMLInputElement;
let csvTextInput: HTMLTextAreaElement;
let allowShortProfileInput: HTMLInputElement;
let runButton: HTMLButtonElement;
let fileStatus: HTMLElement;
let fileParseStatus: HTMLElement | null;
let frontendModuleStatus: HTMLElement | null;
let messageBox: HTMLElement;
let emptyState: HTMLElement;
let resultsSection: HTMLElement;
let chartsSection: HTMLElement;
let comparisonSection: HTMLElement;
let profileSummary: HTMLElement;
let profileWarnings: HTMLElement;
let recommendedScenario: HTMLElement;
let kpiCards: HTMLElement;
let scenarioTableBody: HTMLTableSectionElement;
let emsInterpretation: HTMLElement;
let downloadCsvButton: HTMLButtonElement;
let downloadJsonButton: HTMLButtonElement;
let minSocRangeInput: HTMLInputElement;
let minSocValueLabel: HTMLElement;
let minSocCardValue: HTMLElement;

document.addEventListener("DOMContentLoaded", () => {
  initOcpEmsSimulatorPage();
});

if (document.readyState !== "loading") {
  initOcpEmsSimulatorPage();
}

function initOcpEmsSimulatorPage(): void {
  if (pageInitialized) return;
  pageInitialized = true;

  try {
    markFrontendLoaded();
    cacheRequiredElements();
    bindMinSocInput();
    bindProfileFileInput();
    bindRunStudyButton();
    initDownloadButtons();
  } catch (error) {
    showVisibleError(formatError(error));
  }
}

function markFrontendLoaded(message = "Frontend module loaded. Upload handler ready."): void {
  frontendModuleStatus = getOptionalElement<HTMLElement>("frontendModuleStatus");
  if (frontendModuleStatus) {
    frontendModuleStatus.textContent = message;
    frontendModuleStatus.classList.remove("text-red-700", "bg-red-50", "border", "border-red-200", "rounded-xl", "p-3");
  }
}

function cacheRequiredElements(): void {
  csvFileInput = getElement<HTMLInputElement>("csvFileInput");
  csvTextInput = getElement<HTMLTextAreaElement>("csv-text");
  allowShortProfileInput = getElement<HTMLInputElement>("allow-short-profile");
  runButton = getElement<HTMLButtonElement>("run-study-button");
  fileStatus = getElement<HTMLElement>("selectedFileStatus");
  fileParseStatus = getOptionalElement<HTMLElement>("fileParseStatus");
  frontendModuleStatus = getOptionalElement<HTMLElement>("frontendModuleStatus");
  messageBox = getElement<HTMLElement>("message-box");
  emptyState = getElement<HTMLElement>("empty-state");
  resultsSection = getElement<HTMLElement>("results-section");
  chartsSection = getElement<HTMLElement>("charts-section");
  comparisonSection = getElement<HTMLElement>("comparison-section");
  profileSummary = getElement<HTMLElement>("profile-summary");
  profileWarnings = getElement<HTMLElement>("profile-warnings");
  recommendedScenario = getElement<HTMLElement>("recommended-scenario");
  kpiCards = getElement<HTMLElement>("kpi-cards");
  scenarioTableBody = getElement<HTMLTableSectionElement>("scenario-table-body");
  emsInterpretation = getElement<HTMLElement>("ems-interpretation");
  downloadCsvButton = getElement<HTMLButtonElement>("download-csv-button");
  downloadJsonButton = getElement<HTMLButtonElement>("download-json-button");
  minSocRangeInput = getElement<HTMLInputElement>("minSocRange");
  minSocValueLabel = getElement<HTMLElement>("minSocValueLabel");
  minSocCardValue = getElement<HTMLElement>("minSocCardValue");
}

function bindMinSocInput(): void {
  minSocRangeInput.value = String(getSelectedMinSocPercent());
  updateMinSocDisplay();
  minSocRangeInput.addEventListener("input", updateMinSocDisplay);
}

function getSelectedMinSocPercent(): number {
  return clampPercent(Number(minSocRangeInput?.value ?? DEFAULT_MIN_SOC_PERCENT));
}

function updateMinSocDisplay(): void {
  const minSocPercent = getSelectedMinSocPercent();
  minSocRangeInput.value = String(minSocPercent);
  minSocValueLabel.textContent = `Min SoC: ${formatNumber(minSocPercent, 0)}%`;
  minSocCardValue.textContent = `${formatNumber(minSocPercent, 0)}%`;
}

function bindProfileFileInput(): void {
  const fileInput = document.getElementById("csvFileInput") as HTMLInputElement | null;
  const selectedFileStatus = document.getElementById("selectedFileStatus");
  const parseStatus = document.getElementById("fileParseStatus");

  if (!fileInput) {
    showVisibleError("Upload input not found: csvFileInput");
    return;
  }

  markFrontendLoaded("Upload input found and handler attached.");

  fileInput.onchange = () => {
    const file = fileInput.files?.[0] ?? null;

    selectedProfileFile = file;
    selectedProfileCsvText = null;
    selectedProfileSource = file ? "file" : null;

    if (!file) {
      if (selectedFileStatus) selectedFileStatus.textContent = "No file selected.";
      if (parseStatus) parseStatus.textContent = "";
      return;
    }

    const extension = file.name.split(".").pop()?.toLowerCase() || "";
    const detectedFormat = extension === "xlsx" || extension === "xls"
      ? "Excel"
      : extension === "csv"
        ? "CSV"
        : "Unknown";

    if (selectedFileStatus) {
      selectedFileStatus.textContent = `Selected file: ${file.name}`;
    }

    if (parseStatus) {
      parseStatus.innerHTML = `
        <div>Detected format: ${detectedFormat}</div>
        <div>Ready to run simulation.</div>
      `;
    }
  };
}

function bindRunStudyButton(): void {
  runButton.onclick = () => {
    void runStudy();
  };
}

function initDownloadButtons(): void {
  downloadCsvButton.onclick = () => {
    if (lastStudyResult) {
      downloadText("ocp_scenario_comparison.csv", lastStudyResult.scenarioComparisonCsv, "text/csv");
    }
  };

  downloadJsonButton.onclick = () => {
    if (lastStudyResult) {
      downloadText("ocp_study_summary.json", lastStudyResult.studySummaryJson, "application/json");
    }
  };
}

async function runStudy(): Promise<void> {
  setLoading(true);
  hideMessage();

  try {
    const { csvText, source } = await getProfileCsvForRun();
    const finalCsvText = selectedProfileCsvText ?? csvText;
    const finalSource = selectedProfileSource ?? source;
    validateProfileCsvForRun(finalCsvText, selectedProfileFile);

    const result = runOcpStudyFromCsv(finalCsvText, {
      constantLoadMW: OCP_LOAD_MW,
      minSocPercent: getSelectedMinSocPercent(),
      allowShortProfileForTesting: allowShortProfileInput.checked,
    });

    lastStudyResult = result;
    renderResults(result);
    showMessage(`OCP EMS study completed successfully using ${finalSource}.`, "success");
  } catch (error) {
    resultsSection.classList.add("hidden");
    chartsSection.classList.add("hidden");
    comparisonSection.classList.add("hidden");
    emptyState.classList.remove("hidden");
    showMessage(formatError(error), "error");
  } finally {
    setLoading(false);
  }
}

async function getProfileCsvForRun(): Promise<{ csvText: string; source: InputSource }> {
  const uploadedFile = selectedProfileFile ?? csvFileInput.files?.[0] ?? null;

  if (uploadedFile) {
    selectedProfileFile = uploadedFile;
    selectedProfileSource = "file";
    renderFileStatus(`Selected file: ${uploadedFile.name}`, `Using uploaded file: ${uploadedFile.name}`);
    const parsedProfile = await readUploadedProfileFile(uploadedFile, {
      defaultStartDate: "2025-01-01T00:00:00",
      defaultConstantLoadMWh: OCP_LOAD_MW,
      generateConstantLoadIfMissing: true,
      annualRowsUseGeneratedTimestamps: true,
    });
    selectedProfileCsvText = parsedProfile.csvText;

    if (parsedProfile.detectedFormat === "Excel") {
      const parseStatusLines = [
        `Detected format: ${parsedProfile.detectedFormat}`,
        `Excel profile parsed successfully: ${formatNumber(parsedProfile.rowCount, 0)} rows.`,
      ];

      if (parsedProfile.generatedHourlyTimestamps) {
        parseStatusLines.push("Hourly timestamps generated from 2025-01-01T00:00:00.");
      }
      if (parsedProfile.generatedConstantLoad) {
        parseStatusLines.push(`Missing consumption column: constant ${OCP_LOAD_MW} MWh hourly load generated.`);
      }

      renderFileStatus(`Selected file: ${uploadedFile.name}`, parseStatusLines);
    } else {
      renderFileStatus(`Selected file: ${uploadedFile.name}`, [
        "Detected format: CSV",
        "CSV profile loaded successfully.",
        `Rows: ${formatNumber(parsedProfile.rowCount, 0)}.`,
      ]);
    }
    return {
      csvText: parsedProfile.csvText,
      source: "file",
    };
  }

  const pasted = csvTextInput.value.trim();
  if (!pasted) {
    throw new Error("Please upload a CSV/Excel file or paste CSV data.");
  }

  selectedProfileSource = "textarea";
  selectedProfileCsvText = pasted;
  renderFileStatus("No file selected.", "Using pasted CSV data.");

  return {
    csvText: pasted,
    source: "textarea",
  };
}

function validateProfileCsvForRun(csvText: string, uploadedFile: File | null): void {
  const trimmed = csvText.trim();
  if (trimmed.length === 0) {
    throw new Error("Please upload a CSV/Excel file or paste CSV data.");
  }

  const conflictMarkers = ["<".repeat(7), "=".repeat(7), ">".repeat(7)];
  if (conflictMarkers.some((marker) => trimmed.includes(marker))) {
    throw new Error("Profile CSV contains unresolved conflict markers. Remove them before running the study.");
  }

  const lines = trimmed.split(/\r?\n/);
  if (lines.length <= 1) {
    throw new Error("Profile CSV must contain a header and at least one data row.");
  }

  const normalizedHeader = lines[0].trim().toLowerCase();
  if (!normalizedHeader.includes("timestamp") || !normalizedHeader.includes("production_mwh") || !normalizedHeader.includes("consumption_mwh")) {
    throw new Error("Profile CSV header must contain timestamp, production_mwh, and consumption_mwh.");
  }

  if (uploadedFile && getUploadedFileFormat(uploadedFile) === "Excel") {
    const dataRowCount = lines.length - 1;
    if ((dataRowCount === 8760 || dataRowCount === 8784) && lines.length !== dataRowCount + 1) {
      throw new Error("Excel profile normalization failed: final CSV line count does not match the annual row count.");
    }
  }
}

function renderFileStatus(selectedStatus: string, parseStatus: string | string[] = ""): void {
  fileStatus.textContent = selectedStatus;
  if (fileParseStatus) {
    const lines = Array.isArray(parseStatus) ? parseStatus : [parseStatus];
    fileParseStatus.innerHTML = lines
      .filter((line) => line.length > 0)
      .map((line) => `<div>${escapeHtml(line)}</div>`)
      .join("");
  }
}

function getUploadedFileFormat(file: File): "CSV" | "Excel" {
  return file.name.toLowerCase().match(/\.xlsx?$/) ? "Excel" : "CSV";
}

function renderResults(result: OcpCsvStudyResult): void {
  destroyExistingCharts();
  emptyState.classList.add("hidden");
  resultsSection.classList.remove("hidden");
  chartsSection.classList.remove("hidden");
  comparisonSection.classList.remove("hidden");
  const summary = result.normalizedProfileSummary;
  const recommended = result.studyResult.recommendedScenario;

  profileSummary.innerHTML = [
    metricCard("Rows", formatNumber(summary.rowCount, 0), "Hourly records"),
    metricCard("Total PV Production", `${formatNumber(summary.totalPVProductionMWh, 2)} MWh`, "Normalized PV input"),
    metricCard("Total Load", `${formatNumber(summary.totalLoadMWh, 2)} MWh`, "Annual or sample demand"),
    metricCard("Load Mode", summary.generatedConstantLoad ? "Generated" : "Detected", summary.generatedConstantLoad ? `${OCP_LOAD_MW} MWh per hour` : "Loaded from CSV"),
    metricCard("Warnings", formatNumber(summary.warnings.length, 0), summary.warnings.length > 0 ? "Review alert below" : "No warnings"),
  ].join("");

  profileWarnings.innerHTML = summary.warnings.length > 0
    ? `<div class="rounded-xl border border-amber-200 bg-amber-50 text-amber-800 p-4 text-sm font-medium">${summary.warnings.map(escapeHtml).join("<br>")}</div>`
    : "";

  recommendedScenario.innerHTML = recommended
    ? [
      metricCard("Scenario", recommended.scenarioName, "Selected by gain and cycling logic"),
      metricCard("Annual Net Gain", `${formatNumber(recommended.annualNetGainDh, 0)} DH`, "Estimated yearly value"),
      metricCard("Grid Purchase", `${formatNumber(recommended.annualGridPurchaseMWh, 2)} MWh`, "Grid supply after EMS"),
      metricCard("Battery Discharge", `${formatNumber(recommended.annualBatteryDischargeMWh, 2)} MWh`, "Energy served from BESS"),
      metricCard("Wheeling", `${formatNumber(recommended.annualWheelingMWh, 2)} MWh`, "Exported surplus value"),
      metricCard("Curtailment", `${formatNumber(recommended.annualCurtailmentMWh, 2)} MWh`, "Unused PV energy"),
      metricCard("Equivalent Cycles", formatNumber(recommended.bessEquivalentCycles, 2), "Annual cycling estimate"),
      metricCard("Minimum SoC", `${formatNumber(recommended.minSocPercent, 2)}%`, "Lowest simulated state"),
    ].join("")
    : `<div class="rounded-xl border border-red-200 bg-red-50 text-red-700 p-4 font-medium">No recommended scenario was found.</div>`;

  kpiCards.innerHTML = recommended
    ? [
      metricCard("Annual Net Gain", `${formatNumber(recommended.annualNetGainDh, 0)} DH`, "Economic performance"),
      metricCard("Grid Purchase", `${formatNumber(recommended.annualGridPurchaseMWh, 2)} MWh`, "Imported energy"),
      metricCard("Battery Discharge", `${formatNumber(recommended.annualBatteryDischargeMWh, 2)} MWh`, "BESS contribution"),
      metricCard("Wheeling Energy", `${formatNumber(recommended.annualWheelingMWh, 2)} MWh`, "Surplus monetized"),
      metricCard("Curtailment", `${formatNumber(recommended.annualCurtailmentMWh, 2)} MWh`, "Lost PV opportunity"),
      metricCard("Equivalent Cycles", formatNumber(recommended.bessEquivalentCycles, 2), "Battery utilization"),
      metricCard("Minimum SoC", `${formatNumber(recommended.minSocPercent, 2)}%`, "Reserve check"),
      metricCard("Average Cost", `${formatNumber(recommended.averageCostDhPerKWh, 4)} DH/kWh`, "Supplied energy cost"),
    ].join("")
    : "";

  scenarioTableBody.innerHTML = result.studyResult.scenarioComparison.map(renderComparisonRow).join("");
  renderCharts(result);
  renderInterpretation(result);
}

function renderComparisonRow(row: ScenarioComparisonRow): string {
  return `
    <tr>
      <td>${escapeHtml(row.scenarioName)}</td>
      <td>${formatNumber(row.annualGridPurchaseMWh, 2)}</td>
      <td>${formatNumber(row.annualBatteryDischargeMWh, 2)}</td>
      <td>${formatNumber(row.annualWheelingMWh, 2)}</td>
      <td>${formatNumber(row.annualCurtailmentMWh, 2)}</td>
      <td>${formatNumber(row.pvSelfConsumptionRatePercent, 2)}</td>
      <td>${formatNumber(row.pvSelfSufficiencyRatePercent, 2)}</td>
      <td>${formatNumber(row.bessEquivalentCycles, 2)}</td>
      <td>${formatNumber(row.minSocPercent, 2)}</td>
      <td>${formatNumber(row.annualNetGainDh, 0)}</td>
      <td>${formatNumber(row.averageCostDhPerKWh, 4)}</td>
    </tr>
  `;
}

function renderCharts(result: OcpCsvStudyResult): void {
  const recommendedResult = getRecommendedScenarioResult(result);
  if (!recommendedResult) return;

  const monthlyResults = getMonthlyResults(recommendedResult);
  const hourlyResults = getHourlyResults(recommendedResult);
  const comparisonRows = result.studyResult.scenarioComparison;
  const months = monthlyResults.map((row) => formatMonthLabel(row.month));
  const firstWeek = hourlyResults.slice(0, 168);
  const socStep = hourlyResults.length > 2000 ? 24 : Math.max(1, Math.ceil(hourlyResults.length / 365));
  const socSamples = downsampleHourlyResults(hourlyResults, socStep);

  createChart("monthlyEnergyBalanceChart", {
    type: "bar",
    data: {
      labels: months,
      datasets: [
        stackedDataset("PV to Load", monthlyResults.map((row) => row.pvToLoadMWh), "#f59e0b"),
        stackedDataset("PV to Battery", monthlyResults.map((row) => row.pvToBatteryMWh), "#6366f1"),
        stackedDataset("Battery to Load", monthlyResults.map((row) => row.batteryToLoadMWh), "#06b6d4"),
        stackedDataset("Grid to Load", monthlyResults.map((row) => row.gridToLoadMWh), "#334155"),
        stackedDataset("Grid to Battery", monthlyResults.map((row) => row.gridToBatteryMWh), "#8b5cf6"),
        stackedDataset("Wheeling", monthlyResults.map((row) => row.wheelingMWh), "#10b981"),
        stackedDataset("Curtailment", monthlyResults.map((row) => row.curtailedMWh), "#ef4444"),
      ],
    },
    options: baseChartOptions("Energy (MWh)", true),
  });

  createChart("monthlyPvLoadChart", {
    type: "bar",
    data: {
      labels: months,
      datasets: [
        barDataset("PV Production MWh", monthlyResults.map((row) => row.pvProductionMWh), "#f59e0b"),
        barDataset("OCP Load MWh", monthlyResults.map((row) => row.loadMWh), "#10b981"),
      ],
    },
    options: baseChartOptions("Energy (MWh)", false),
  });

  createChart("socProfileChart", {
    type: "line",
    data: {
      labels: socSamples.map((row, index) => index % 7 === 0 ? row.timestamp.slice(5, 10) : ""),
      datasets: [
        lineDataset("SoC %", socSamples.map((row) => row.socPercent), "#059669", 2),
        lineDataset(`Security reserve ${formatNumber(getSelectedMinSocPercent(), 0)}%`, socSamples.map(() => getSelectedMinSocPercent()), "#ef4444", 1, true),
      ],
    },
    options: baseChartOptions("SoC (%)", false),
  });

  createChart("hourlyOperationChart", {
    type: "line",
    data: {
      labels: firstWeek.map((_, index) => `H${index}`),
      datasets: [
        lineDataset("PV Production", firstWeek.map((row) => row.pvProductionMWh), "#f59e0b", 2),
        lineDataset("OCP Load", firstWeek.map((row) => row.loadMWh), "#10b981", 2),
        lineDataset("Grid to Load", firstWeek.map((row) => row.gridToLoadMWh), "#334155", 2),
        lineDataset("Battery to Load", firstWeek.map((row) => row.batteryToLoadMWh), "#06b6d4", 2),
      ],
    },
    options: baseChartOptions("MWh", false),
  });

  createChart("batteryChargeDischargeChart", {
    type: "bar",
    data: {
      labels: firstWeek.map((_, index) => `H${index}`),
      datasets: [
        stackedDataset("PV to Battery", firstWeek.map((row) => row.pvToBatteryMWh), "#6366f1"),
        stackedDataset("Grid to Battery", firstWeek.map((row) => row.gridToBatteryMWh), "#8b5cf6"),
        stackedDataset("Battery to Load", firstWeek.map((row) => -row.batteryToLoadMWh), "#06b6d4"),
      ],
    },
    options: baseChartOptions("MWh", true),
  });

  createChart("scenarioGainChart", {
    type: "bar",
    data: {
      labels: comparisonRows.map(shortScenarioName),
      datasets: [barDataset("Annual Net Gain DH", comparisonRows.map((row) => row.annualNetGainDh), "#059669")],
    },
    options: baseChartOptions("DH", false, true),
  });

  createChart("scenarioGridPurchaseChart", {
    type: "bar",
    data: {
      labels: comparisonRows.map(shortScenarioName),
      datasets: [barDataset("Grid Purchase MWh", comparisonRows.map((row) => row.annualGridPurchaseMWh), "#334155")],
    },
    options: baseChartOptions("MWh", false),
  });

  createChart("scenarioCyclesChart", {
    type: "bar",
    data: {
      labels: comparisonRows.map(shortScenarioName),
      datasets: [barDataset("Equivalent cycles/year", comparisonRows.map((row) => row.bessEquivalentCycles), "#6366f1")],
    },
    options: baseChartOptions("Cycles/year", false),
  });

  createChart("wheelingCurtailmentChart", {
    type: "bar",
    data: {
      labels: comparisonRows.map(shortScenarioName),
      datasets: [
        barDataset("Wheeling MWh", comparisonRows.map((row) => row.annualWheelingMWh), "#10b981"),
        barDataset("Curtailment MWh", comparisonRows.map((row) => row.annualCurtailmentMWh), "#ef4444"),
      ],
    },
    options: baseChartOptions("MWh", false),
  });

  const flowTotals = getHourlyFlowTotals(hourlyResults);
  createChart("energyFlowDistributionChart", {
    type: "doughnut",
    data: {
      labels: ["PV to Load", "Battery to Load", "Grid to Load", "Wheeling", "Curtailment"],
      datasets: [{
        data: [
          flowTotals.pvToLoadMWh,
          flowTotals.batteryToLoadMWh,
          flowTotals.gridToLoadMWh,
          flowTotals.wheelingMWh,
          flowTotals.curtailedMWh,
        ].map((value) => Math.max(0, value)),
        backgroundColor: ["#f59e0b", "#06b6d4", "#334155", "#10b981", "#ef4444"],
        borderWidth: 0,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
        tooltip: { callbacks: { label: (context) => `${context.label}: ${formatNumber(Number(context.raw), 2)} MWh` } },
      },
    },
  });
}

function renderInterpretation(result: OcpCsvStudyResult): void {
  const recommended = result.studyResult.recommendedScenario;
  const recommendedResult = getRecommendedScenarioResult(result);
  if (!recommended || !recommendedResult) {
    emsInterpretation.innerHTML = `<div class="rounded-xl border border-red-200 bg-red-50 text-red-700 p-4 font-medium">No recommended scenario is available for interpretation.</div>`;
    return;
  }

  const annual = recommendedResult.result.annualKpis;
  const curtailmentRate = annual.annualPVProductionMWh > 0 ? (annual.annualCurtailmentMWh / annual.annualPVProductionMWh) * 100 : 0;
  const gridShare = annual.annualLoadMWh > 0 ? (annual.annualGridPurchaseMWh / annual.annualLoadMWh) * 100 : 0;
  const messages = [
    `The recommended scenario is ${recommended.scenarioName} because it provides strong economic gain while respecting the SoC security reserve.`,
    "The BESS reduces expensive-period grid purchases by discharging to the industrial load when stored energy is available.",
    annual.annualWheelingMWh > 0 ? "Wheeling helps valorize surplus PV when direct load supply and battery charging are saturated." : "Wheeling is limited in this result, so most PV value comes from direct consumption and battery dispatch.",
    annual.bessEquivalentCycles <= 365 ? `Equivalent cycles remain within the annual design limit (${formatNumber(annual.bessEquivalentCycles, 1)} <= 365).` : `Equivalent cycles exceed the nominal annual design limit (${formatNumber(annual.bessEquivalentCycles, 1)} > 365).`,
    curtailmentRate <= 5 ? `Curtailment is low at ${formatNumber(curtailmentRate, 2)}% of PV production.` : `Curtailment is ${formatNumber(curtailmentRate, 2)}%; surplus valorization should be improved.`,
    `Grid purchase still supplies ${formatNumber(gridShare, 1)}% of annual load, which is useful for sizing future PV, BESS, or wheeling improvements.`,
  ];
  const badges = [
    warningBadge(annual.minimumSocPercent < getSelectedMinSocPercent(), `Min SoC below ${formatNumber(getSelectedMinSocPercent(), 0)}%`, "SoC security reserve respected"),
    warningBadge(annual.bessEquivalentCycles > 365, "Cycles above 365/year", "Cycles within 365/year"),
    warningBadge(curtailmentRate > 10, "Curtailment above 10%", "Curtailment acceptable"),
    warningBadge(annual.annualNetGainDh <= 0, "Net gain is not positive", "Positive annual net gain"),
    `<span class="inline-flex rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700">Grid purchase ${formatNumber(gridShare, 1)}% of load</span>`,
  ];

  emsInterpretation.innerHTML = `
    <div class="flex flex-wrap gap-2">${badges.join("")}</div>
    <ul class="list-disc pl-5 space-y-2 mt-4">${messages.map((message) => `<li>${escapeHtml(message)}</li>`).join("")}</ul>
  `;
}

function getRecommendedScenarioResult(result: OcpCsvStudyResult): ScenarioStudyResult | undefined {
  const recommended = result.studyResult.recommendedScenario;
  if (!recommended) return undefined;
  return result.studyResult.scenarioResults.find((scenario) => scenario.scenarioId === recommended.scenarioId);
}

function getMonthlyResults(scenario: ScenarioStudyResult): MonthlyEnergySummary[] {
  return scenario.result.monthlyResults;
}

function getHourlyResults(scenario: ScenarioStudyResult): DispatchStepResult[] {
  return scenario.result.hourlyResults;
}

function downsampleHourlyResults(hourlyResults: DispatchStepResult[], step: number): DispatchStepResult[] {
  return hourlyResults.filter((_, index) => index % step === 0);
}

function getHourlyFlowTotals(hourlyResults: DispatchStepResult[]) {
  return hourlyResults.reduce(
    (totals, row) => {
      totals.pvToLoadMWh += row.pvToLoadMWh;
      totals.batteryToLoadMWh += row.batteryToLoadMWh;
      totals.gridToLoadMWh += row.gridToLoadMWh;
      totals.wheelingMWh += row.wheelingMWh;
      totals.curtailedMWh += row.curtailedMWh;
      return totals;
    },
    {
      pvToLoadMWh: 0,
      batteryToLoadMWh: 0,
      gridToLoadMWh: 0,
      wheelingMWh: 0,
      curtailedMWh: 0,
    },
  );
}

function formatMonthLabel(month: string): string {
  const monthIndex = Number(month.slice(5, 7)) - 1;
  return ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][monthIndex] ?? month;
}

function destroyExistingCharts(): void {
  activeCharts.forEach((chart) => chart.destroy());
  activeCharts = [];
}

function createChart(canvasId: string, config: ChartConfiguration<keyof ChartTypeRegistry, unknown, unknown>): void {
  const canvas = getElement<HTMLCanvasElement>(canvasId);
  activeCharts.push(new Chart(canvas, config));
}

function barDataset(label: string, data: number[], color: string) {
  return {
    label,
    data,
    backgroundColor: color,
    borderColor: color,
    borderWidth: 1,
    borderRadius: 5,
  };
}

function stackedDataset(label: string, data: number[], color: string) {
  return {
    ...barDataset(label, data, color),
    stack: "energy",
  };
}

function lineDataset(label: string, data: number[], color: string, width: number, dashed = false) {
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: color,
    borderWidth: width,
    borderDash: dashed ? [6, 6] : undefined,
    pointRadius: 0,
    tension: 0.25,
    type: "line" as const,
  };
}

function baseChartOptions(yTitle: string, stacked: boolean, dhTooltip = false) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { intersect: false, mode: "index" as const },
    plugins: {
      legend: { position: "bottom" as const },
      tooltip: {
        callbacks: {
          label: (context: { dataset: { label?: string }; raw: unknown }) => {
            const value = Number(context.raw);
            const suffix = dhTooltip ? " DH" : yTitle.includes("%") ? "%" : yTitle.includes("Cycles") ? "" : ` ${yTitle}`;
            return `${context.dataset.label ?? ""}: ${dhTooltip ? formatDh(value) : formatNumber(value, 2)}${dhTooltip ? "" : suffix}`;
          },
        },
      },
    },
    scales: {
      x: { stacked, ticks: { maxRotation: 0, autoSkip: true } },
      y: {
        stacked,
        beginAtZero: true,
        title: { display: true, text: yTitle },
        ticks: {
          callback: (value: string | number) => dhTooltip ? formatCompactDh(Number(value)) : formatCompactNumber(Number(value)),
        },
      },
    },
  };
}

function shortScenarioName(row: ScenarioComparisonRow): string {
  return row.scenarioId
    .replace("PV_BESS_", "")
    .replace("BASE_GRID_ONLY", "Grid only")
    .replace("PV_ONLY", "PV only")
    .replace("OPTIMIZED_PLACEHOLDER", "Optimized")
    .replace("GRID_OFFPEAK", "Off-peak")
    .replace("WHEELING", "Wheeling")
    .replace("BASIC", "Basic");
}

function warningBadge(isWarning: boolean, warningText: string, okText: string): string {
  return isWarning
    ? `<span class="inline-flex rounded-full bg-amber-50 px-3 py-1 text-xs font-semibold text-amber-700 ring-1 ring-amber-200">${escapeHtml(warningText)}</span>`
    : `<span class="inline-flex rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700 ring-1 ring-emerald-200">${escapeHtml(okText)}</span>`;
}

function formatDh(value: number): string {
  return `${formatNumber(value, 0)} DH`;
}

function formatCompactDh(value: number): string {
  if (Math.abs(value) >= 1_000_000) return `${formatNumber(value / 1_000_000, 1)}M`;
  if (Math.abs(value) >= 1_000) return `${formatNumber(value / 1_000, 1)}k`;
  return formatNumber(value, 0);
}

function formatCompactNumber(value: number): string {
  if (Math.abs(value) >= 1_000_000) return `${formatNumber(value / 1_000_000, 1)}M`;
  if (Math.abs(value) >= 1_000) return `${formatNumber(value / 1_000, 1)}k`;
  return formatNumber(value, 0);
}

function metricCard(label: string, value: string, description = ""): string {
  return `
    <div class="metric-card">
      <p class="text-xs font-semibold uppercase tracking-wide text-slate-500">${escapeHtml(label)}</p>
      <p class="text-xl font-extrabold text-slate-900 mt-2 break-words">${escapeHtml(value)}</p>
      ${description ? `<p class="text-xs text-slate-500 mt-2">${escapeHtml(description)}</p>` : ""}
    </div>
  `;
}

function showMessage(message: string, type: "success" | "error"): void {
  messageBox.classList.remove("hidden", "border-emerald-200", "bg-emerald-50", "text-emerald-800", "border-red-200", "bg-red-50", "text-red-700");
  if (type === "success") {
    messageBox.classList.add("border-emerald-200", "bg-emerald-50", "text-emerald-800");
  } else {
    messageBox.classList.add("border-red-200", "bg-red-50", "text-red-700");
  }
  messageBox.textContent = message;
}

function showVisibleError(message: string): void {
  const status = frontendModuleStatus ?? getOptionalElement<HTMLElement>("frontendModuleStatus");
  if (status) {
    status.textContent = message;
    status.classList.add("text-red-700", "bg-red-50", "border", "border-red-200", "rounded-xl", "p-3");
    return;
  }

  const fallbackMessageBox = getOptionalElement<HTMLElement>("message-box");
  if (fallbackMessageBox) {
    messageBox = fallbackMessageBox;
    showMessage(message, "error");
  }
}

function hideMessage(): void {
  messageBox.classList.add("hidden");
  messageBox.textContent = "";
}

function setLoading(isLoading: boolean): void {
  runButton.disabled = isLoading;
  runButton.textContent = isLoading ? "Running study..." : "Run OCP EMS Study";
  runButton.classList.toggle("opacity-60", isLoading);
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

function formatError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (message.includes("Missing required profile columns")) {
    return "Missing PV production column. Accepted names include production_mwh, pv_production_mwh, pv_mwh, production, energie_pv, energy_pv, pv, and pvsyst_mwh.";
  }
  if (message.includes("non-numeric")) {
    return "Invalid numeric value found. Check that production and load columns contain only numbers.";
  }
  if (message.includes("negative")) {
    return "Negative production or load values are not allowed.";
  }
  if (message.includes("8760")) {
    return `${message} Enable short sample profiles only for testing, or provide a real 8760-hour annual CSV.`;
  }
  if (message.includes("timestamp")) {
    return `Timestamp issue: ${message}`;
  }

  return `Unable to run OCP EMS study: ${message}`;
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

function getElement<T extends HTMLElement>(id: string, fallbackId?: string): T {
  const element = document.getElementById(id) ?? (fallbackId ? document.getElementById(fallbackId) : null);
  if (!element) {
    throw new Error(`Missing page element: ${id}`);
  }
  return element as T;
}

function getOptionalElement<T extends HTMLElement>(id: string): T | null {
  return document.getElementById(id) as T | null;
}
