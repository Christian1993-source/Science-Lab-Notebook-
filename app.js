const STORAGE_KEY = "libretaLaboratorio.draft";
const REPORT_ID_KEY = "libretaLaboratorio.reportId";
const REPORT_STARTED_AT_KEY = "libretaLaboratorio.startedAt";
const PROGRAM_KEY = "libretaLaboratorio.program";
const REPORT_TOKEN_KEY = "libretaLaboratorio.reportToken";
const REPORT_SCHEMA_VERSION = 2;
const REPORT_TIME_ZONE = "America/Puerto_Rico";

const sectionKeys = [
  "researchQuestion",
  "backgroundInformation",
  "variables",
  "hypothesis",
  "materials",
  "procedure",
  "rawDataNotes",
  "processedDataNotes",
  "processedDataSampleCalculations",
  "conclusion",
  "evaluation",
  "improvements",
  "references",
  "dpResearchQuestion",
  "dpBackgroundInformation",
  "dpVariables",
  "dpHypothesis",
  "dpMaterials",
  "dpProcedure",
  "dpRawDataNotes",
  "dpProcessedDataNotes",
  "dpProcessedDataSampleCalculations",
  "dpConclusion",
  "dpEvaluation",
  "dpImprovements",
  "dpReferences"
];

const SELECTABLE_INPUT_TYPES = new Set(["text", "search", "url", "tel", "password", "email", "number"]);
const sectionOrder = [
  { type: "text", key: "researchQuestion", label: "Research Question" },
  { type: "text", key: "backgroundInformation", label: "Background Information" },
  { type: "text", key: "variables", label: "Variables" },
  { type: "text", key: "hypothesis", label: "Hypothesis" },
  { type: "text", key: "materials", label: "Materials" },
  { type: "text", key: "procedure", label: "Procedure" },
  { type: "data", key: "rawData", notesKey: "rawDataNotes", label: "Raw Data" },
  {
    type: "data",
    key: "processedData",
    notesKey: "processedDataNotes",
    sampleCalculationsKey: "processedDataSampleCalculations",
    label: "Processed Data"
  },
  { type: "text", key: "conclusion", label: "Conclusion" },
  { type: "text", key: "evaluation", label: "Evaluation" },
  { type: "text", key: "improvements", label: "Improvements" },
  { type: "text", key: "references", label: "References (APA 7)", program: "myp" },
  { type: "text", key: "dpResearchQuestion", label: "Research Question", program: "dp" },
  { type: "text", key: "dpBackgroundInformation", label: "Background Information", program: "dp" },
  { type: "text", key: "dpVariables", label: "Variables", program: "dp" },
  { type: "text", key: "dpHypothesis", label: "Hypothesis", program: "dp" },
  { type: "text", key: "dpMaterials", label: "Materials", program: "dp" },
  { type: "text", key: "dpProcedure", label: "Procedure", program: "dp" },
  { type: "data", key: "dpRawData", notesKey: "dpRawDataNotes", label: "Raw Data", program: "dp" },
  {
    type: "data",
    key: "dpProcessedData",
    notesKey: "dpProcessedDataNotes",
    sampleCalculationsKey: "dpProcessedDataSampleCalculations",
    label: "Processed Data",
    program: "dp"
  },
  { type: "text", key: "dpConclusion", label: "Conclusion", program: "dp" },
  { type: "text", key: "dpEvaluation", label: "Evaluation", program: "dp" },
  { type: "text", key: "dpImprovements", label: "Improvements", program: "dp" },
  { type: "text", key: "dpReferences", label: "References (APA 7)", program: "dp" }
];

sectionOrder.forEach((section) => {
  if (!section.program) {
    section.program = "myp";
  }
});

const PROGRAM_CONFIGS = {
  myp: {
    name: "MYP",
    fullName: "Middle Years Programme",
    sections: [
      "researchQuestion", "backgroundInformation", "variables", "hypothesis", "materials", "procedure",
      "rawData", "processedData", "conclusion", "evaluation", "improvements", "references"
    ]
  },
  dp: {
    name: "DP",
    fullName: "Diploma Programme",
    sections: [
      "dpResearchQuestion", "dpBackgroundInformation", "dpVariables", "dpHypothesis", "dpMaterials", "dpProcedure",
      "dpRawData", "dpProcessedData", "dpConclusion", "dpEvaluation", "dpImprovements", "dpReferences"
    ]
  }
};

function createDefaultActiveSections() {
  return Object.fromEntries(Object.entries(PROGRAM_CONFIGS).map(([program, config]) => [program, config.sections.slice()]));
}

const scienceTableTemplates = {
  rawData: ["Trial", "", "", "", ""],
  processedData: ["Trial", "", "", "", ""],
  dpRawData: ["Trial", "", "", "", ""],
  dpProcessedData: ["Trial", "", "", "", ""]
};
const legacyScienceHeaderPatterns = [
  /^independent/i,
  /^dependent/i,
  /^unit$/i,
  /^observation/i,
  /^measure/i,
  /^processed/i,
  /^sample/i,
  /^uncertainty/i,
  /^interpretation/i,
  /^column\s+\d+$/i
];

const state = {
  reportId: generateId(),
  reportToken: localStorage.getItem(REPORT_TOKEN_KEY) || generateId(),
  startedAt: 0,
  status: "Draft",
  program: localStorage.getItem(PROGRAM_KEY) === "dp" ? "dp" : "myp",
  classCode: "",
  activeSections: createDefaultActiveSections(),
  blockedAttempts: 0,
  programmaticUpdate: false,
  tables: {
    rawData: defaultTableList("rawData"),
    processedData: defaultTableList("processedData"),
    dpRawData: defaultTableList("dpRawData"),
    dpProcessedData: defaultTableList("dpProcessedData")
  },
  isSaving: false,
  pendingSave: false,
  remoteDraftEnabled: true,
  idleTimer: null,
  intervalTimer: null,
  restrictionAlertAt: 0
};

localStorage.setItem(REPORT_ID_KEY, state.reportId);
localStorage.setItem(REPORT_TOKEN_KEY, state.reportToken);
localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));
localStorage.setItem(PROGRAM_KEY, state.program);

const elements = {
  title: document.getElementById("title"),
  teacher: document.getElementById("teacher"),
  studentName: document.getElementById("studentName"),
  date: document.getElementById("date"),
  time: document.getElementById("time"),
  classCode: document.getElementById("classCode"),
  selectedProgram: document.getElementById("selectedProgram"),
  programBadge: document.getElementById("programBadge"),
  outlineList: document.getElementById("outlineList"),
  removedSections: document.getElementById("removedSections"),
  restoreSectionButtons: document.getElementById("restoreSectionButtons"),
  resetBtn: document.getElementById("resetBtn"),
  saveDraftBtn: document.getElementById("saveDraftBtn"),
  loadSavedDraftBtn: document.getElementById("loadSavedDraftBtn"),
  submitBtn: document.getElementById("submitBtn"),
  saveState: document.getElementById("saveState"),
  statusBadge: document.getElementById("documentStatus"),
  rawDataEditor: document.getElementById("rawDataEditor"),
  processedDataEditor: document.getElementById("processedDataEditor"),
  dpRawDataEditor: document.getElementById("dpRawDataEditor"),
  dpProcessedDataEditor: document.getElementById("dpProcessedDataEditor")
};

const sectionInputs = {
  researchQuestion: document.getElementById("section-researchQuestion"),
  backgroundInformation: document.getElementById("section-backgroundInformation"),
  variables: document.getElementById("section-variables"),
  hypothesis: document.getElementById("section-hypothesis"),
  materials: document.getElementById("section-materials"),
  procedure: document.getElementById("section-procedure"),
  rawDataNotes: document.getElementById("section-rawDataNotes"),
  processedDataNotes: document.getElementById("section-processedDataNotes"),
  processedDataSampleCalculations: document.getElementById("section-processedDataSampleCalculations"),
  conclusion: document.getElementById("section-conclusion"),
  evaluation: document.getElementById("section-evaluation"),
  improvements: document.getElementById("section-improvements"),
  references: document.getElementById("section-references"),
  dpResearchQuestion: document.getElementById("section-dpResearchQuestion"),
  dpBackgroundInformation: document.getElementById("section-dpBackgroundInformation"),
  dpVariables: document.getElementById("section-dpVariables"),
  dpHypothesis: document.getElementById("section-dpHypothesis"),
  dpMaterials: document.getElementById("section-dpMaterials"),
  dpProcedure: document.getElementById("section-dpProcedure"),
  dpRawDataNotes: document.getElementById("section-dpRawDataNotes"),
  dpProcessedDataNotes: document.getElementById("section-dpProcessedDataNotes"),
  dpProcessedDataSampleCalculations: document.getElementById("section-dpProcessedDataSampleCalculations"),
  dpConclusion: document.getElementById("section-dpConclusion"),
  dpEvaluation: document.getElementById("section-dpEvaluation"),
  dpImprovements: document.getElementById("section-dpImprovements"),
  dpReferences: document.getElementById("section-dpReferences")
};

init();

function init() {
  attachRestrictions();
  attachInputListeners();
  renderTableEditor("rawData", elements.rawDataEditor);
  renderTableEditor("processedData", elements.processedDataEditor);
  renderTableEditor("dpRawData", elements.dpRawDataEditor);
  renderTableEditor("dpProcessedData", elements.dpProcessedDataEditor);
  resetAllReport({
    requireConfirmation: false,
    restartTimer: false,
    statusMessage: "New blank report started automatically."
  });
  renderProgramUI();
  updateStatusBadge();
  setFormLocked(state.status === "Submitted");

  state.intervalTimer = setInterval(() => {
    void saveDraft("interval");
  }, 15000);
}

function isLegacyExampleDraft(report) {
  if (!report || typeof report !== "object") {
    return false;
  }
  const title = String(report.title || "").trim().toLowerCase();
  const background = String(report.sections?.backgroundInformation || "").toLowerCase();
  return (title.includes("analysis of motion on an inclined track") || background.includes("linear velocity-time relationship v = v0 + at"));
}

function clearLegacyExampleDraft() {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(REPORT_ID_KEY);
  localStorage.removeItem(REPORT_TOKEN_KEY);
  localStorage.removeItem(REPORT_STARTED_AT_KEY);

  state.reportId = generateId();
  state.reportToken = generateId();
  state.startedAt = 0;
  state.status = "Draft";
  state.classCode = "";
  state.activeSections = createDefaultActiveSections();
  state.blockedAttempts = 0;
  state.tables = {
    rawData: defaultTableList("rawData"),
    processedData: defaultTableList("processedData"),
    dpRawData: defaultTableList("dpRawData"),
    dpProcessedData: defaultTableList("dpProcessedData")
  };

  localStorage.setItem(REPORT_ID_KEY, state.reportId);
  localStorage.setItem(REPORT_TOKEN_KEY, state.reportToken);
  localStorage.setItem(REPORT_STARTED_AT_KEY, "0");
}

function generateId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `report-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function formatAutomaticDateTime(timestamp) {
  const dateParts = new Intl.DateTimeFormat("en-US", {
    timeZone: REPORT_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date(timestamp));
  const part = (type) => dateParts.find((item) => item.type === type)?.value || "";
  const date = `${part("year")}-${part("month")}-${part("day")}`;
  const time = new Intl.DateTimeFormat("en-US", {
    timeZone: REPORT_TIME_ZONE,
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  }).format(new Date(timestamp));
  return { date, time };
}

async function initializeAutomaticDateTime(force = false) {
  if (!force && elements.date.value && elements.time.value) {
    return;
  }

  let timestamp = Date.now();
  try {
    const response = await fetch(`./index.html?clock=${Date.now()}`, { method: "HEAD", cache: "no-store" });
    const serverDate = Date.parse(response.headers.get("date") || "");
    if (Number.isFinite(serverDate)) {
      timestamp = serverDate;
    }
  } catch (_error) {
    // Fall back to the device clock when the platform clock is unavailable.
  }

  const automatic = formatAutomaticDateTime(timestamp);
  elements.date.value = automatic.date;
  elements.time.value = automatic.time;
  persistLocalBackup();
}

function getTemplateHeaders(tableKey) {
  return scienceTableTemplates[tableKey] || [];
}

function isScienceTable(tableKey) {
  return getTemplateHeaders(tableKey).length > 0;
}

function getDefaultHeaders(tableKey, width) {
  const safeWidth = Math.max(1, Number(width) || 1);
  const template = getTemplateHeaders(tableKey);
  if (template.length === 0) {
    return Array.from({ length: safeWidth }, (_, index) => `Column ${index + 1}`);
  }
  return Array.from({ length: safeWidth }, (_, index) => {
    if (index === 0) {
      return "Trial";
    }
    return template[index] ?? "";
  });
}

function buildEmptyRow(columnCount, trialLabel = "") {
  const row = Array(Math.max(1, columnCount)).fill("");
  if (trialLabel) {
    row[0] = trialLabel;
  }
  return row;
}

function buildInitialRows(tableKey, columnCount) {
  if (isScienceTable(tableKey)) {
    return [1, 2, 3].map((trialNumber) => buildEmptyRow(columnCount, `Trial ${trialNumber}`));
  }
  return [buildEmptyRow(columnCount)];
}

function shouldBlankLegacyScienceHeader(headerValue, columnIndex) {
  if (columnIndex === 0) {
    return false;
  }
  const normalized = String(headerValue || "").trim();
  if (!normalized) {
    return false;
  }
  return legacyScienceHeaderPatterns.some((pattern) => pattern.test(normalized));
}

function defaultTable(tableKey = "generic") {
  const minColumns = isScienceTable(tableKey) ? getTemplateHeaders(tableKey).length : 2;
  const headers = getDefaultHeaders(tableKey, minColumns);
  return {
    title: "",
    headers,
    rows: buildInitialRows(tableKey, headers.length)
  };
}

function defaultTableList(tableKey = "generic") {
  return [defaultTable(tableKey)];
}

function normalizeTable(table, tableKey = "generic") {
  if (!table || typeof table !== "object") {
    return defaultTable(tableKey);
  }

  const title = String(table.title ?? "").trim();

  let headers = Array.isArray(table.headers)
    ? table.headers.map((header) => String(header ?? "").trim()).slice(0, 12)
    : [];

  let rows = Array.isArray(table.rows) ? table.rows : [];
  const maxColumns = rows.reduce((max, row) => {
    if (!Array.isArray(row)) {
      return max;
    }
    return Math.max(max, row.length);
  }, 0);

  const width = Math.max(headers.length, maxColumns, 1);
  const defaultHeaders = getDefaultHeaders(tableKey, width);
  if (headers.length === 0) {
    headers = defaultHeaders.slice();
  }
  while (headers.length < width) {
    headers.push(defaultHeaders[headers.length] ?? `Column ${headers.length + 1}`);
  }

  if (isScienceTable(tableKey)) {
    headers = headers.map((header, index) => {
      if (shouldBlankLegacyScienceHeader(header, index)) {
        return "";
      }
      return header;
    });
  }

  rows = rows
    .filter((row) => Array.isArray(row))
    .map((row) => {
      const normalized = Array.from({ length: headers.length }, (_, index) => row[index] ?? "");
      return normalized.map((cell) => String(cell ?? ""));
    });

  if (rows.length === 0) {
    rows = buildInitialRows(tableKey, headers.length);
  }

  return { title, headers, rows };
}

function normalizeTableList(tableValue, tableKey = "generic") {
  if (Array.isArray(tableValue)) {
    const normalized = tableValue.map((table) => normalizeTable(table, tableKey));
    return normalized.length > 0 ? normalized : defaultTableList(tableKey);
  }
  if (tableValue && typeof tableValue === "object") {
    return [normalizeTable(tableValue, tableKey)];
  }
  return defaultTableList(tableKey);
}

function maybeStartTimerFromStudentName() {
  if (state.status === "Submitted") {
    return;
  }
  if (state.startedAt > 0) {
    return;
  }
  if (!elements.studentName.value.trim()) {
    return;
  }

  state.startedAt = Date.now();
  localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));
}

function attachInputListeners() {
  const standardInputs = [elements.title, elements.teacher, elements.studentName, elements.date, elements.classCode, ...Object.values(sectionInputs)];

  standardInputs.forEach((input) => {
    input.addEventListener("input", () => {
      if (state.status === "Submitted") {
        return;
      }
      if (input === elements.studentName) {
        maybeStartTimerFromStudentName();
      }
      if (input === elements.classCode) {
        state.classCode = input.value.trim().toUpperCase();
        input.value = state.classCode;
      }
      persistLocalBackup();
      queueIdleSave();
    });
  });

  document.querySelectorAll(".program-option").forEach((button) => {
    button.addEventListener("click", () => setProgram(button.dataset.program));
  });

  document.addEventListener("click", (event) => {
    const removeButton = event.target.closest('[data-action="remove-section"]');
    if (removeButton) {
      const section = removeButton.closest(".report-section");
      if (section) removeSection(section.dataset.sectionKey);
      return;
    }
    const restoreButton = event.target.closest('[data-action="restore-section"]');
    if (restoreButton) restoreSection(restoreButton.dataset.sectionKey);
  });

  elements.saveDraftBtn.addEventListener("click", () => {
    void saveDraft("manual");
  });

  if (elements.loadSavedDraftBtn) {
    elements.loadSavedDraftBtn.addEventListener("click", () => {
      const localDraft = safeParseLocalDraft();
      if (!localDraft) {
        elements.saveState.textContent = "No local draft found in this browser.";
        return;
      }
      applyReportToUI(localDraft);
      elements.saveState.textContent = `Local draft loaded at ${new Date().toLocaleTimeString()}.`;
    });
  }

  elements.resetBtn.addEventListener("click", () => {
    resetAllReport();
  });

  elements.submitBtn.addEventListener("click", () => {
    void submitFinalReport();
  });
}

function setProgram(program) {
  if (!PROGRAM_CONFIGS[program] || state.status === "Submitted") return;
  state.program = program;
  localStorage.setItem(PROGRAM_KEY, program);
  renderProgramUI();
  persistLocalBackup();
  queueIdleSave();
}

function removeSection(sectionKey) {
  const active = state.activeSections[state.program];
  if (!active.includes(sectionKey)) return;
  if (active.length === 1) {
    window.alert("At least one report section must remain active.");
    return;
  }
  state.activeSections[state.program] = active.filter((key) => key !== sectionKey);
  renderProgramUI();
  persistLocalBackup();
}

function restoreSection(sectionKey) {
  const configured = PROGRAM_CONFIGS[state.program].sections;
  if (!configured.includes(sectionKey)) return;
  const activeSet = new Set([...state.activeSections[state.program], sectionKey]);
  state.activeSections[state.program] = configured.filter((key) => activeSet.has(key));
  renderProgramUI();
  persistLocalBackup();
}

function getSectionLabel(sectionKey) {
  return sectionOrder.find((section) => section.key === sectionKey)?.label || sectionKey;
}

function renderProgramUI() {
  const config = PROGRAM_CONFIGS[state.program];
  const active = state.activeSections[state.program];
  document.body.classList.toggle("program-myp", state.program === "myp");
  document.body.classList.toggle("program-dp", state.program === "dp");
  document.querySelectorAll(".program-option").forEach((button) => {
    const selected = button.dataset.program === state.program;
    button.classList.toggle("is-active", selected);
    button.setAttribute("aria-pressed", String(selected));
  });
  document.querySelectorAll(".report-section").forEach((section) => {
    section.hidden = section.dataset.program !== state.program || !active.includes(section.dataset.sectionKey);
  });
  elements.programBadge.textContent = config.name;
  elements.selectedProgram.value = config.name;
  elements.outlineList.replaceChildren();
  const studentItem = document.createElement("li");
  studentItem.innerHTML = '<a href="#studentInfo">Student Information</a>';
  elements.outlineList.appendChild(studentItem);
  active.forEach((key) => {
    const item = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#sec-${key}`;
    link.textContent = getSectionLabel(key);
    item.appendChild(link);
    elements.outlineList.appendChild(item);
  });
  const removed = config.sections.filter((key) => !active.includes(key));
  elements.removedSections.hidden = removed.length === 0;
  elements.restoreSectionButtons.replaceChildren();
  removed.forEach((key) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "restore-section-btn";
    button.dataset.action = "restore-section";
    button.dataset.sectionKey = key;
    button.textContent = `+ ${getSectionLabel(key)}`;
    elements.restoreSectionButtons.appendChild(button);
  });
}

function getChemistryExampleReport() {
  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    id: state.reportId,
    accessToken: state.reportToken,
    teacherEmail: "",
    teacher: "Mr. Mercado",
    startedAt: 0,
    title: "Experimental Determination of Density for Unknown Liquid and Solid Samples",
    studentName: "Alexandra M. Cruz",
    date: "2026-02-10",
    status: "Draft",
    sections: {
      researchQuestion:
        "How accurately can the density of an unknown liquid and an irregular solid be determined when the independent variable is sample type (liquid vs. solid) and the dependent variable is calculated density (g/mL or g/cm^3)?",
      backgroundInformation:
        "Density is an intensive property defined by rho = m/V and does not depend on sample size when temperature and pressure are controlled (Brown et al., 2018). Accurate density determination requires reliable mass and volume measurements and careful uncertainty handling (Harris, 2020). For liquids, volumetric glassware and consistent meniscus reading reduce systematic error (Skoog et al., 2018). For irregular solids, water displacement provides experimental volume, but trapped air and wet-surface effects can bias results (Zumdahl & Zumdahl, 2017). Repeated trials and uncertainty reporting improve confidence in final density values and allow meaningful comparison with accepted reference data (Atkins et al., 2018).",
      variables:
        "Independent variable: sample type and trial repetition. Dependent variable: calculated density (g/mL for liquid, g/cm^3 for solid). Controlled variables: calibrated balance, constant room temperature, meniscus-reading technique, and identical displacement method for all solid trials.",
      hypothesis:
        "If mass and volume are measured consistently, then calculated densities will cluster tightly by trial and align closely with accepted values for the unknowns.",
      materials:
        "Analytical balance (0.01 g resolution), 50.0 mL graduated cylinder (0.5 mL graduations), 100 mL beaker, 25.0 mL volumetric pipette, distilled water at 23 +/- 1 degrees C, unknown liquid sample (~60 mL), irregular metal sample (~26 g), thin thread (~30 cm), and absorbent paper towels.",
      procedure:
        "Measure mass of each sample using the same calibrated balance. For liquid trials, transfer fixed volume and compute rho = m/V. For solid trials, determine displaced volume by immersion and compute rho = m/V. Repeat each condition for three trials and report mean density, spread, and percent error.",
      rawDataNotes:
        "Raw measurements were stable across trials with no gross outliers. Meniscus readings were taken at eye level and balance zero was rechecked before each trial.",
      processedDataNotes:
        "Computed densities were consistent within each sample type. The liquid showed lower relative spread than the solid, consistent with lower volume-reading uncertainty in direct volumetric measurement.",
      processedDataSampleCalculations:
        "Liquid Trial 2: m = 19.82 g and V = 20.00 mL, so rho = 19.82 / 20.00 = 0.991 g/mL. Solid Trial 3: m = 26.44 g and displaced V = 3.00 cm^3, so rho = 26.44 / 3.00 = 8.81 g/cm^3.",
      conclusion:
        "The hypothesis was supported. Both samples produced tightly clustered density values across repeated trials, and final mean values were close to accepted references within expected laboratory uncertainty.",
      evaluation:
        "Primary uncertainty sources were meniscus reading, trapped bubbles during displacement, and residual water on the solid surface. These factors predominantly affect volume and therefore density.",
      improvements:
        "Use a pycnometer for liquid density, vacuum-degas displacement water to reduce bubbles, and increase to five trials for narrower confidence intervals.",
      references:
        "Atkins, P., de Paula, J., & Keeler, J. (2018). Atkins' physical chemistry (11th ed.). Oxford University Press.\nBrown, T. L., LeMay, H. E., Bursten, B. E., Murphy, C., Woodward, P., & Stoltzfus, M. W. (2018). Chemistry: The central science (14th ed.). Pearson.\nHarris, D. C. (2020). Quantitative chemical analysis (10th ed.). W. H. Freeman and Company.\nSkoog, D. A., Holler, F. J., & Crouch, S. R. (2018). Principles of instrumental analysis (7th ed.). Cengage.\nZumdahl, S. S., & Zumdahl, S. A. (2017). Chemistry (10th ed.). Cengage."
    },
    tables: {
      rawData: {
        title: "Table 1. Raw Measurements for Density Determination",
        headers: ["Trial", "Sample", "Mass (g)", "Volume (mL or cm^3)", "Density (g/mL or g/cm^3)", "Uncertainty"],
        rows: [
          ["Trial 1", "Unknown Liquid", "19.76", "20.00", "0.988", "+/- 0.013 g/mL"],
          ["Trial 2", "Unknown Liquid", "19.82", "20.00", "0.991", "+/- 0.013 g/mL"],
          ["Trial 3", "Unknown Liquid", "19.70", "20.00", "0.985", "+/- 0.013 g/mL"],
          ["Trial 1", "Unknown Solid", "26.35", "3.00", "8.78", "+/- 0.15 g/cm^3"],
          ["Trial 2", "Unknown Solid", "26.41", "3.00", "8.80", "+/- 0.15 g/cm^3"],
          ["Trial 3", "Unknown Solid", "26.44", "3.00", "8.81", "+/- 0.15 g/cm^3"]
        ]
      },
      processedData: {
        title: "Table 2. Processed Density Results and Quality Metrics",
        headers: ["Trial", "Sample", "Mean Density", "Std. Spread", "Percent Uncertainty (%)", "Percent Error (%)"],
        rows: [
          ["Trial 1", "Unknown Liquid", "0.988", "0.003", "1.32", "1.20"],
          ["Trial 2", "Unknown Liquid", "0.991", "0.003", "1.31", "0.90"],
          ["Trial 3", "Unknown Liquid", "0.985", "0.003", "1.33", "1.50"],
          ["Mean", "Unknown Liquid", "0.988", "0.003", "1.32", "1.20"],
          ["Mean", "Unknown Solid", "8.80", "0.02", "1.70", "0.95"]
        ]
      }
    }
  };
}

function resetAllReport({
  requireConfirmation = true,
  restartTimer = true,
  statusMessage = "Report reset. You can start a new draft."
} = {}) {
  if (requireConfirmation && !window.confirm("Are you sure you want to delete all information? This action cannot be undone.")) {
    return;
  }

  if (state.intervalTimer) {
    clearInterval(state.intervalTimer);
    state.intervalTimer = null;
  }
  if (state.idleTimer) {
    clearTimeout(state.idleTimer);
    state.idleTimer = null;
  }

  state.reportId = generateId();
  state.reportToken = generateId();
  state.startedAt = 0;
  state.status = "Draft";
  state.classCode = "";
  state.blockedAttempts = 0;
  state.activeSections = createDefaultActiveSections();
  state.tables = {
    rawData: defaultTableList("rawData"),
    processedData: defaultTableList("processedData"),
    dpRawData: defaultTableList("dpRawData"),
    dpProcessedData: defaultTableList("dpProcessedData")
  };

  localStorage.removeItem(STORAGE_KEY);
  localStorage.setItem(REPORT_ID_KEY, state.reportId);
  localStorage.setItem(REPORT_TOKEN_KEY, state.reportToken);
  localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));

  applyReportToUI({
    id: state.reportId,
    teacher: "",
    classCode: "",
    program: state.program,
    activeSections: state.activeSections,
    blockedAttempts: 0,
    title: "",
    studentName: "",
    date: "",
    status: "Draft",
    sections: {},
    tables: {
      rawData: defaultTableList("rawData"),
      processedData: defaultTableList("processedData"),
      dpRawData: defaultTableList("dpRawData"),
      dpProcessedData: defaultTableList("dpProcessedData")
    },
    startedAt: state.startedAt
  });
  void initializeAutomaticDateTime(true);

  if (restartTimer) {
    state.intervalTimer = setInterval(() => {
      void saveDraft("interval");
    }, 15000);
  }

  elements.saveState.textContent = statusMessage;
}

function tableHasContent(table, tableKey = "generic") {
  if (!table || !Array.isArray(table.headers) || !Array.isArray(table.rows)) {
    return false;
  }

  const hasTitle = String(table.title || "").trim().length > 0;
  const hasCellContent = table.rows.some((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim()));
  const defaultHeaders = getDefaultHeaders(tableKey, table.headers.length);
  const hasCustomHeader = table.headers.some((header, index) => {
    const fallback = defaultHeaders[index] || `Column ${index + 1}`;
    const genericFallback = `Column ${index + 1}`;
    const clean = String(header || "").trim();
    return clean && clean !== fallback && clean !== genericFallback;
  });
  return hasTitle || hasCellContent || hasCustomHeader;
}

function buildPrintableSections(report) {
  const sections = [];
  const program = PROGRAM_CONFIGS[report.program] ? report.program : "myp";
  const active = Array.isArray(report.activeSections?.[program])
    ? report.activeSections[program]
    : PROGRAM_CONFIGS[program].sections;

  sectionOrder.forEach((section) => {
    if (section.program !== program || !active.includes(section.key)) {
      return;
    }
    if (section.type === "text") {
      const text = String(report.sections?.[section.key] || "").trim();
      if (text) {
        sections.push({ type: "text", label: section.label, text });
      }
      return;
    }

    const notes = String(report.sections?.[section.notesKey] || "").trim();
    const sampleCalculations = String(report.sections?.[section.sampleCalculationsKey] || "").trim();
    const tableList = normalizeTableList(report.tables?.[section.key], section.key);
    const contentTables = tableList.filter((table) => tableHasContent(table, section.key));
    if (notes || sampleCalculations || contentTables.length > 0) {
      sections.push({
        type: "data",
        tableKey: section.key,
        label: section.label,
        notes,
        sampleCalculations,
        tables: contentTables
      });
    }
  });

  return sections;
}

function wrapPlainText(text, maxChars = 95) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";

  words.forEach((word) => {
    if ((line + " " + word).trim().length <= maxChars) {
      line = (line + " " + word).trim();
    } else {
      if (line) {
        lines.push(line);
      }
      line = word;
    }
  });

  if (line) {
    lines.push(line);
  }

  return lines;
}

function generateBasicPdfBlob(report) {
  const printableSections = buildPrintableSections(report);
  const lines = [];
  lines.push(report.title || "Lab Report");
  lines.push(`Teacher: ${report.teacher || "Not specified"}`);
  lines.push(`Student: ${report.studentName || ""}`);
  lines.push(`Date: ${report.date || ""}`);
  lines.push(`Time: ${report.time || ""}`);
  lines.push(`Programme: ${PROGRAM_CONFIGS[report.program]?.name || "MYP"}`);
  lines.push(`Class Code: ${report.classCode || ""}`);
  lines.push(`Copy and Paste Attempts: ${report.blockedAttempts || 0}`);
  lines.push(`Time Spent: ${formatDuration(report.timeSpentSeconds || getTimeSpentSeconds())}`);
  lines.push("");

  printableSections.forEach((section, index) => {
    lines.push(`${index + 1}. ${section.label}`);
    if (section.type === "text") {
      lines.push(...wrapPlainText(section.text));
      lines.push("");
      return;
    }

    if (section.notes) {
      lines.push(...wrapPlainText(section.notes));
    }
    if (section.sampleCalculations) {
      lines.push(...wrapPlainText(`Sample Calculations: ${section.sampleCalculations}`));
    }

    const tables = Array.isArray(section.tables) ? section.tables : [];
    tables.forEach((table, tableIndex) => {
      if (table.title) {
        lines.push(String(table.title));
      }
      if (tables.length > 1) {
        lines.push(`Table ${tableIndex + 1}`);
      }
      lines.push(table.headers.join(" | "));
      table.rows
        .filter((row) => row.some((cell) => String(cell || "").trim()))
        .forEach((row) => lines.push(row.map((cell) => String(cell || "").trim()).join(" | ")));
      lines.push("");
    });
    if (tables.length === 0 && (section.notes || section.sampleCalculations)) {
      lines.push("No table rows included.");
    }
    lines.push("");
  });

  const escapePdfText = (value) => String(value).replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
  const contentLines = ["BT", "/F1 11 Tf"];
  let y = 740;
  lines.forEach((line) => {
    if (y < 52) {
      return;
    }
    contentLines.push(`1 0 0 1 72 ${y} Tm (${escapePdfText(line)}) Tj`);
    y -= 14;
  });
  contentLines.push("ET");

  const stream = `${contentLines.join("\n")}\n`;
  const objects = [];
  objects.push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  objects.push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");
  objects.push(
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n"
  );
  objects.push(`4 0 obj\n<< /Length ${stream.length} >>\nstream\n${stream}endstream\nendobj\n`);
  objects.push("5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>\nendobj\n");

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((obj) => {
    offsets.push(pdf.length);
    pdf += obj;
  });
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += "0000000000 65535 f \n";
  for (let i = 1; i < offsets.length; i += 1) {
    pdf += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  return new Blob([pdf], { type: "application/pdf" });
}

function generatePdfInBrowser(report) {
  if (!window.jspdf || typeof window.jspdf.jsPDF !== "function") {
    return generateBasicPdfBlob(report);
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 72;
  const maxTextWidth = pageWidth - margin * 2;
  let y = margin;

  const ensureSpace = (needed) => {
    if (y + needed > pageHeight - margin) {
      doc.addPage();
      y = margin;
    }
  };

  const drawParagraph = (text, { bold = false, size = 12, lineHeight = 16, align = "left" } = {}) => {
    const lines = doc.setFont("times", bold ? "bold" : "normal").setFontSize(size).splitTextToSize(text, maxTextWidth);
    const needed = lines.length * lineHeight + 4;
    ensureSpace(needed);
    doc.text(lines, align === "center" ? pageWidth / 2 : margin, y, { align, baseline: "top" });
    y += needed;
  };

  drawParagraph(report.title || "Lab Report", { bold: true, size: 20, lineHeight: 24, align: "center" });
  drawParagraph(`Teacher: ${report.teacher || "Not specified"}`, {
    size: 12,
    align: "center",
    lineHeight: 16
  });
  drawParagraph(`Student: ${report.studentName || ""}`, { size: 12, align: "center", lineHeight: 16 });
  drawParagraph(`Date: ${report.date || ""}`, { size: 12, align: "center", lineHeight: 16 });
  drawParagraph(`Time: ${report.time || ""}`, { size: 12, align: "center", lineHeight: 16 });
  drawParagraph(`Programme: ${PROGRAM_CONFIGS[report.program]?.name || "MYP"}  |  Class Code: ${report.classCode || ""}`, {
    size: 11,
    align: "center",
    lineHeight: 15
  });
  drawParagraph(`Copy and Paste Attempts: ${report.blockedAttempts || 0}`, {
    size: 10,
    align: "center",
    lineHeight: 14
  });
  drawParagraph(`Time Spent: ${formatDuration(report.timeSpentSeconds || getTimeSpentSeconds())}`, {
    size: 12,
    align: "center",
    lineHeight: 16
  });
  y += 8;

  const printableSections = buildPrintableSections(report);
  printableSections.forEach((section, index) => {
    drawParagraph(`${index + 1}. ${section.label}`, { bold: true, size: 13, lineHeight: 18 });

    if (section.type === "text") {
      drawParagraph(section.text, { size: 12, lineHeight: 17 });
      y += 6;
      return;
    }

    if (section.notes) {
      drawParagraph(section.notes, { size: 12, lineHeight: 17 });
      y += 4;
    }

    if (section.sampleCalculations) {
      drawParagraph("Sample Calculations", { bold: true, size: 12, lineHeight: 16 });
      drawParagraph(section.sampleCalculations, { size: 12, lineHeight: 17 });
      y += 4;
    }

    const sectionTables = Array.isArray(section.tables) ? section.tables : [];
    if (sectionTables.length > 0 && typeof doc.autoTable === "function") {
      sectionTables.forEach((table, tableIndex) => {
        const tableLabelCount = Number(Boolean(table.title)) + Number(sectionTables.length > 1);
        ensureSpace(82 + tableLabelCount * 18);
        if (table.title) {
          drawParagraph(String(table.title), { bold: true, size: 11, lineHeight: 15 });
        }
        if (sectionTables.length > 1) {
          drawParagraph(`Table ${tableIndex + 1}`, { bold: true, size: 11, lineHeight: 15 });
        }
        const columnCount = Math.max(1, table.headers.length);
        const columnWidth = maxTextWidth / columnCount;
        const columnStyles = Object.fromEntries(
          Array.from({ length: columnCount }, (_, columnIndex) => [columnIndex, { cellWidth: columnWidth }])
        );
        const tableFontSize = columnCount <= 5 ? 10 : columnCount <= 7 ? 9 : 8;
        doc.autoTable({
          startY: y,
          head: [table.headers],
          body: table.rows.filter((row) => row.some((cell) => String(cell || "").trim())),
          theme: "grid",
          tableWidth: maxTextWidth,
          showHead: "everyPage",
          pageBreak: "auto",
          rowPageBreak: "avoid",
          styles: {
            font: "times",
            fontSize: tableFontSize,
            cellPadding: { top: 7, right: 5, bottom: 7, left: 5 },
            minCellHeight: 32,
            overflow: "linebreak",
            valign: "middle",
            lineColor: [82, 120, 102],
            lineWidth: 0.5
          },
          headStyles: {
            fillColor: [232, 241, 236],
            textColor: [20, 52, 39],
            fontStyle: "bold",
            minCellHeight: 34
          },
          bodyStyles: { minCellHeight: 38 },
          columnStyles,
          margin: { top: margin, bottom: margin, left: margin, right: margin }
        });
        y = doc.lastAutoTable.finalY + 14;
      });
    } else if (sectionTables.length > 0) {
      drawParagraph("Table data included in saved draft.", { size: 11, lineHeight: 15 });
    } else {
      y += 6;
    }
  });

  return doc.output("blob");
}

function attachRestrictions() {
  const blockEvent = (event) => {
    event.preventDefault();
    event.stopPropagation();
    state.blockedAttempts += 1;
    persistLocalBackup();
    showRestrictionAlert();
  };

  ["paste", "copy", "cut", "drop", "dragstart"].forEach((eventName) => {
    document.addEventListener(eventName, blockEvent, true);
  });

  document.addEventListener(
    "beforeinput",
    (event) => {
      const blockedTypes = new Set([
        "insertFromPaste",
        "insertFromDrop",
        "insertFromYank",
        "deleteByCut",
        "insertFromPasteAsQuotation"
      ]);
      if (blockedTypes.has(event.inputType)) {
        blockEvent(event);
      }
    },
    true
  );

  document.addEventListener(
    "keydown",
    (event) => {
      const key = event.key.toLowerCase();
      const code = String(event.code || "").toLowerCase();
      const withCommandKey = event.ctrlKey || event.metaKey;
      const refreshShortcut = key === "f5" || (withCommandKey && (key === "r" || code === "keyr"));
      if (refreshShortcut) {
        return;
      }
      const blockedShortcuts = withCommandKey && ["c", "v", "x", "insert"].includes(key);
      const shiftInsert = event.shiftKey && key === "insert";
      if (blockedShortcuts || shiftInsert) {
        blockEvent(event);
      }
    },
    true
  );

  document.addEventListener("contextmenu", blockEvent, true);
  document.addEventListener("selectstart", blockEvent, true);

  document.addEventListener("focusin", (event) => {
    const field = event.target;
    if ((field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement) && !field.readOnly) {
      field.dataset.safeTypedValue = field.value;
    }
  }, true);

  document.addEventListener("input", (event) => {
    if (state.programmaticUpdate || event.isComposing) return;
    const field = event.target;
    if (!(field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement) || field.readOnly || field.type === "date") return;
    const previous = field.dataset.safeTypedValue ?? "";
    const current = field.value;
    const insertedCount = Math.max(0, current.length - previous.length);
    const inputType = String(event.inputType || "");
    const prohibitedInput = /insertFromPaste|insertFromDrop|insertFromYank|insertReplacementText/i.test(inputType);
    const unexplainedBulkInsertion = insertedCount > 2;
    if (prohibitedInput || unexplainedBulkInsertion) {
      field.value = previous;
      state.blockedAttempts += 1;
      persistLocalBackup();
      showRestrictionAlert();
      event.stopImmediatePropagation();
      return;
    }
    field.dataset.safeTypedValue = current;
  }, true);

  document.addEventListener(
    "selectionchange",
    () => {
      const active = document.activeElement;
      if (active instanceof HTMLTextAreaElement) {
        if (typeof active.selectionStart === "number" && typeof active.selectionEnd === "number") {
          if (active.selectionStart !== active.selectionEnd) {
            active.setSelectionRange(active.selectionEnd, active.selectionEnd);
          }
        }
        return;
      }

      if (active instanceof HTMLInputElement) {
        const supportsSelectionRange = SELECTABLE_INPUT_TYPES.has(active.type);

        if (supportsSelectionRange) {
          if (typeof active.selectionStart === "number" && typeof active.selectionEnd === "number") {
            if (active.selectionStart !== active.selectionEnd) {
              active.setSelectionRange(active.selectionEnd, active.selectionEnd);
            }
          }
        }
      } else {
        const selection = window.getSelection();
        if (selection && !selection.isCollapsed) {
          selection.removeAllRanges();
        }
      }
    },
    true
  );
}

function showRestrictionAlert() {
  const now = Date.now();
  if (now - state.restrictionAlertAt < 1500) {
    return;
  }
  state.restrictionAlertAt = now;
  window.alert("This protected notebook accepts original typing only.");
}

function isTrialHeader(headerValue) {
  return String(headerValue || "")
    .trim()
    .toLowerCase()
    .startsWith("trial");
}

function parseTrialNumber(cellValue) {
  const match = String(cellValue || "")
    .trim()
    .match(/(\d+)/);
  return match ? Number(match[1]) : 0;
}

function getNextTrialNumber(rows) {
  const maxTrial = rows.reduce((max, row) => {
    if (!Array.isArray(row) || row.length === 0) {
      return max;
    }
    return Math.max(max, parseTrialNumber(row[0]));
  }, 0);
  return Math.max(1, maxTrial + 1);
}

function getNextColumnHeader(headers, tableKey) {
  const nextIndex = headers.length;
  const templateHeaders = getTemplateHeaders(tableKey);
  if (templateHeaders[nextIndex]) {
    return templateHeaders[nextIndex];
  }
  if (isScienceTable(tableKey)) {
    return "";
  }
  return `Column ${nextIndex + 1}`;
}

function renderTableEditor(tableKey, container) {
  const tableList = normalizeTableList(state.tables[tableKey], tableKey);
  state.tables[tableKey] = tableList;
  container.innerHTML = "";

  const setControls = document.createElement("div");
  setControls.className = "table-set-controls";

  const addTableBtn = document.createElement("button");
  addTableBtn.type = "button";
  addTableBtn.textContent = "Add New Table";
  addTableBtn.addEventListener("click", () => {
    tableList.push(defaultTable(tableKey));
    onTableChange();
    renderTableEditor(tableKey, container);
  });
  setControls.appendChild(addTableBtn);
  container.appendChild(setControls);

  tableList.forEach((tableData, tableIndex) => {
    const tableBlock = document.createElement("div");
    tableBlock.className = "table-block";

    const blockHeader = document.createElement("div");
    blockHeader.className = "table-block-header";

    const blockTitle = document.createElement("span");
    blockTitle.className = "table-block-title";
    blockTitle.textContent = `Table ${tableIndex + 1}`;
    blockHeader.appendChild(blockTitle);

    const tableTitleInput = document.createElement("input");
    tableTitleInput.type = "text";
    tableTitleInput.className = "table-title-input";
    tableTitleInput.placeholder = "Table title";
    tableTitleInput.value = tableData.title || "";
    tableTitleInput.addEventListener("input", (event) => {
      tableData.title = event.target.value;
      onTableChange();
    });
    blockHeader.appendChild(tableTitleInput);

    if (tableList.length > 1) {
      const removeTableBtn = document.createElement("button");
      removeTableBtn.type = "button";
      removeTableBtn.className = "table-remove-btn";
      removeTableBtn.textContent = "Remove Table";
      removeTableBtn.addEventListener("click", () => {
        if (tableList.length === 1) {
          tableList[0] = defaultTable(tableKey);
        } else {
          tableList.splice(tableIndex, 1);
        }
        onTableChange();
        renderTableEditor(tableKey, container);
      });
      blockHeader.appendChild(removeTableBtn);
    }

    tableBlock.appendChild(blockHeader);

    const controls = document.createElement("div");
    controls.className = "table-controls";

    if (isScienceTable(tableKey)) {
      const addTrialBtn = document.createElement("button");
      addTrialBtn.type = "button";
      addTrialBtn.textContent = "Add Trial Row";
      addTrialBtn.addEventListener("click", () => {
        const trialLabel = isTrialHeader(tableData.headers[0]) ? `Trial ${getNextTrialNumber(tableData.rows)}` : "";
        tableData.rows.push(buildEmptyRow(tableData.headers.length, trialLabel));
        onTableChange();
        renderTableEditor(tableKey, container);
      });
      controls.append(addTrialBtn);
    }

    const addRowBtn = document.createElement("button");
    addRowBtn.type = "button";
    addRowBtn.textContent = "Add Row";
    addRowBtn.addEventListener("click", () => {
      const trialLabel = isScienceTable(tableKey) && isTrialHeader(tableData.headers[0]) ? `Trial ${getNextTrialNumber(tableData.rows)}` : "";
      tableData.rows.push(buildEmptyRow(tableData.headers.length, trialLabel));
      onTableChange();
      renderTableEditor(tableKey, container);
    });

    const addColumnBtn = document.createElement("button");
    addColumnBtn.type = "button";
    addColumnBtn.textContent = "Add Column";
    addColumnBtn.addEventListener("click", () => {
      tableData.headers.push(getNextColumnHeader(tableData.headers, tableKey));
      tableData.rows = tableData.rows.map((row) => [...row, ""]);
      onTableChange();
      renderTableEditor(tableKey, container);
    });

    const deleteColumnBtn = document.createElement("button");
    deleteColumnBtn.type = "button";
    deleteColumnBtn.textContent = "Delete Last Column";
    deleteColumnBtn.addEventListener("click", () => {
      if (tableData.headers.length <= 1) {
        tableData.headers = isScienceTable(tableKey) ? ["Trial"] : ["Column 1"];
        tableData.rows = tableData.rows.map(() => buildEmptyRow(1, isScienceTable(tableKey) ? "Trial 1" : ""));
      } else {
        tableData.headers.pop();
        tableData.rows = tableData.rows.map((row) => row.slice(0, tableData.headers.length));
      }
      onTableChange();
      renderTableEditor(tableKey, container);
    });

    if (isScienceTable(tableKey)) {
      controls.append(addColumnBtn, deleteColumnBtn);
    } else {
      controls.append(addRowBtn, addColumnBtn, deleteColumnBtn);
    }

    const tableViewport = document.createElement("div");
    tableViewport.className = "table-viewport";

    const tableElement = document.createElement("table");
    tableElement.className = "table-grid";
    tableElement.style.minWidth = `${Math.max(700, tableData.headers.length * 136 + 96)}px`;

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    tableData.headers.forEach((header, columnIndex) => {
      const th = document.createElement("th");
      const headerWrap = document.createElement("div");
      headerWrap.className = "header-input-wrap";

      const input = document.createElement("input");
      input.type = "text";
      input.value = header;
      input.addEventListener("input", (event) => {
        tableData.headers[columnIndex] = event.target.value;
        onTableChange();
      });

      const removeColumnBtn = document.createElement("button");
      removeColumnBtn.type = "button";
      removeColumnBtn.className = "col-delete";
      removeColumnBtn.textContent = "Delete";
      removeColumnBtn.addEventListener("click", () => {
        if (tableData.headers.length <= 1) {
          tableData.headers = isScienceTable(tableKey) ? ["Trial"] : ["Column 1"];
          tableData.rows = tableData.rows.map(() => buildEmptyRow(1, isScienceTable(tableKey) ? "Trial 1" : ""));
        } else {
          tableData.headers.splice(columnIndex, 1);
          tableData.rows = tableData.rows.map((row) => {
            const updatedRow = [...row];
            updatedRow.splice(columnIndex, 1);
            if (updatedRow.length > 0) {
              return updatedRow;
            }
            return buildEmptyRow(1, isScienceTable(tableKey) ? "Trial 1" : "");
          });
        }
        onTableChange();
        renderTableEditor(tableKey, container);
      });

      headerWrap.append(input, removeColumnBtn);
      th.appendChild(headerWrap);
      headerRow.appendChild(th);
    });

    const actionsHeader = document.createElement("th");
    actionsHeader.textContent = "Delete Row";
    headerRow.appendChild(actionsHeader);
    thead.appendChild(headerRow);
    tableElement.appendChild(thead);

    const tbody = document.createElement("tbody");
    tableData.rows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");

      row.forEach((cell, colIndex) => {
        const td = document.createElement("td");
        const input = document.createElement("textarea");
        input.className = "table-cell-input";
        input.rows = 2;
        input.value = cell;
        input.addEventListener("input", (event) => {
          tableData.rows[rowIndex][colIndex] = event.target.value;
          onTableChange();
        });
        td.appendChild(input);
        tr.appendChild(td);
      });

      const actionTd = document.createElement("td");
      const removeRowBtn = document.createElement("button");
      removeRowBtn.type = "button";
      removeRowBtn.className = "row-delete";
      removeRowBtn.textContent = "Delete";
      removeRowBtn.addEventListener("click", () => {
        if (tableData.rows.length === 1) {
          const trialLabel = isScienceTable(tableKey) && isTrialHeader(tableData.headers[0]) ? "Trial 1" : "";
          tableData.rows[0] = buildEmptyRow(tableData.headers.length, trialLabel);
        } else {
          tableData.rows.splice(rowIndex, 1);
        }
        onTableChange();
        renderTableEditor(tableKey, container);
      });
      actionTd.appendChild(removeRowBtn);
      tr.appendChild(actionTd);

      tbody.appendChild(tr);
    });

    tableElement.appendChild(tbody);
    tableViewport.appendChild(tableElement);
    tableBlock.append(controls, tableViewport);
    container.appendChild(tableBlock);
  });

  applyLockedState();
}

function onTableChange() {
  if (state.status === "Submitted") {
    return;
  }
  persistLocalBackup();
  queueIdleSave();
}

function getTimeSpentSeconds() {
  const start = Number(state.startedAt);
  if (!Number.isFinite(start) || start <= 0) {
    return 0;
  }
  return Math.max(0, Math.round((Date.now() - start) / 1000));
}

function formatDuration(seconds) {
  const safeSeconds = Math.max(0, Number(seconds) || 0);
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const secs = safeSeconds % 60;

  if (hours > 0) {
    return `${hours}h ${minutes}m ${secs}s`;
  }
  if (minutes > 0) {
    return `${minutes}m ${secs}s`;
  }
  return `${secs}s`;
}

function collectReport() {
  const sections = {};
  sectionKeys.forEach((sectionKey) => {
    sections[sectionKey] = sectionInputs[sectionKey].value.trim();
  });

  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    id: state.reportId,
    accessToken: state.reportToken,
    teacherEmail: "",
    teacher: elements.teacher.value.trim(),
    classCode: state.classCode || elements.classCode.value.trim().toUpperCase(),
    program: state.program,
    activeSections: state.activeSections,
    blockedAttempts: state.blockedAttempts,
    title: elements.title.value.trim(),
    studentName: elements.studentName.value.trim(),
    date: elements.date.value,
    time: elements.time.value,
    startedAt: state.startedAt,
    timeSpentSeconds: getTimeSpentSeconds(),
    status: state.status,
    sections,
    tables: {
      rawData: state.tables.rawData,
      processedData: state.tables.processedData,
      dpRawData: state.tables.dpRawData,
      dpProcessedData: state.tables.dpProcessedData
    }
  };
}

function applyReportToUI(report) {
  const normalizedReport = report && typeof report === "object" ? report : {};

  if (normalizedReport.id) {
    state.reportId = normalizedReport.id;
    localStorage.setItem(REPORT_ID_KEY, normalizedReport.id);
  }

  if (normalizedReport.accessToken) {
    state.reportToken = String(normalizedReport.accessToken);
    localStorage.setItem(REPORT_TOKEN_KEY, state.reportToken);
  }

  if (Object.prototype.hasOwnProperty.call(normalizedReport, "startedAt")) {
    const parsedStartedAt = Number(normalizedReport.startedAt);
    state.startedAt = !Number.isNaN(parsedStartedAt) && parsedStartedAt > 0 ? parsedStartedAt : 0;
    localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));
  }

  state.programmaticUpdate = true;
  state.program = PROGRAM_CONFIGS[normalizedReport.program] ? normalizedReport.program : state.program;
  if (Object.prototype.hasOwnProperty.call(normalizedReport, "classCode")) {
    state.classCode = String(normalizedReport.classCode || "").trim().toUpperCase();
  }
  state.blockedAttempts = Math.max(0, Number(normalizedReport.blockedAttempts) || 0);
  const defaults = createDefaultActiveSections();
  state.activeSections = Object.fromEntries(Object.keys(PROGRAM_CONFIGS).map((program) => {
    const source = normalizedReport.activeSections?.[program];
    if (program === "dp" && Number(normalizedReport.schemaVersion) < REPORT_SCHEMA_VERSION) {
      return [program, defaults[program]];
    }
    const valid = Array.isArray(source) ? defaults[program].filter((key) => source.includes(key)) : defaults[program];
    return [program, valid.length ? valid : defaults[program]];
  }));
  localStorage.setItem(PROGRAM_KEY, state.program);

  elements.teacher.value = normalizedReport.teacher || "";
  elements.title.value = normalizedReport.title || "";
  elements.studentName.value = normalizedReport.studentName || "";
  elements.date.value = normalizedReport.date || "";
  elements.time.value = normalizedReport.time || "";
  elements.classCode.value = state.classCode;

  sectionKeys.forEach((sectionKey) => {
    sectionInputs[sectionKey].value = normalizedReport.sections?.[sectionKey] || "";
  });

  state.tables.rawData = normalizeTableList(normalizedReport.tables?.rawData, "rawData");
  state.tables.processedData = normalizeTableList(normalizedReport.tables?.processedData, "processedData");
  state.tables.dpRawData = normalizeTableList(normalizedReport.tables?.dpRawData, "dpRawData");
  state.tables.dpProcessedData = normalizeTableList(normalizedReport.tables?.dpProcessedData, "dpProcessedData");
  state.status = normalizedReport.status === "Submitted" ? "Submitted" : "Draft";

  renderTableEditor("rawData", elements.rawDataEditor);
  renderTableEditor("processedData", elements.processedDataEditor);
  renderTableEditor("dpRawData", elements.dpRawDataEditor);
  renderTableEditor("dpProcessedData", elements.dpProcessedDataEditor);
  renderProgramUI();
  updateStatusBadge();
  setFormLocked(state.status === "Submitted");
  state.programmaticUpdate = false;
}

function persistLocalBackup() {
  localStorage.setItem(REPORT_ID_KEY, state.reportId);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(collectReport()));
}

function loadLocalDraft() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return;
  }

  try {
    const report = JSON.parse(raw);
    applyReportToUI(report);
  } catch (_error) {
    elements.saveState.textContent = "Local backup could not be read.";
  }
}

async function syncDraftFromServer() {
  if (!state.reportId) {
    return;
  }

  try {
    const response = await fetch(`/api/report/${encodeURIComponent(state.reportId)}`, {
      headers: { "X-Report-Token": state.reportToken }
    });
    if (!response.ok) {
      return;
    }

    const payload = await response.json();
    if (!payload.report) {
      return;
    }

    const localDraft = safeParseLocalDraft();
    if (!localDraft || payload.report.status === "Submitted") {
      applyReportToUI(payload.report);
      persistLocalBackup();
    }
  } catch (_error) {
    // silent fallback to local mode
  }
}

function safeParseLocalDraft() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (_error) {
    return null;
  }
}

function queueIdleSave() {
  clearTimeout(state.idleTimer);
  state.idleTimer = setTimeout(() => {
    void saveDraft("idle");
  }, 3000);
}

function sanitizeServerErrorMessage(message) {
  const raw = String(message || "").trim();
  if (!raw) {
    return "Cloud draft server unavailable on this host. Saved locally in this browser.";
  }
  const looksLikeHtml = /<\/?[a-z][\s\S]*>/i.test(raw) || /^<!doctype html>/i.test(raw);
  if (looksLikeHtml) {
    return "Cloud draft server unavailable on this host. Saved locally in this browser.";
  }
  return raw;
}

async function saveDraft(trigger) {
  if (state.status === "Submitted") {
    return;
  }

  if (state.isSaving) {
    state.pendingSave = true;
    return;
  }

  persistLocalBackup();
  const report = collectReport();
  state.isSaving = true;

  if (trigger === "manual") {
    elements.saveState.textContent = "Saving draft...";
  }

  if (!state.remoteDraftEnabled) {
    elements.saveState.textContent = `Draft saved locally at ${new Date().toLocaleTimeString()}.`;
    state.isSaving = false;
    if (state.pendingSave) {
      state.pendingSave = false;
      void saveDraft("queued");
    }
    return;
  }

  try {
    const response = await fetch("/api/draft", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ report })
    });

    const payload = await readResponse(response);
    if (!response.ok) {
      const serverError = new Error(payload.error || "Failed to save draft.");
      serverError.isServerResponse = true;
       serverError.status = response.status;
      throw serverError;
    }

    state.status = "Draft";
    updateStatusBadge();
    elements.saveState.textContent = `Draft saved at ${new Date().toLocaleTimeString()}.`;
  } catch (error) {
    if (error && error.isServerResponse) {
      if ([404, 405, 501].includes(Number(error.status))) {
        state.remoteDraftEnabled = false;
        elements.saveState.textContent = `Draft saved locally at ${new Date().toLocaleTimeString()}.`;
      } else {
        elements.saveState.textContent = sanitizeServerErrorMessage(error.message);
      }
    } else {
      elements.saveState.textContent = `Draft saved locally at ${new Date().toLocaleTimeString()}.`;
    }
  } finally {
    state.isSaving = false;
    if (state.pendingSave) {
      state.pendingSave = false;
      void saveDraft("queued");
    }
  }
}

async function submitFinalReport() {
  if (state.status === "Submitted") {
    return;
  }

  const report = collectReport();
  if (!report.classCode) {
    window.alert("Enter the Class Code provided by your teacher before downloading the report.");
    elements.classCode.focus();
    return;
  }
  if (!report.title || !report.studentName || !report.date || !report.time) {
    window.alert("Title of Experiment, Student Name, Date, and Time are required.");
    return;
  }

  elements.submitBtn.disabled = true;
  elements.saveState.textContent = "Generating final PDF...";

  try {
    let pdfBlob;
    const openedAsLocalFile = window.location.protocol === "file:";

    if (openedAsLocalFile) {
      pdfBlob = generatePdfInBrowser(report);
    } else {
      try {
        const response = await fetch("/api/submit", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ report })
        });
        if (!response.ok) {
          if ([404, 405, 501].includes(response.status)) {
            pdfBlob = generatePdfInBrowser(report);
          } else {
            const payload = await readResponse(response);
            throw new Error(payload.error || "The report service could not generate the PDF.");
          }
        } else {
          pdfBlob = await response.blob();
        }
      } catch (error) {
        if (error instanceof TypeError) {
          pdfBlob = generatePdfInBrowser(report);
        } else {
          throw error;
        }
      }
    }

    downloadPdf(pdfBlob, `${safeFileName(report.title)}.pdf`);

    state.status = "Submitted";
    persistLocalBackup();
    updateStatusBadge();
    setFormLocked(true);
    clearInterval(state.intervalTimer);
    state.intervalTimer = null;
    clearTimeout(state.idleTimer);
    state.idleTimer = null;
    elements.saveState.textContent = "Final report downloaded. Editing is now locked.";
    window.alert("Final report downloaded successfully.");
  } catch (error) {
    elements.saveState.textContent = error.message || "Failed to generate final PDF.";
    elements.submitBtn.disabled = false;
  }
}

function safeFileName(value) {
  const clean = String(value || "")
    .replace(/[^a-zA-Z0-9-_ ]+/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
  return clean || "lab-report";
}

function downloadPdf(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function updateStatusBadge() {
  if (!elements.statusBadge) {
    return;
  }
  elements.statusBadge.textContent = state.status;
  elements.statusBadge.classList.toggle("status-draft", state.status === "Draft");
  elements.statusBadge.classList.toggle("status-submitted", state.status === "Submitted");
}

function setFormLocked(locked) {
  const controls = document.querySelectorAll("input, textarea, button");
  controls.forEach((control) => {
    if (control === elements.statusBadge) {
      return;
    }
    if (control === elements.resetBtn) {
      return;
    }
    control.disabled = locked;
  });
  applyLockedState();
}

function applyLockedState() {
  if (state.status !== "Submitted") {
    return;
  }
  const tableButtons = document.querySelectorAll(
    ".table-set-controls button, .table-remove-btn, .table-controls button, .row-delete, .col-delete"
  );
  tableButtons.forEach((button) => {
    button.disabled = true;
  });
}

async function readResponse(response) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    return response.json();
  }
  const text = await response.text();
  if (!text) {
    return { error: "Unknown error." };
  }
  return { error: sanitizeServerErrorMessage(text) };
}
