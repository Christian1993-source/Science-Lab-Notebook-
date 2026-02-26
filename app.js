const STORAGE_KEY = "libretaLaboratorio.draft";
const REPORT_ID_KEY = "libretaLaboratorio.reportId";
const REPORT_STARTED_AT_KEY = "libretaLaboratorio.startedAt";

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
  "references"
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
  { type: "text", key: "references", label: "References (APA 7)" }
];

const scienceTableTemplates = {
  rawData: ["Trial", "", "", "", ""],
  processedData: ["Trial", "", "", "", ""]
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
  startedAt: 0,
  status: "Draft",
  tables: {
    rawData: defaultTableList("rawData"),
    processedData: defaultTableList("processedData")
  },
  isSaving: false,
  pendingSave: false,
  remoteDraftEnabled: true,
  idleTimer: null,
  intervalTimer: null,
  restrictionAlertAt: 0
};

localStorage.setItem(REPORT_ID_KEY, state.reportId);
localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));

const elements = {
  title: document.getElementById("title"),
  teacher: document.getElementById("teacher"),
  studentName: document.getElementById("studentName"),
  date: document.getElementById("date"),
  resetBtn: document.getElementById("resetBtn"),
  loadChemistryExampleBtn: document.getElementById("loadChemistryExampleBtn"),
  loadPhysicsExampleBtn: document.getElementById("loadPhysicsExampleBtn"),
  saveDraftBtn: document.getElementById("saveDraftBtn"),
  submitBtn: document.getElementById("submitBtn"),
  saveState: document.getElementById("saveState"),
  statusBadge: document.getElementById("documentStatus"),
  rawDataEditor: document.getElementById("rawDataEditor"),
  processedDataEditor: document.getElementById("processedDataEditor")
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
  references: document.getElementById("section-references")
};

init();

function init() {
  attachRestrictions();
  attachInputListeners();
  renderTableEditor("rawData", elements.rawDataEditor);
  renderTableEditor("processedData", elements.processedDataEditor);
  updateStatusBadge();
  setFormLocked(state.status === "Submitted");

  state.intervalTimer = setInterval(() => {
    void saveDraft("interval");
  }, 15000);
}

function generateId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `report-${Date.now()}-${Math.random().toString(16).slice(2)}`;
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
  const minColumns = isScienceTable(tableKey) ? getTemplateHeaders(tableKey).length : 2;
  const maxColumns = rows.reduce((max, row) => {
    if (!Array.isArray(row)) {
      return max;
    }
    return Math.max(max, row.length);
  }, 0);

  const width = Math.max(headers.length, maxColumns, minColumns);
  const defaultHeaders = getDefaultHeaders(tableKey, width);
  if (headers.length === 0) {
    headers = defaultHeaders.slice();
  }
  while (headers.length < width) {
    headers.push(defaultHeaders[headers.length] ?? `Column ${headers.length + 1}`);
  }

  if (isScienceTable(tableKey)) {
    headers = headers.map((header, index) => {
      if (index === 0) {
        return "Trial";
      }
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
  const standardInputs = [elements.title, elements.teacher, elements.studentName, elements.date, ...Object.values(sectionInputs)];

  standardInputs.forEach((input) => {
    input.addEventListener("input", () => {
      if (state.status === "Submitted") {
        return;
      }
      if (input === elements.studentName) {
        maybeStartTimerFromStudentName();
      }
      persistLocalBackup();
      queueIdleSave();
    });
  });

  elements.saveDraftBtn.addEventListener("click", () => {
    void saveDraft("manual");
  });

  elements.resetBtn.addEventListener("click", () => {
    resetAllReport();
  });

  elements.loadChemistryExampleBtn.addEventListener("click", () => {
    loadExampleReport("chemistry");
  });

  elements.loadPhysicsExampleBtn.addEventListener("click", () => {
    loadExampleReport("physics");
  });

  elements.submitBtn.addEventListener("click", () => {
    void submitFinalReport();
  });
}

function getChemistryExampleReport() {
  return {
    id: state.reportId,
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

function getPhysicsExampleReport() {
  return {
    id: state.reportId,
    teacherEmail: "",
    teacher: "Mr. Mercado",
    startedAt: 0,
    title: "Kinematic Analysis of Motion on an Inclined Track",
    studentName: "Daniel R. Morales",
    date: "2026-02-12",
    status: "Draft",
    sections: {
      researchQuestion:
        "How does a cart's velocity change with time on an inclined track when the independent variable is elapsed time (s) and the dependent variable is velocity (m/s), and what acceleration with uncertainty is obtained from repeated trials?",
      backgroundInformation:
        "For uniformly accelerated motion, kinematics predicts a linear velocity-time relationship v = v0 + at and a quadratic position-time relationship x = x0 + v0t + 1/2at^2 (Giancoli, 2016). On a fixed incline with limited friction, the component of gravity along the track produces approximately constant acceleration (Serway & Jewett, 2018). Estimating acceleration from the slope of v versus t is a standard approach in introductory mechanics (Young & Freedman, 2020). Reliable conclusions require explicit uncertainty analysis from instrument resolution and repeated trials (Halliday et al., 2018). Reporting mean values with percent uncertainty strengthens scientific validity and supports model-based interpretation (Knight, 2017).",
      variables:
        "Independent variable: elapsed time t. Dependent variable: cart velocity v and derived acceleration a. Controlled variables: incline angle, cart mass, release method, same track surface, and same photogate timing setup.",
      hypothesis:
        "If the incline angle remains fixed and friction is small, then velocity will increase linearly with time and calculated acceleration will be approximately constant across trials.",
      materials:
        "Dynamics cart (0.250 +/- 0.001 kg), aluminum track (1.20 +/- 0.01 m), fixed incline angle 11.0 +/- 0.2 degrees, two photogates with timer resolution +/- 0.001 s, meter stick (1 mm resolution), digital balance (0.01 g resolution), clamp stand, and data sheet.",
      procedure:
        "Set track angle to a constant value and verify alignment. Release the cart from the same start point without push. Record velocity at fixed time intervals for each trial, estimate measurement uncertainty from instrument resolution, repeat for three trials, and compute acceleration from the slope of v versus t in each trial.",
      rawDataNotes:
        "Raw velocity values increased steadily with time in every run. Instrument uncertainties were tracked for both time and velocity, and the spread remained small across trials.",
      processedDataNotes:
        "Linear fits of velocity versus time yielded high coefficients (R^2 > 0.99), supporting near-constant acceleration. Mean acceleration across trials was stable within experimental uncertainty and consistent with the expected incline model.",
      processedDataSampleCalculations:
        "Trial 2 slope method: two points (t = 0.40 +/- 0.001 s, v = 0.83 +/- 0.03 m/s) and (t = 1.20 +/- 0.001 s, v = 2.42 +/- 0.03 m/s). a = Delta v / Delta t = (2.42 - 0.83) / (1.20 - 0.40) = 1.99 m/s^2. Propagated uncertainty: Delta a / a ~= sqrt((Delta(Delta v)/Delta v)^2 + (Delta(Delta t)/Delta t)^2) ~= 0.027, so Delta a ~= 0.05 m/s^2.",
      conclusion:
        "The data support the hypothesis: velocity increased linearly with time and acceleration remained nearly constant across trials. The kinematic model for uniformly accelerated motion describes the system well.",
      evaluation:
        "Main uncertainties came from slight release inconsistencies, photogate alignment, and rolling friction. These effects shifted absolute slope modestly but did not alter linear v-t behavior.",
      improvements:
        "Use a longer track interval for improved slope precision, calibrate sensors before each run, and add five or more trials to reduce uncertainty in mean acceleration.",
      references:
        "Giancoli, D. C. (2016). Physics: Principles with applications (7th ed.). Pearson.\nHalliday, D., Resnick, R., & Walker, J. (2018). Fundamentals of physics (11th ed.). Wiley.\nKnight, R. D. (2017). Physics for scientists and engineers: A strategic approach (4th ed.). Pearson.\nSerway, R. A., & Jewett, J. W. (2018). Physics for scientists and engineers (10th ed.). Cengage.\nYoung, H. D., & Freedman, R. A. (2020). University physics with modern physics (15th ed.). Pearson."
    },
    tables: {
      rawData: {
        title: "Table 1. Kinematics Raw Data with Measurement Uncertainty",
        headers: ["Trial", "Time t (s)", "Delta t (s)", "Velocity v (m/s)", "Delta v (m/s)", "Position x (m)"],
        rows: [
          ["Trial 1", "0.40", "0.001", "0.81", "0.03", "0.16"],
          ["Trial 2", "0.80", "0.001", "1.62", "0.03", "0.64"],
          ["Trial 3", "1.20", "0.001", "2.43", "0.03", "1.44"]
        ]
      },
      processedData: {
        title: "Table 2. Processed Kinematics Results and Uncertainty",
        headers: ["Trial", "Acceleration a (m/s^2)", "Delta a (m/s^2)", "Percent Uncertainty (%)", "R^2", "Percent Difference (%)"],
        rows: [
          ["Trial 1", "1.98", "0.05", "2.53", "0.993", "0.50"],
          ["Trial 2", "1.99", "0.05", "2.51", "0.995", "0.00"],
          ["Trial 3", "2.02", "0.05", "2.48", "0.992", "1.51"],
          ["Mean", "2.00", "0.05", "2.50", "0.993", "0.67"]
        ]
      }
    }
  };
}

function loadExampleReport(exampleType) {
  if (state.status === "Submitted") {
    return;
  }

  const exampleReport = exampleType === "physics" ? getPhysicsExampleReport() : getChemistryExampleReport();

  applyReportToUI(exampleReport);
  persistLocalBackup();
  elements.saveState.textContent =
    exampleType === "physics"
      ? "Physics example loaded with scientific, high-achievement report content."
      : "Chemistry example loaded with scientific, high-achievement report content.";
  queueIdleSave();
}

function resetAllReport() {
  if (!window.confirm("Reset all fields and start a new report?")) {
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
  state.startedAt = 0;
  state.status = "Draft";
  state.tables = {
    rawData: defaultTableList("rawData"),
    processedData: defaultTableList("processedData")
  };

  localStorage.removeItem(STORAGE_KEY);
  localStorage.setItem(REPORT_ID_KEY, state.reportId);
  localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));

  applyReportToUI({
    id: state.reportId,
    teacher: "",
    title: "",
    studentName: "",
    date: "",
    status: "Draft",
    sections: {},
    tables: {
      rawData: defaultTableList("rawData"),
      processedData: defaultTableList("processedData")
    },
    startedAt: state.startedAt
  });

  state.intervalTimer = setInterval(() => {
    void saveDraft("interval");
  }, 15000);

  elements.saveState.textContent = "Report reset. You can start a new draft.";
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

  sectionOrder.forEach((section) => {
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
        if (table.title) {
          drawParagraph(String(table.title), { bold: true, size: 11, lineHeight: 15 });
        }
        if (sectionTables.length > 1) {
          drawParagraph(`Table ${tableIndex + 1}`, { bold: true, size: 11, lineHeight: 15 });
        }
        doc.autoTable({
          startY: y,
          head: [table.headers],
          body: table.rows.filter((row) => row.some((cell) => String(cell || "").trim())),
          theme: "grid",
          styles: {
            font: "times",
            fontSize: 10,
            cellPadding: 4,
            lineColor: [82, 120, 102],
            lineWidth: 0.5
          },
          headStyles: { fillColor: [232, 241, 236], textColor: [20, 52, 39] },
          margin: { left: margin, right: margin }
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
  window.alert("Copy and paste are disabled. Please write your own work.");
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

    const tableElement = document.createElement("table");
    tableElement.className = "table-grid";

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
        const input = document.createElement("input");
        input.type = "text";
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
    tableBlock.append(controls, tableElement);
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
    id: state.reportId,
    teacherEmail: "",
    teacher: elements.teacher.value.trim(),
    title: elements.title.value.trim(),
    studentName: elements.studentName.value.trim(),
    date: elements.date.value,
    startedAt: state.startedAt,
    timeSpentSeconds: getTimeSpentSeconds(),
    status: state.status,
    sections,
    tables: {
      rawData: state.tables.rawData,
      processedData: state.tables.processedData
    }
  };
}

function applyReportToUI(report) {
  const normalizedReport = report && typeof report === "object" ? report : {};

  if (normalizedReport.id) {
    state.reportId = normalizedReport.id;
    localStorage.setItem(REPORT_ID_KEY, normalizedReport.id);
  }

  if (Object.prototype.hasOwnProperty.call(normalizedReport, "startedAt")) {
    const parsedStartedAt = Number(normalizedReport.startedAt);
    state.startedAt = !Number.isNaN(parsedStartedAt) && parsedStartedAt > 0 ? parsedStartedAt : 0;
    localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));
  }

  elements.teacher.value = normalizedReport.teacher || "";
  elements.title.value = normalizedReport.title || "";
  elements.studentName.value = normalizedReport.studentName || "";
  elements.date.value = normalizedReport.date || "";

  sectionKeys.forEach((sectionKey) => {
    sectionInputs[sectionKey].value = normalizedReport.sections?.[sectionKey] || "";
  });

  state.tables.rawData = normalizeTableList(normalizedReport.tables?.rawData, "rawData");
  state.tables.processedData = normalizeTableList(normalizedReport.tables?.processedData, "processedData");
  state.status = normalizedReport.status === "Submitted" ? "Submitted" : "Draft";

  renderTableEditor("rawData", elements.rawDataEditor);
  renderTableEditor("processedData", elements.processedDataEditor);
  updateStatusBadge();
  setFormLocked(state.status === "Submitted");
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
    const response = await fetch(`/api/report/${encodeURIComponent(state.reportId)}`);
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
    return "Draft server unavailable. Saved locally.";
  }
  const looksLikeHtml = /<\/?[a-z][\s\S]*>/i.test(raw) || /^<!doctype html>/i.test(raw);
  if (looksLikeHtml) {
    return "Draft server unavailable. Saved locally.";
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
  if (!report.title || !report.studentName || !report.date) {
    window.alert("Title of Experiment, Student Name, and Date are required.");
    return;
  }

  elements.submitBtn.disabled = true;
  elements.saveState.textContent = "Generating final PDF...";

  try {
    const pdfBlob = generatePdfInBrowser(report);
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
