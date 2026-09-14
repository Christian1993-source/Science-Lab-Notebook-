const STORAGE_KEY = "libretaLaboratorio.draft";
const REPORT_ID_KEY = "libretaLaboratorio.reportId";
const REPORT_STARTED_AT_KEY = "libretaLaboratorio.startedAt";
const PROGRAM_KEY = "libretaLaboratorio.program";
const REPORT_TOKEN_KEY = "libretaLaboratorio.reportToken";
const REPORT_SCHEMA_VERSION = 4;
const REPORT_TIME_ZONE = "America/Puerto_Rico";

const sectionKeys = [
  "researchQuestion",
  "backgroundInformation",
  "backgroundPurpose",
  "backgroundScience",
  "variables",
  "independentVariable",
  "dependentVariable",
  "controlledVariables",
  "hypothesis",
  "materials",
  "procedure",
  "rawDataNotes",
  "processedDataNotes",
  "processedDataSampleCalculations",
  "conclusion",
  "evaluation",
  "improvements",
  "safetyConsiderations",
  "pilotObservations",
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
  {
    type: "background",
    key: "backgroundInformation",
    label: "Background Information",
    fieldKeys: ["backgroundPurpose", "backgroundScience"]
  },
  {
    type: "variables",
    key: "variables",
    label: "Variables",
    fieldKeys: ["independentVariable", "dependentVariable", "controlledVariables"]
  },
  { type: "text", key: "hypothesis", label: "Hypothesis" },
  { type: "text", key: "materials", label: "Materials" },
  { type: "text", key: "procedure", label: "Procedure" },
  { type: "diagram", key: "experimentalSetup", label: "Experimental Setup / Diagram" },
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
  { type: "text", key: "safetyConsiderations", label: "Safety, Ethical & Environmental Considerations" },
  { type: "text", key: "pilotObservations", label: "Pilot Test / Preliminary Observations (DP)", dpOnly: true },
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
      "experimentalSetup", "rawData", "processedData", "conclusion", "evaluation", "improvements",
      "safetyConsiderations", "pilotObservations", "references"
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
  program: "myp",
  classCode: "",
  activeSections: createDefaultActiveSections(),
  blockedAttempts: 0,
  programmaticUpdate: false,
  figures: [],
  setupDiagram: { dataUrl: "", title: "", description: "" },
  imageUploadPending: false,
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
  experimentalSetupImage: document.getElementById("experimentalSetupImage"),
  experimentalSetupPreview: document.getElementById("experimentalSetupPreview"),
  experimentalSetupTitle: document.getElementById("experimentalSetupTitle"),
  experimentalSetupDescription: document.getElementById("experimentalSetupDescription"),
  experimentalSetupStatus: document.getElementById("experimentalSetupStatus"),
  removeExperimentalSetupImage: document.getElementById("removeExperimentalSetupImage"),
  outlineList: document.getElementById("outlineList"),
  removedSections: document.getElementById("removedSections"),
  restoreSectionButtons: document.getElementById("restoreSectionButtons"),
  resetBtn: document.getElementById("resetBtn"),
  saveDraftBtn: document.getElementById("saveDraftBtn"),
  loadSavedDraftBtn: document.getElementById("loadSavedDraftBtn"),
  loadPhysicsExampleBtn: document.getElementById("loadPhysicsExampleBtn"),
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
  backgroundPurpose: document.getElementById("section-backgroundPurpose"),
  backgroundScience: document.getElementById("section-backgroundScience"),
  variables: document.getElementById("section-variables"),
  independentVariable: document.getElementById("section-independentVariable"),
  dependentVariable: document.getElementById("section-dependentVariable"),
  controlledVariables: document.getElementById("section-controlledVariables"),
  hypothesis: document.getElementById("section-hypothesis"),
  materials: document.getElementById("section-materials"),
  procedure: document.getElementById("section-procedure"),
  rawDataNotes: document.getElementById("section-rawDataNotes"),
  processedDataNotes: document.getElementById("section-processedDataNotes"),
  processedDataSampleCalculations: document.getElementById("section-processedDataSampleCalculations"),
  conclusion: document.getElementById("section-conclusion"),
  evaluation: document.getElementById("section-evaluation"),
  improvements: document.getElementById("section-improvements"),
  safetyConsiderations: document.getElementById("section-safetyConsiderations"),
  pilotObservations: document.getElementById("section-pilotObservations"),
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
  elements.experimentalSetupImage.addEventListener("change", uploadExperimentalSetupImage);
  elements.removeExperimentalSetupImage.addEventListener("click", () => {
    if (state.status === "Submitted" || state.imageUploadPending) return;
    state.setupDiagram = { dataUrl: "", title: "", description: "" };
    renderExperimentalSetup();
    persistLocalBackup();
    queueIdleSave();
    elements.experimentalSetupStatus.textContent = "Diagram image removed.";
  });
  [[elements.experimentalSetupTitle, "title"], [elements.experimentalSetupDescription, "description"]].forEach(([field, key]) => {
    field.addEventListener("input", () => {
      if (state.status === "Submitted") return;
      state.setupDiagram[key] = field.value;
      if (key === "title") elements.experimentalSetupPreview.alt = field.value || "Experimental setup diagram preview";
      persistLocalBackup();
      queueIdleSave();
    });
  });
  document.getElementById("addGraph").addEventListener("click", () => {
    if (state.status === "Submitted" || state.imageUploadPending) return;
    if (state.figures.length >= LabFigures.MAX_COUNT) {
      document.getElementById("graphUploadStatus").textContent = "You can add up to 6 graph images.";
      return;
    }
    state.figures.push({ dataUrl: "", title: "", description: "" });
    renderGraphFigures();
    persistLocalBackup();
  });
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
  updateSessionGate();
}

function isLegacyExampleDraft(report) {
  if (!report || typeof report !== "object") {
    return false;
  }
  const title = String(report.title || "").trim().toLowerCase();
  const background = [
    report.sections?.backgroundInformation,
    report.sections?.backgroundPurpose,
    report.sections?.backgroundScience
  ].join(" ").toLowerCase();
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

async function prepareGraphImage(file) {
  if (!["image/png", "image/jpeg", "image/webp"].includes(file.type) || file.size > 8 * 1024 * 1024) {
    throw new Error("Choose PNG, JPG or WebP images up to 8 MB each.");
  }
  const url = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("The selected image could not be read."));
    reader.readAsDataURL(file);
  });
    const image = new Image();
    image.src = url;
    await image.decode();
    if (!image.naturalWidth || image.naturalWidth * image.naturalHeight > 40000000) {
      throw new Error("This image is too large. Export a smaller graph from your software.");
    }
    const canvas = document.createElement("canvas");
    const scale = Math.min(1, 1800 / Math.max(image.naturalWidth, image.naturalHeight));
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    const context = canvas.getContext("2d");
    context.fillStyle = "white";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    let dataUrl = canvas.toDataURL("image/png");
    for (const quality of [0.9, 0.8, 0.7]) {
      if (dataUrl.length <= LabFigures.MAX_IMAGE_LENGTH) break;
      dataUrl = canvas.toDataURL("image/jpeg", quality);
    }
    return LabFigures.normalize([{ dataUrl, title: "", description: "" }])[0];
}

function normalizeSingleFigure(value) {
  return LabFigures.normalize(value && typeof value === "object" ? [value] : [])[0]
    || { dataUrl: "", title: "", description: "" };
}

async function uploadExperimentalSetupImage(event) {
  const input = event.target;
  if (state.status === "Submitted" || state.imageUploadPending || !input.files?.length) return;
  const reportId = state.reportId;
  state.imageUploadPending = true;
  input.disabled = true;
  elements.experimentalSetupStatus.textContent = "Preparing diagram image…";
  try {
    const prepared = await prepareGraphImage(input.files[0]);
    if (reportId !== state.reportId || state.status === "Submitted") return;
    const candidate = normalizeSingleFigure({
      dataUrl: prepared.dataUrl,
      title: elements.experimentalSetupTitle.value,
      description: elements.experimentalSetupDescription.value
    });
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...collectReport(), setupDiagram: candidate }));
    state.setupDiagram = candidate;
    renderExperimentalSetup();
    queueIdleSave();
    elements.experimentalSetupStatus.textContent = "Diagram image saved with your draft.";
  } catch (error) {
    if (reportId !== state.reportId) return;
    elements.experimentalSetupStatus.textContent = error.name === "QuotaExceededError"
      ? "Browser storage is full. The existing diagram is unchanged. Try a smaller image."
      : error.message || "The diagram image could not be opened.";
  } finally {
    if (reportId !== state.reportId) return;
    input.value = "";
    state.imageUploadPending = false;
    input.disabled = state.status === "Submitted";
  }
}

function renderExperimentalSetup() {
  const diagram = normalizeSingleFigure(state.setupDiagram);
  state.setupDiagram = diagram;
  elements.experimentalSetupTitle.value = diagram.title;
  elements.experimentalSetupDescription.value = diagram.description;
  elements.experimentalSetupTitle.dataset.safeTypedValue = diagram.title;
  elements.experimentalSetupDescription.dataset.safeTypedValue = diagram.description;
  if (diagram.dataUrl) {
    elements.experimentalSetupPreview.src = diagram.dataUrl;
  } else {
    elements.experimentalSetupPreview.removeAttribute("src");
  }
  elements.experimentalSetupPreview.alt = diagram.title || "Experimental setup diagram preview";
  elements.experimentalSetupPreview.hidden = !diagram.dataUrl;
  elements.removeExperimentalSetupImage.disabled = state.status === "Submitted" || !diagram.dataUrl;
}

async function uploadGraphImages(event) {
  const input = event.target;
  if (state.status === "Submitted" || state.imageUploadPending) return;
  const files = Array.from(input.files || []);
  if (!files.length) return;
  const message = document.getElementById("graphUploadStatus");
  const reportId = state.reportId;
  state.imageUploadPending = true;
  input.disabled = true;
  message.textContent = "Preparing graph images…";
  try {
    const index = Number(input.dataset.figureIndex);
    const target = state.figures[index];
    const added = await prepareGraphImage(files[0]);
    if (reportId !== state.reportId || state.status === "Submitted") return;
    if (state.figures[index] !== target) return;
    const candidate = LabFigures.normalize(state.figures.map((figure, position) => position === index ? { ...figure, dataUrl: added.dataUrl } : figure));
    // Check available draft storage before replacing the current figures.
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...collectReport(), figures: candidate }));
    // Keep the existing fields and their listeners while the student is typing.
    target.dataUrl = added.dataUrl;
    const preview = input.parentElement.querySelector("img");
    if (preview) {
      preview.src = added.dataUrl;
      preview.alt = target.title || `Graph preview ${index + 1}`;
      preview.hidden = false;
    }
    queueIdleSave();
    message.textContent = "Graph image saved with your draft.";
  } catch (error) {
    if (reportId !== state.reportId) return;
    message.textContent = error.name === "QuotaExceededError"
      ? "Browser storage is full. Your existing figures are unchanged. Try a smaller image."
      : error.message || "The image could not be opened. Try exporting it again.";
  } finally {
    if (reportId !== state.reportId) return;
    input.value = "";
    state.imageUploadPending = false;
    input.disabled = state.status === "Submitted";
  }
}

function renderGraphFigures() {
  const container = document.getElementById("graphFigures");
  container.replaceChildren();
  if (!state.figures.length && state.status !== "Submitted") state.figures.push({ dataUrl: "", title: "", description: "" });
  state.figures.forEach((figure, index) => {
    const card = document.createElement("div");
    card.className = "graph-figure";
    const heading = document.createElement("h4");
    heading.textContent = `Figure ${index + 1}`;
    const preview = document.createElement("img");
    if (figure.dataUrl) preview.src = figure.dataUrl;
    preview.hidden = !figure.dataUrl;
    preview.alt = figure.title || `Graph preview ${index + 1}`;
    card.append(heading);
    for (const [key, labelText, tag, limit] of [["title", "Graph title", "input", 160], ["description", "Image description", "textarea", 2000]]) {
      const label = document.createElement("label");
      const field = document.createElement(tag);
      field.id = `graph-${index}-${key}`;
      label.htmlFor = field.id;
      label.textContent = `${labelText} — Figure ${index + 1}`;
      field.value = figure[key];
      field.maxLength = limit;
      field.dataset.safeTypedValue = field.value;
      field.disabled = state.status === "Submitted";
      field.addEventListener("input", () => {
        if (state.status === "Submitted") return;
        figure[key] = field.value;
        if (key === "title") preview.alt = field.value || `Graph preview ${index + 1}`;
        persistLocalBackup();
        queueIdleSave();
      });
      card.append(label, field);
      if (key === "title") {
        const uploadLabel = document.createElement("label");
        const upload = document.createElement("input");
        upload.id = `graph-${index}-image`;
        uploadLabel.htmlFor = upload.id;
        uploadLabel.textContent = `Upload image — Figure ${index + 1}`;
        upload.type = "file";
        upload.accept = "image/png,image/jpeg,image/webp";
        upload.dataset.figureIndex = String(index);
        upload.disabled = state.status === "Submitted";
        upload.addEventListener("change", uploadGraphImages);
        card.append(uploadLabel, upload, preview);
      }
    }
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn-secondary";
    remove.textContent = `Remove Figure ${index + 1}`;
    remove.disabled = state.status === "Submitted";
    remove.addEventListener("click", () => {
      if (state.status === "Submitted" || state.imageUploadPending) return;
      state.figures.splice(index, 1);
      renderGraphFigures();
      persistLocalBackup();
      queueIdleSave();
    });
    card.append(remove);
    container.append(card);
  });
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

function updateSessionGate() {
  const ready = Boolean(elements.studentName.value.trim() && elements.title.value.trim());
  document.getElementById("sessionGateNotice").hidden = ready;
  document.getElementById("reportFields").disabled = !ready || state.status === "Submitted";
  elements.submitBtn.disabled = !ready || state.status === "Submitted";
  document.querySelectorAll("#restoreSectionButtons button").forEach(button => { button.disabled = !ready; });
}

function maybeStartTimerFromStudentName() {
  if (state.status === "Submitted") {
    return;
  }
  if (state.startedAt > 0) {
    return;
  }
  if (!elements.studentName.value.trim() || !elements.title.value.trim()) {
    return;
  }

  state.startedAt = Date.now();
  localStorage.setItem(REPORT_STARTED_AT_KEY, String(state.startedAt));
}

function attachInputListeners() {
  elements.selectedProgram.addEventListener("change", () => {
    renderProgramUI();
    persistLocalBackup();
    queueIdleSave();
  });
  [sectionInputs.materials, sectionInputs.dpMaterials].forEach(field => {
    const update = (value, caret = value.length) => {
      field.value = value;
      field.dataset.safeTypedValue = value;
      field.setSelectionRange(caret, caret);
      persistLocalBackup();
      queueIdleSave();
    };
    field.addEventListener("focus", () => {
      if (state.status === "Submitted") return;
      if (!field.value.trim()) update("1. ");
    });
    field.addEventListener("keydown", event => {
      if (event.key !== "Enter" || event.isComposing || state.status === "Submitted") return;
      event.preventDefault();
      const start = field.selectionStart;
      const end = field.selectionEnd;
      const before = field.value.slice(0, start);
      const currentLine = before.slice(before.lastIndexOf("\n") + 1);
      if (!currentLine.replace(/^\s*\d+[.)]\s*/, "").trim()) return;
      const nextLine = before.split("\n").length;
      const lines = (before + "\n" + field.value.slice(end)).split("\n");
      const numbered = lines.map((line, index) => `${index + 1}. ${line.replace(/^\s*(?:\d+[.)]|[-*•])\s*/, "")}`);
      const caret = numbered.slice(0, nextLine).join("\n").length + 1 + `${nextLine + 1}. `.length;
      update(numbered.join("\n"), caret);
    });
    field.addEventListener("blur", () => {
      if (state.status !== "Submitted") update(LabFigures.numberedMaterials(field.value));
    });
  });
  const standardInputs = [elements.title, elements.teacher, elements.studentName, elements.date, elements.classCode, ...Object.values(sectionInputs)];

  standardInputs.forEach((input) => {
    input.addEventListener("input", () => {
      if (state.status === "Submitted") {
        return;
      }
      if (input === elements.studentName || input === elements.title) {
        maybeStartTimerFromStudentName();
        updateSessionGate();
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

  elements.loadPhysicsExampleBtn.addEventListener("click", () => {
    if (!window.confirm("Load the complete physics example? This will replace all information currently entered in the notebook.")) {
      return;
    }
    applyReportToUI(getPhysicsExampleReport());
    persistLocalBackup();
    elements.saveState.textContent = "Complete physics example loaded. You can review, edit, or download it.";
  });

  elements.resetBtn.addEventListener("click", () => {
    resetAllReport();
  });

  elements.submitBtn.addEventListener("click", () => {
    void submitFinalReport();
  });
}

function setProgram(program) {
  program = "myp";
  if (!PROGRAM_CONFIGS[program] || state.status === "Submitted") return;
  state.program = program;
  localStorage.setItem(PROGRAM_KEY, program);
  renderProgramUI();
  persistLocalBackup();
  queueIdleSave();
}

function removeSection(sectionKey) {
  if (state.status === "Submitted") return;
  const active = state.activeSections[state.program];
  if (!active.includes(sectionKey)) return;
  if (active.length === 1) {
    window.alert("At least one report section must remain active.");
    return;
  }
  state.activeSections[state.program] = active.filter((key) => key !== sectionKey);
  renderProgramUI();
  persistLocalBackup();
  document.getElementById(`removed-${sectionKey}`)?.querySelector("button")?.focus();
}

function restoreSection(sectionKey) {
  if (state.status === "Submitted") return;
  const configured = PROGRAM_CONFIGS[state.program].sections;
  if (!configured.includes(sectionKey)) return;
  const activeSet = new Set([...state.activeSections[state.program], sectionKey]);
  state.activeSections[state.program] = configured.filter((key) => activeSet.has(key));
  renderProgramUI();
  persistLocalBackup();
  const restored = document.getElementById(`sec-${sectionKey}`);
  restored?.querySelector("textarea, input, button")?.focus();
}

function getSectionLabel(sectionKey) {
  return sectionOrder.find((section) => section.key === sectionKey)?.label || sectionKey;
}

function toRoman(number) {
  const values = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let remaining = Math.max(1, Math.floor(number));
  let result = "";
  values.forEach(([value, numeral]) => {
    while (remaining >= value) {
      result += numeral;
      remaining -= value;
    }
  });
  return result;
}

function renderProgramUI() {
  const config = PROGRAM_CONFIGS[state.program];
  const active = state.activeSections[state.program];
  const isDp = elements.selectedProgram.value === "DP";
  const visibleSections = config.sections.filter((key) => {
    const definition = sectionOrder.find((section) => section.key === key && section.program === state.program);
    return !definition?.dpOnly || isDp;
  });
  document.body.classList.toggle("program-myp", state.program === "myp");
  document.body.classList.toggle("program-dp", state.program === "dp");
  document.querySelectorAll(".program-option").forEach((button) => {
    const selected = button.dataset.program === state.program;
    button.classList.toggle("is-active", selected);
    button.setAttribute("aria-pressed", String(selected));
  });
  document.querySelectorAll(".report-section").forEach((section) => {
    const key = section.dataset.sectionKey;
    const belongsToProgram = section.dataset.program === state.program;
    const availableForStudent = visibleSections.includes(key);
    section.hidden = !belongsToProgram || !availableForStudent || !active.includes(key);
    if (belongsToProgram && availableForStudent) {
      const heading = section.querySelector(".section-heading h2");
      if (heading) heading.textContent = `${toRoman(visibleSections.indexOf(key) + 1)}. ${getSectionLabel(key)}`;
    }
    let placeholder = document.getElementById(`removed-${key}`);
    if (!placeholder) {
      placeholder = document.createElement("div");
      placeholder.id = `removed-${key}`;
      placeholder.className = "removed-part-placeholder";
      const copy = document.createElement("div");
      const title = document.createElement("strong");
      title.textContent = `${getSectionLabel(key)} — removed from report`;
      const help = document.createElement("p");
      help.textContent = "Your work is kept while this notebook is open. Restore this part anytime.";
      copy.append(title, help);
      const button = document.createElement("button");
      button.type = "button";
      button.className = "restore-section-btn";
      button.dataset.action = "restore-section";
      button.dataset.sectionKey = key;
      button.textContent = "Restore part";
      button.setAttribute("aria-label", `Restore ${getSectionLabel(key)}`);
      placeholder.append(copy, button);
      section.after(placeholder);
    }
    placeholder.hidden = !belongsToProgram || !availableForStudent || active.includes(key);
  });
  elements.outlineList.replaceChildren();
  const studentItem = document.createElement("li");
  studentItem.innerHTML = '<a href="#studentInfo">Student Information</a>';
  elements.outlineList.appendChild(studentItem);
  visibleSections.filter((key) => active.includes(key)).forEach((key) => {
    const item = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#sec-${key}`;
    link.textContent = getSectionLabel(key);
    item.appendChild(link);
    elements.outlineList.appendChild(item);
  });
  const removed = visibleSections.filter((key) => !active.includes(key));
  elements.removedSections.hidden = removed.length === 0;
  elements.restoreSectionButtons.replaceChildren();
  removed.forEach((key) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "restore-section-btn";
    button.dataset.action = "restore-section";
    button.dataset.sectionKey = key;
    button.textContent = `+ Restore ${getSectionLabel(key)}`;
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
      backgroundPurpose:
        "The purpose of this investigation is to determine the density of unknown liquid and solid samples and compare the experimental values with accepted reference data. Repeated trials and uncertainty reporting improve confidence in the calculated results (Atkins et al., 2018).",
      backgroundScience:
        "Density is an intensive property defined by rho = m/V and does not depend on sample size when temperature and pressure are controlled (Brown et al., 2018). Accurate density determination requires reliable mass and volume measurements (Harris, 2020). For liquids, consistent meniscus readings reduce systematic error, while water displacement can determine the volume of irregular solids (Skoog et al., 2018; Zumdahl & Zumdahl, 2017).",
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

function createPhysicsExampleImages() {
  const makeCanvas = (width, height, draw) => {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    draw(context, width, height);
    return canvas.toDataURL("image/png");
  };

  const diagram = makeCanvas(900, 560, (context) => {
    context.strokeStyle = "#143b2f";
    context.fillStyle = "#143b2f";
    context.lineWidth = 8;
    context.lineCap = "round";
    context.beginPath();
    context.moveTo(170, 485);
    context.lineTo(730, 485);
    context.moveTo(260, 485);
    context.lineTo(260, 90);
    context.moveTo(210, 90);
    context.lineTo(650, 90);
    context.stroke();

    const pivotX = 520;
    const pivotY = 90;
    const bobX = 625;
    const bobY = 390;
    context.fillStyle = "#0f8f8c";
    context.beginPath();
    context.arc(pivotX, pivotY, 13, 0, Math.PI * 2);
    context.fill();
    context.strokeStyle = "#1f2937";
    context.lineWidth = 5;
    context.beginPath();
    context.moveTo(pivotX, pivotY + 10);
    context.lineTo(bobX, bobY - 24);
    context.stroke();
    context.fillStyle = "#b53b3b";
    context.beginPath();
    context.arc(bobX, bobY, 30, 0, Math.PI * 2);
    context.fill();

    context.setLineDash([10, 10]);
    context.strokeStyle = "#7b8790";
    context.lineWidth = 3;
    context.beginPath();
    context.moveTo(pivotX, pivotY + 15);
    context.lineTo(pivotX, 430);
    context.stroke();
    context.setLineDash([]);

    context.strokeStyle = "#0f8f8c";
    context.lineWidth = 5;
    context.strokeRect(690, 325, 70, 100);
    context.beginPath();
    context.moveTo(725, 425);
    context.lineTo(725, 485);
    context.stroke();

    context.fillStyle = "#143b2f";
    context.font = "bold 28px Arial";
    context.fillText("Clamp stand", 75, 70);
    context.fillText("Pivot", 545, 78);
    context.fillText("String length, L", 555, 240);
    context.fillText("50 g bob", 650, 390);
    context.fillText("Photogate", 680, 310);
    context.font = "24px Arial";
    context.fillText("Release angle = 10°", 300, 455);
  });

  const graph = makeCanvas(900, 560, (context, width, height) => {
    const left = 110;
    const right = width - 60;
    const top = 55;
    const bottom = height - 90;
    const plotWidth = right - left;
    const plotHeight = bottom - top;
    const points = [[0.2, 0.81], [0.4, 1.61], [0.6, 2.42], [0.8, 3.22], [1.0, 4.03]];
    const xPosition = (value) => left + value * plotWidth;
    const yPosition = (value) => bottom - (value / 4.2) * plotHeight;

    context.strokeStyle = "#d5ded9";
    context.lineWidth = 2;
    context.font = "20px Arial";
    context.fillStyle = "#374151";
    context.textAlign = "center";
    for (let index = 0; index <= 5; index += 1) {
      const xValue = index * 0.2;
      const x = xPosition(xValue);
      context.beginPath();
      context.moveTo(x, top);
      context.lineTo(x, bottom);
      context.stroke();
      context.fillText(xValue.toFixed(1), x, bottom + 30);
    }
    context.textAlign = "right";
    for (let index = 0; index <= 4; index += 1) {
      const y = yPosition(index);
      context.beginPath();
      context.moveTo(left, y);
      context.lineTo(right, y);
      context.stroke();
      context.fillText(String(index), left - 18, y + 7);
    }

    context.strokeStyle = "#143b2f";
    context.lineWidth = 5;
    context.beginPath();
    context.moveTo(left, top);
    context.lineTo(left, bottom);
    context.lineTo(right, bottom);
    context.stroke();

    context.strokeStyle = "#b53b3b";
    context.lineWidth = 5;
    context.beginPath();
    context.moveTo(xPosition(0), yPosition(0));
    context.lineTo(xPosition(1), yPosition(4.03));
    context.stroke();

    context.fillStyle = "#0f8f8c";
    points.forEach(([xValue, yValue]) => {
      context.beginPath();
      context.arc(xPosition(xValue), yPosition(yValue), 10, 0, Math.PI * 2);
      context.fill();
    });

    context.fillStyle = "#143b2f";
    context.font = "bold 25px Arial";
    context.textAlign = "center";
    context.fillText("Pendulum length, L (m)", (left + right) / 2, height - 30);
    context.save();
    context.translate(35, (top + bottom) / 2);
    context.rotate(-Math.PI / 2);
    context.fillText("Period squared, T² (s²)", 0, 0);
    context.restore();
    context.font = "22px Arial";
    context.textAlign = "left";
    context.fillStyle = "#b53b3b";
    context.fillText("Best-fit line: T² = 4.03L", 500, 95);
    context.fillText("R² = 0.999", 500, 125);
  });

  return { diagram, graph };
}

function getPhysicsExampleReport() {
  const images = createPhysicsExampleImages();
  const studentProgramme = elements.selectedProgram.value === "DP" ? "DP" : "MYP";
  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    id: generateId(),
    accessToken: generateId(),
    teacherEmail: "",
    teacher: "Physics Department",
    classCode: "EXAMPLE",
    program: "myp",
    studentProgramme,
    activeSections: createDefaultActiveSections(),
    blockedAttempts: 0,
    startedAt: Date.now(),
    title: "Investigating the Relationship Between Pendulum Length and Period",
    studentName: "Sample Student",
    date: elements.date.value,
    time: elements.time.value,
    status: "Draft",
    sections: {
      researchQuestion:
        "How does changing the length of a simple pendulum from 0.20 m to 1.00 m affect its period, while the bob mass, release angle, measurement method, and testing location are controlled?",
      backgroundPurpose:
        "The purpose of this investigation is to determine how pendulum length affects period and to use the relationship between T² and L to calculate an experimental value of gravitational field strength. Timing several oscillations reduces the percentage effect of reaction time on a single period (Ling et al., 2016).",
      backgroundScience:
        "The independent variable is pendulum length, L, and the dependent variable is period, T. For small release angles, a simple pendulum follows T = 2π√(L/g), where g is gravitational field strength. Squaring the equation gives T² = (4π²/g)L, predicting a direct linear relationship between T² and L (Serway & Jewett, 2018).",
      independentVariable:
        "Pendulum length, L (m), measured from the pivot to the center of the bob and changed through 0.20, 0.40, 0.60, 0.80, and 1.00 m.",
      dependentVariable:
        "Period, T (s), calculated by timing 10 complete oscillations and dividing the measured time by 10. T² (s²) is used for graphical analysis.",
      controlledVariables:
        "Bob mass (50 g), release angle (10°), same string and pivot, same photogate position, same release method, same room, and three trials at every length.",
      hypothesis:
        "If pendulum length increases, then the period will increase in proportion to the square root of length. Therefore, T² plotted against L will form a straight line through or close to the origin.",
      materials:
        "Retort stand with heavy base\nClamp and pivot support\nInextensible string (at least 1.10 m)\n50 g pendulum bob\nMeter ruler (+/- 0.001 m)\nProtractor (+/- 1°)\nPhotogate and electronic timer (+/- 0.01 s)\nSafety goggles",
      procedure:
        "1. Assemble the clamp stand on a level bench and attach the string and 50 g bob.\n2. Measure 0.20 m from the pivot to the center of the bob.\n3. Position the photogate at the equilibrium point.\n4. Use a protractor to displace the bob by 10° and release it without pushing.\n5. Record the time for 10 complete oscillations.\n6. Repeat the measurement three times at the same length.\n7. Repeat steps 2–6 for 0.40, 0.60, 0.80, and 1.00 m.\n8. Calculate the mean time, period, and period squared for every length.\n9. Plot T² against L and add a best-fit line.",
      rawDataNotes:
        "The pendulum motion remained in one vertical plane. No trial showed contact between the bob and the photogate. The amplitude decreased slightly during each timed set but remained below 10°.",
      processedDataNotes:
        "T² increased linearly as pendulum length increased. The best-fit line T² = 4.03L had R² = 0.999, indicating a very strong positive linear relationship. No point was an obvious outlier.",
      processedDataSampleCalculations:
        "For L = 0.60 m: mean time for 10 oscillations = (15.54 + 15.58 + 15.55) / 3 = 15.56 s. Period T = 15.56 / 10 = 1.556 s. T² = (1.556)² = 2.42 s². From slope m = 4.03 s²/m, g = 4π²/m = 9.80 m/s².",
      conclusion:
        "The results support the hypothesis. Increasing pendulum length increased the period, and the linear T² versus L graph agreed with the simple-pendulum model. The calculated gravitational field strength was 9.80 m/s², which is close to the accepted local value of approximately 9.81 m/s².",
      evaluation:
        "The strongest feature was the repeated timing across five evenly spaced lengths. Remaining limitations include uncertainty in locating the bob's center, small variation in release angle, pivot friction, and slight damping. The very high R² shows strong consistency, but it does not eliminate possible systematic error in length measurement.",
      improvements:
        "Use a fixed mechanical release to keep the starting angle constant, measure length with a set square at the pivot and bob center, collect five trials per length, and include additional lengths to strengthen the regression.",
      safetyConsiderations:
        "Wear goggles, secure the stand with a heavy base, keep faces and hands outside the bob's path, and stop the bob before adjusting length. No human or animal participants are involved. Reuse the string and bob, switch off the photogate after use, and avoid unnecessary material waste.",
      pilotObservations:
        "A pilot test at L = 0.60 m showed that timing one oscillation produced excessive relative uncertainty. Timing 10 oscillations gave repeatable values near 15.6 s. A 10° release angle produced stable motion and remained within the small-angle approximation.",
      references:
        "Ling, S. J., Moebs, W., & Sanny, J. (2016). University physics volume 1. OpenStax.\nSerway, R. A., & Jewett, J. W. (2018). Physics for scientists and engineers with modern physics (10th ed.). Cengage Learning."
    },
    setupDiagram: {
      dataUrl: images.diagram,
      title: "Simple pendulum experimental setup",
      description:
        "The string is attached to a fixed pivot on a secured clamp stand. Length L is measured from the pivot to the center of the 50 g bob. The bob is released from 10°, and a photogate at equilibrium records the oscillation time."
    },
    figures: [{
      dataUrl: images.graph,
      title: "Period squared versus pendulum length",
      description:
        "The data form a straight-line pattern with a positive slope of 4.03 s²/m and R² = 0.999. No measured point is an obvious outlier."
    }],
    tables: {
      rawData: [{
        title: "Table 1. Raw Timing Data for 10 Oscillations",
        headers: ["Length L (m)", "Trial 1 (s)", "Trial 2 (s)", "Trial 3 (s)", "Qualitative observation"],
        rows: [
          ["0.20", "8.99", "9.05", "9.02", "Stable, small amplitude"],
          ["0.40", "12.68", "12.73", "12.70", "Stable, small amplitude"],
          ["0.60", "15.54", "15.58", "15.55", "Stable, small amplitude"],
          ["0.80", "17.95", "18.00", "17.97", "Stable, small amplitude"],
          ["1.00", "20.04", "20.10", "20.07", "Stable, small amplitude"]
        ]
      }],
      processedData: [{
        title: "Table 2. Processed Pendulum Data",
        headers: ["Length L (m)", "Mean time for 10 (s)", "Period T (s)", "T² (s²)", "Time uncertainty (s)"],
        rows: [
          ["0.20", "9.02", "0.902", "0.81", "+/- 0.03"],
          ["0.40", "12.70", "1.270", "1.61", "+/- 0.03"],
          ["0.60", "15.56", "1.556", "2.42", "+/- 0.02"],
          ["0.80", "17.97", "1.797", "3.22", "+/- 0.03"],
          ["1.00", "20.07", "2.007", "4.03", "+/- 0.03"]
        ]
      }],
      dpRawData: defaultTableList("dpRawData"),
      dpProcessedData: defaultTableList("dpProcessedData")
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
  state.isSaving = false;
  state.pendingSave = false;
  state.imageUploadPending = false;
  state.setupDiagram = { dataUrl: "", title: "", description: "" };
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
    setupDiagram: state.setupDiagram,
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
  return LabFigures.hasTableContent(table);
}

function buildPrintableSections(report) {
  const sections = [];
  const program = PROGRAM_CONFIGS[report.program] ? report.program : "myp";
  const active = Array.isArray(report.activeSections?.[program])
    ? report.activeSections[program]
    : PROGRAM_CONFIGS[program].sections;

  sectionOrder.forEach((section) => {
    if (section.program !== program || !active.includes(section.key) || (section.dpOnly && report.studentProgramme !== "DP")) {
      return;
    }
    if (section.type === "background") {
      const labels = {
        backgroundPurpose: "Paragraph 1 - Purpose of the Investigation",
        backgroundScience: "Paragraph 2 - Scientific Information About the IV, DV, and Their Relationship"
      };
      const parts = section.fieldKeys
        .map((key) => ({ label: labels[key], text: String(report.sections?.[key] || "").trim() }))
        .filter((part) => part.text);
      const legacyText = String(report.sections?.backgroundInformation || "").trim();
      if (legacyText && parts.length === 0) {
        parts.push({ label: labels.backgroundPurpose, text: legacyText });
      }
      if (parts.length) {
        sections.push({ type: "structuredText", label: section.label, parts });
      }
      return;
    }
    if (section.type === "variables") {
      const labels = {
        independentVariable: "Independent Variable",
        dependentVariable: "Dependent Variable",
        controlledVariables: "Controlled Variables"
      };
      const parts = section.fieldKeys
        .map((key) => ({ label: labels[key], text: String(report.sections?.[key] || "").trim() }))
        .filter((part) => part.text);
      const legacyText = String(report.sections?.variables || "").trim();
      if (legacyText && parts.length === 0) {
        sections.push({ type: "text", label: section.label, text: legacyText });
        return;
      }
      if (parts.length) {
        sections.push({
          type: "text",
          label: section.label,
          text: parts.map((part) => `${part.label}\n${part.text}`).join("\n\n")
        });
      }
      return;
    }
    if (section.type === "diagram") {
      const diagram = normalizeSingleFigure(report.setupDiagram);
      if (diagram.dataUrl) {
        sections.push({
          type: "data",
          label: section.label,
          notes: "",
          sampleCalculations: "",
          tables: [],
          figures: [diagram]
        });
      }
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
    const figures = section.key === "processedData" ? LabFigures.normalize(report.figures).filter(figure => figure.dataUrl) : [];
    if (notes || sampleCalculations || contentTables.length > 0 || figures.length) {
      sections.push({
        type: "data",
        tableKey: section.key,
        label: section.label,
        notes,
        sampleCalculations,
        tables: contentTables,
        figures
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
  lines.push(`Programme: ${report.studentProgramme || "MYP"}`);
  lines.push(`Class Code: ${report.classCode || ""}`);
  lines.push(`Copy and Paste Attempts: ${report.blockedAttempts || 0}`);
  lines.push(`Time Spent: ${formatDuration(report.timeSpentSeconds || getTimeSpentSeconds())}`);
  lines.push("");

  printableSections.forEach((section, index) => {
    lines.push(`${index + 1}. ${section.label}`);
    if (section.type === "structuredText") {
      section.parts.forEach((part) => {
        lines.push(part.label);
        lines.push(...wrapPlainText(part.text));
        lines.push("");
      });
      return;
    }
    if (section.type === "text") {
      if (section.label === "Materials") {
        section.text.split("\n").forEach(item => lines.push(...wrapPlainText(item)));
      } else {
        lines.push(...wrapPlainText(section.text));
      }
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
    throw new Error("The PDF library is unavailable. Keep this page open so you do not lose your work, and try downloading again when it is available.");
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "letter" });
  if (!window.LabReportFonts?.normal || !window.LabReportFonts?.bold) {
    throw new Error("The PDF fonts are unavailable. Keep this page open so you do not lose your work, and try downloading again when they are available.");
  }
  doc.addFileToVFS("LiberationSerif-Regular.ttf", window.LabReportFonts.normal);
  doc.addFileToVFS("LiberationSerif-Bold.ttf", window.LabReportFonts.bold);
  doc.addFont("LiberationSerif-Regular.ttf", "LabReportSerif", "normal");
  doc.addFont("LiberationSerif-Bold.ttf", "LabReportSerif", "bold");
  const printableSections = buildPrintableSections(report);
  if (printableSections.some(section => section.tables?.length) && typeof doc.autoTable !== "function") {
    throw new Error("The PDF table library is unavailable. Keep this page open so you do not lose your work, and try downloading again when it is available.");
  }
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
    const lines = doc.setFont("LabReportSerif", bold ? "bold" : "normal").setFontSize(size).splitTextToSize(text, maxTextWidth);
    let nextLine = 0;
    // A paragraph can be longer than one page. Render only the lines that fit,
    // then continue on another page instead of drawing text below its edge.
    while (nextLine < lines.length) {
      ensureSpace(lineHeight + 4);
      const availableLines = Math.max(1, Math.floor((pageHeight - margin - y - 4) / lineHeight));
      const pageLines = lines.slice(nextLine, nextLine + availableLines);
      doc.text(pageLines, align === "center" ? pageWidth / 2 : margin, y, {
        align,
        baseline: "top",
        lineHeightFactor: lineHeight / size
      });
      y += pageLines.length * lineHeight;
      nextLine += pageLines.length;
    }
    y += 4;
  };

  const measureParagraphHeight = (text, { bold = false, size = 12, lineHeight = 16 } = {}) => {
    const lines = doc
      .setFont("LabReportSerif", bold ? "bold" : "normal")
      .setFontSize(size)
      .splitTextToSize(String(text || " "), maxTextWidth);
    return Math.max(1, lines.length) * lineHeight + 4;
  };

  const measureTableHeight = (table, includeTableNumber) => {
    let height = 14;
    if (table.title) {
      height += measureParagraphHeight(String(table.title), { bold: true, size: 11, lineHeight: 15 });
    }
    if (includeTableNumber) {
      height += measureParagraphHeight("Table 1", { bold: true, size: 11, lineHeight: 15 });
    }

    const columnCount = Math.max(1, table.headers.length);
    const columnWidth = maxTextWidth / columnCount;
    const tableFontSize = columnCount <= 5 ? 10 : columnCount <= 7 ? 9 : 8;
    const rowHeight = (row, minimumHeight, bold) => {
      doc.setFont("LabReportSerif", bold ? "bold" : "normal").setFontSize(tableFontSize);
      const contentHeight = Math.max(...Array.from({ length: columnCount }, (_, columnIndex) => {
        const cell = String(row[columnIndex] || " ");
        const lines = doc.splitTextToSize(cell, Math.max(12, columnWidth - 10));
        return Math.max(1, lines.length) * tableFontSize * 1.2 + 14;
      }));
      return Math.max(minimumHeight, contentHeight);
    };

    height += rowHeight(table.headers, 34, true);
    table.rows
      .filter((row) => row.some((cell) => String(cell || "").trim()))
      .forEach((row) => {
        height += rowHeight(row, 38, false);
      });
    return height;
  };

  const measureFigureHeight = (figure, figureIndex) => {
    const properties = doc.getImageProperties(figure.dataUrl);
    const scale = Math.min(maxTextWidth / properties.width, 300 / properties.height);
    const imageHeight = properties.height * scale;
    const title = `Figure ${figureIndex + 1}${figure.title ? `. ${figure.title}` : ""}`;
    let height = measureParagraphHeight(title, { bold: true, size: 12, lineHeight: 16 });
    height += imageHeight + 24;
    if (figure.description) {
      height += measureParagraphHeight(figure.description, { size: 11, lineHeight: 15 });
    }
    return height;
  };

  const measureSectionHeight = (section, index) => {
    let height = measureParagraphHeight(`${index + 1}. ${section.label}`, {
      bold: true,
      size: 13,
      lineHeight: 18
    });

    if (section.type === "text") {
      return height + measureParagraphHeight(section.text, { size: 12, lineHeight: 17 }) + 12;
    }
    if (section.type === "structuredText") {
      section.parts.forEach((part) => {
        height += measureParagraphHeight(part.label, { bold: true, size: 12, lineHeight: 16 });
        height += measureParagraphHeight(part.text, { size: 12, lineHeight: 17 }) + 8;
      });
      return height + 8;
    }

    if (section.notes) {
      height += measureParagraphHeight(section.notes, { size: 12, lineHeight: 17 }) + 4;
    }
    if (section.sampleCalculations) {
      height += measureParagraphHeight("Sample Calculations", { bold: true, size: 12, lineHeight: 16 });
      height += measureParagraphHeight(section.sampleCalculations, { size: 12, lineHeight: 17 }) + 4;
    }

    const sectionTables = Array.isArray(section.tables) ? section.tables : [];
    sectionTables.forEach((table) => {
      height += measureTableHeight(table, sectionTables.length > 1);
    });
    if (sectionTables.length === 0) {
      height += 6;
    }
    (section.figures || []).forEach((figure, figureIndex) => {
      height += measureFigureHeight(figure, figureIndex);
    });
    return height + 12;
  };

  const startSectionOnWholePageWhenPossible = (section, index) => {
    const bottomLimit = pageHeight - margin;
    const requiredHeight = measureSectionHeight(section, index);
    if (y > margin + 1 && y + requiredHeight > bottomLimit) {
      doc.addPage();
      y = margin;
    }
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
  drawParagraph(`Programme: ${report.studentProgramme || "MYP"}  |  Class Code: ${report.classCode || ""}`, {
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

  printableSections.forEach((section, index) => {
    startSectionOnWholePageWhenPossible(section, index);
    drawParagraph(`${index + 1}. ${section.label}`, { bold: true, size: 13, lineHeight: 18 });

    if (section.type === "text") {
      drawParagraph(section.text, { size: 12, lineHeight: 17 });
      y += 6;
      return;
    }
    if (section.type === "structuredText") {
      section.parts.forEach((part) => {
        drawParagraph(part.label, { bold: true, size: 12, lineHeight: 16 });
        drawParagraph(part.text, { size: 12, lineHeight: 17 });
        y += 4;
      });
      y += 2;
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
            font: "LabReportSerif",
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
    (section.figures || []).forEach((figure, figureIndex) => {
      const properties = doc.getImageProperties(figure.dataUrl);
      const scale = Math.min(maxTextWidth / properties.width, 300 / properties.height);
      const width = properties.width * scale;
      const height = properties.height * scale;
      const title = `Figure ${figureIndex + 1}${figure.title ? `. ${figure.title}` : ""}`;
      const titleHeight = doc.setFont("LabReportSerif", "bold").setFontSize(12).splitTextToSize(title, maxTextWidth).length * 16 + 4;
      ensureSpace(height + titleHeight + 24);
      drawParagraph(title, { bold: true });
      doc.addImage(figure.dataUrl, properties.fileType, (pageWidth - width) / 2, y, width, height);
      y += height + 12;
      if (figure.description) drawParagraph(figure.description, { size: 11, lineHeight: 15 });
      y += 12;
    });
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
    if (!(field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement) || field.readOnly || field.type === "date" || field.type === "file") return;
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
    sections[sectionKey] = ["materials", "dpMaterials"].includes(sectionKey)
      ? LabFigures.numberedMaterials(sectionInputs[sectionKey].value)
      : sectionInputs[sectionKey].value.trim();
  });

  return {
    schemaVersion: REPORT_SCHEMA_VERSION,
    id: state.reportId,
    accessToken: state.reportToken,
    teacherEmail: "",
    teacher: elements.teacher.value.trim(),
    classCode: state.classCode || elements.classCode.value.trim().toUpperCase(),
    studentProgramme: elements.selectedProgram.value === "DP" ? "DP" : "MYP",
    program: state.program,
    activeSections: state.activeSections,
    blockedAttempts: state.blockedAttempts,
    title: elements.title.value.trim(),
    studentName: elements.studentName.value.trim(),
    date: elements.date.value,
    time: elements.time.value,
    figures: state.figures,
    setupDiagram: normalizeSingleFigure({
      dataUrl: state.setupDiagram.dataUrl,
      title: elements.experimentalSetupTitle.value,
      description: elements.experimentalSetupDescription.value
    }),
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
  const figures = LabFigures.normalize(normalizedReport.figures);
  const setupDiagram = normalizeSingleFigure(normalizedReport.setupDiagram);

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
  state.program = "myp";
  elements.selectedProgram.value = normalizedReport.studentProgramme === "DP" || (!normalizedReport.studentProgramme && normalizedReport.program === "dp") ? "DP" : "MYP";
  if (Object.prototype.hasOwnProperty.call(normalizedReport, "classCode")) {
    state.classCode = String(normalizedReport.classCode || "").trim().toUpperCase();
  }
  state.blockedAttempts = Math.max(0, Number(normalizedReport.blockedAttempts) || 0);
  const defaults = createDefaultActiveSections();
  state.activeSections = Object.fromEntries(Object.keys(PROGRAM_CONFIGS).map((program) => {
    const source = normalizedReport.activeSections?.[program];
    if (Number(normalizedReport.schemaVersion) < REPORT_SCHEMA_VERSION) {
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
    if (["materials", "dpMaterials"].includes(sectionKey)) {
      sectionInputs[sectionKey].value = LabFigures.numberedMaterials(sectionInputs[sectionKey].value);
    }
  });
  if (!sectionInputs.backgroundPurpose.value && !sectionInputs.backgroundScience.value && sectionInputs.backgroundInformation.value) {
    sectionInputs.backgroundScience.value = sectionInputs.backgroundInformation.value;
    sectionInputs.backgroundInformation.value = "";
  }
  if (!sectionInputs.independentVariable.value && !sectionInputs.dependentVariable.value && !sectionInputs.controlledVariables.value && sectionInputs.variables.value) {
    sectionInputs.independentVariable.value = sectionInputs.variables.value;
    sectionInputs.variables.value = "";
  }

  state.tables.rawData = normalizeTableList(normalizedReport.tables?.rawData, "rawData");
  state.tables.processedData = normalizeTableList(normalizedReport.tables?.processedData, "processedData");
  state.tables.dpRawData = normalizeTableList(normalizedReport.tables?.dpRawData, "dpRawData");
  state.tables.dpProcessedData = normalizeTableList(normalizedReport.tables?.dpProcessedData, "dpProcessedData");
  state.status = normalizedReport.status === "Submitted" ? "Submitted" : "Draft";
  state.figures = figures;
  state.setupDiagram = setupDiagram;
  renderExperimentalSetup();
  renderGraphFigures();

  renderTableEditor("rawData", elements.rawDataEditor);
  renderTableEditor("processedData", elements.processedDataEditor);
  renderTableEditor("dpRawData", elements.dpRawDataEditor);
  renderTableEditor("dpProcessedData", elements.dpProcessedDataEditor);
  renderProgramUI();
  updateStatusBadge();
  setFormLocked(state.status === "Submitted");
  state.programmaticUpdate = false;
  maybeStartTimerFromStudentName();
}

function persistLocalBackup() {
  try {
  localStorage.setItem(REPORT_ID_KEY, state.reportId);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(collectReport()));
  return true;
  } catch (_error) {
    elements.saveState.textContent = "Draft could not be saved in this browser. Storage may be full; keep this page open.";
    return false;
  }
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

  if (!persistLocalBackup()) return;
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
    if (report.id !== state.reportId || state.status === "Submitted") return;
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
    if (report.id !== state.reportId || state.status === "Submitted") return;
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
    if (report.id !== state.reportId) return;
    state.isSaving = false;
    if (state.pendingSave) {
      state.pendingSave = false;
      void saveDraft("queued");
    }
  }
}

async function submitFinalReport() {
  if (state.imageUploadPending) {
    elements.saveState.textContent = "Wait for your graph images to finish uploading.";
    return;
  }
  if (state.status === "Submitted") {
    return;
  }

  const report = collectReport();
  if (!report.classCode) {
    window.alert("Enter the Class Code provided by your teacher before downloading the report.");
    elements.classCode.focus();
    return;
  }
  if (report.activeSections.myp.includes("experimentalSetup") && !report.setupDiagram.dataUrl && (report.setupDiagram.title || report.setupDiagram.description)) {
    elements.saveState.textContent = "Upload the experimental setup image, or clear its unfinished title and explanation.";
    return;
  }
  if (report.program === "myp" && report.activeSections.myp.includes("processedData") && report.figures.some(figure => !figure.dataUrl && (figure.title.trim() || figure.description.trim()))) {
    elements.saveState.textContent = "Upload an image for each graph with a title or description, or remove the unfinished graph.";
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

    if (report.id !== state.reportId) return;
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
    if (report.id !== state.reportId) return;
    elements.saveState.textContent = error.message || "Failed to generate final PDF.";
    updateSessionGate();
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
  const controls = document.querySelectorAll("input, textarea, select, button");
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
  updateSessionGate();
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
