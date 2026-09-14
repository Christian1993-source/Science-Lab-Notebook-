const path = require("path");
const { randomUUID, timingSafeEqual } = require("crypto");
const express = require("express");
const dotenv = require("dotenv");
const PDFDocument = require("pdfkit");
const LabFigures = require("./figures.js");
const ReportFonts = require("./report-fonts.js");
const { createClient } = require("@supabase/supabase-js");

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3000);
const SUPABASE_TABLE = process.env.SUPABASE_TABLE || "lab_reports";

app.disable("x-powered-by");

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
  { type: "data", key: "rawData", noteKey: "rawDataNotes", label: "Raw Data" },
  {
    type: "data",
    key: "processedData",
    noteKey: "processedDataNotes",
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
  { type: "data", key: "dpRawData", noteKey: "dpRawDataNotes", label: "Raw Data", program: "dp" },
  {
    type: "data",
    key: "dpProcessedData",
    noteKey: "dpProcessedDataNotes",
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
  if (!section.program) section.program = "myp";
});

const programSections = {
  myp: sectionOrder.filter((section) => section.program === "myp").map((section) => section.key),
  dp: sectionOrder.filter((section) => section.program === "dp").map((section) => section.key)
};

const supabase =
  process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY
    ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY)
    : null;

const memoryStore = new Map();

app.use((_req, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=(), usb=()");
  res.setHeader(
    "Content-Security-Policy",
    "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-src https://christian1993-source.github.io; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'"
  );
  next();
});
app.use(express.json({ limit: "4mb" }));

const REPORT_LIMITS = Object.freeze({
  textLength: 100000,
  tablesPerSection: 20,
  rowsPerTable: 300,
  columnsPerTable: 12,
  references: 30,
  referenceUrlLength: 2048,
  controlledVariables: 30
});

function invalidReport(message, status = 400) {
  const error = new Error(message);
  error.status = status;
  return error;
}

app.get("/chemistry-lab-guide.pdf", (_req, res) => {
  res.sendFile(path.join(__dirname, "chemistry-lab-guide.pdf"));
});
app.get("/dp-physics-sl-lab-guide.pdf", (_req, res) => {
  res.sendFile(path.join(__dirname, "dp-physics-sl-lab-guide.pdf"));
});
const publicFiles = new Set([
  "app.js",
  "figures.js",
  "loader.js",
  "report-fonts.js",
  "styles.css",
  "atlas.css",
  "lab-hero.jpg",
  "favicon.svg",
  "favicon-32.png",
  "apple-touch-icon.png",
  "safari-pinned-tab.svg",
  "teacher.html",
  "teacher.js"
]);
app.get(["/", "/index.html"], (_req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.use("/vendor", express.static(path.join(__dirname, "vendor"), { dotfiles: "deny", fallthrough: false }));
app.get("/:publicFile", (req, res, next) => {
  if (!publicFiles.has(req.params.publicFile)) return next();
  return res.sendFile(path.join(__dirname, req.params.publicFile));
});

function cleanString(value) {
  if (typeof value === "string" && value.length > REPORT_LIMITS.textLength) {
    throw invalidReport(`Each text field supports up to ${REPORT_LIMITS.textLength.toLocaleString("en-US")} characters.`, 413);
  }
  return typeof value === "string" ? value.trim() : "";
}

function cleanMultiline(value) {
  return cleanString(value);
}

function cleanNumber(value) {
  const parsed = Number(value);
  if (Number.isNaN(parsed)) {
    return 0;
  }
  return parsed;
}

function normalizeReferenceUrl(value) {
  const raw = cleanString(value);
  if (!raw) return "";
  if (raw.length > REPORT_LIMITS.referenceUrlLength || /\s/.test(raw)) {
    throw invalidReport("Each reference link must be a valid web address up to 2,048 characters.");
  }
  const candidate = /^[a-z][a-z0-9+.-]*:\/\//i.test(raw) ? raw : `https://${raw}`;
  try {
    const parsed = new URL(candidate);
    if (!["http:", "https:"].includes(parsed.protocol) || !parsed.hostname) throw new Error("Invalid protocol");
    return parsed.href;
  } catch (_error) {
    throw invalidReport("Each reference link must be a valid http or https web address.");
  }
}

function normalizeReferences(value, legacyText = "") {
  if (Array.isArray(value) && value.length > REPORT_LIMITS.references) {
    throw invalidReport(`A report supports up to ${REPORT_LIMITS.references} references.`);
  }
  let source = Array.isArray(value) ? value : [];
  if (!source.length && cleanMultiline(legacyText)) {
    source = cleanMultiline(legacyText).split(/\n+/).map((citation) => ({ citation, url: "" }));
  }
  return source
    .slice(0, REPORT_LIMITS.references)
    .map((entry) => ({
      citation: cleanMultiline(entry?.citation),
      url: normalizeReferenceUrl(entry?.url)
    }))
    .filter((entry) => entry.citation || entry.url)
    .map((entry) => {
      if (entry.url && !entry.citation) {
        throw invalidReport("Every reference link must have typed APA 7 reference text.");
      }
      return entry;
    });
}

function normalizeControlledVariables(value, legacyText = "") {
  if (Array.isArray(value) && value.length > REPORT_LIMITS.controlledVariables) {
    throw invalidReport(`A report supports up to ${REPORT_LIMITS.controlledVariables} controlled variables.`);
  }
  let source = Array.isArray(value) ? value : [];
  if (!source.length && cleanMultiline(legacyText)) {
    source = cleanMultiline(legacyText)
      .split(/\n+/)
      .map((line) => ({ variable: line, control: "" }));
  }
  return source
    .slice(0, REPORT_LIMITS.controlledVariables)
    .map((row) => ({
      variable: cleanMultiline(row?.variable),
      control: cleanMultiline(row?.control)
    }))
    .filter((row) => row.variable || row.control);
}

function controlledVariablesAsTable(rows) {
  return {
    title: "",
    headers: ["Controlled Variable", "How It Will Be Controlled"],
    rows: rows.map((row) => [row.variable, row.control])
  };
}

function defaultTable() {
  return { title: "", headers: ["Column 1", "Column 2"], rows: [["", ""]] };
}

function defaultTableList() {
  return [defaultTable()];
}

function normalizeTable(table) {
  if (!table || typeof table !== "object") {
    return defaultTable();
  }

  const title = cleanString(String(table.title || ""));

  if (Array.isArray(table.headers) && table.headers.length > REPORT_LIMITS.columnsPerTable) {
    throw invalidReport(`Each table supports up to ${REPORT_LIMITS.columnsPerTable} columns.`);
  }
  if (Array.isArray(table.rows) && table.rows.length > REPORT_LIMITS.rowsPerTable) {
    throw invalidReport(`Each table supports up to ${REPORT_LIMITS.rowsPerTable} rows.`);
  }

  let headers = Array.isArray(table.headers)
    ? table.headers.map((header) => cleanString(String(header || "")))
    : [];

  let rows = Array.isArray(table.rows) ? table.rows : [];
  const maxRowColumns = rows.reduce((max, row) => {
    if (!Array.isArray(row)) {
      return max;
    }
    if (row.length > REPORT_LIMITS.columnsPerTable) {
      throw invalidReport(`Each table supports up to ${REPORT_LIMITS.columnsPerTable} columns.`);
    }
    return Math.max(max, row.length);
  }, 0);

  const width = Math.max(headers.length, maxRowColumns, 1);
  if (headers.length === 0) {
    headers = Array.from({ length: width }, (_, index) => `Column ${index + 1}`);
  }
  while (headers.length < width) {
    headers.push(`Column ${headers.length + 1}`);
  }

  rows = rows
    .filter((row) => Array.isArray(row))
    .map((row) => {
      const normalized = Array.from({ length: headers.length }, (_, index) => row[index] ?? "");
      return normalized.map((cell) => cleanString(String(cell)));
    });

  if (rows.length === 0) {
    rows = [Array(headers.length).fill("")];
  }

  return { title, headers, rows };
}

function normalizeTableList(tableValue) {
  if (Array.isArray(tableValue)) {
    if (tableValue.length > REPORT_LIMITS.tablesPerSection) {
      throw invalidReport(`Each data section supports up to ${REPORT_LIMITS.tablesPerSection} tables.`);
    }
    const normalized = tableValue.map((table) => normalizeTable(table));
    return normalized.length > 0 ? normalized : defaultTableList();
  }
  if (tableValue && typeof tableValue === "object") {
    return [normalizeTable(tableValue)];
  }
  return defaultTableList();
}

function sanitizeReport(rawReport) {
  const report = rawReport && typeof rawReport === "object" ? rawReport : {};
  let figures;
  let setupDiagram;
  try {
    figures = LabFigures.normalize(report.figures);
    setupDiagram = LabFigures.normalize(report.setupDiagram && typeof report.setupDiagram === "object" ? [report.setupDiagram] : [])[0]
      || { dataUrl: "", title: "", description: "" };
    const probe = new PDFDocument({ autoFirstPage: false });
    for (const figure of [...figures, setupDiagram]) {
      if (!figure.dataUrl) continue;
      const image = probe.openImage(Buffer.from(figure.dataUrl.split(",")[1], "base64"));
      if (!image.width || !image.height || image.width * image.height > 4000000) throw new Error("Graph dimensions are too large.");
    }
    probe.end();
  } catch (_error) {
    const error = new Error("Invalid uploaded image. Upload a PNG, JPG or WebP using an image section.");
    error.status = 400;
    throw error;
  }
  const sections = {};
  const program = programSections[report.program] ? report.program : "myp";

  sectionOrder.forEach((section) => {
    if (section.type === "background") {
      sections.backgroundInformation = cleanMultiline(report.sections?.backgroundInformation);
      section.fieldKeys.forEach((key) => {
        sections[key] = cleanMultiline(report.sections?.[key]);
      });
      return;
    }
    if (section.type === "variables") {
      sections.variables = cleanMultiline(report.sections?.variables);
      section.fieldKeys.forEach((key) => {
        sections[key] = cleanMultiline(report.sections?.[key]);
      });
      return;
    }
    if (section.type === "text") {
      sections[section.key] = ["materials", "dpMaterials"].includes(section.key)
        ? LabFigures.numberedMaterials(cleanMultiline(report.sections?.[section.key]))
        : cleanMultiline(report.sections?.[section.key]);
      return;
    }
    if (section.type === "data" && section.noteKey) {
      sections[section.noteKey] = cleanMultiline(report.sections?.[section.noteKey]);
      if (section.sampleCalculationsKey) {
        sections[section.sampleCalculationsKey] = cleanMultiline(report.sections?.[section.sampleCalculationsKey]);
      }
    }
  });

  const references = normalizeReferences(
    report.references,
    sections.references || sections.dpReferences
  );
  const controlledVariables = normalizeControlledVariables(
    report.controlledVariables,
    sections.controlledVariables
  );

  return {
    schemaVersion: Math.max(1, Math.floor(cleanNumber(report.schemaVersion) || 1)),
    id: cleanString(report.id) || randomUUID(),
    accessToken: cleanString(report.accessToken).slice(0, 128),
    teacherEmail: cleanString(report.teacherEmail),
    teacher: cleanString(report.teacher),
    classCode: cleanString(report.classCode).slice(0, 24).toUpperCase(),
    program,
    studentProgramme: report.studentProgramme === "DP" ? "DP" : "MYP",
    activeSections: {
      myp: Array.isArray(report.activeSections?.myp)
        ? programSections.myp.filter((key) => report.activeSections.myp.includes(key))
        : programSections.myp,
      dp: Array.isArray(report.activeSections?.dp)
        ? programSections.dp.filter((key) => report.activeSections.dp.includes(key))
        : programSections.dp
    },
    blockedAttempts: Math.max(0, Math.floor(cleanNumber(report.blockedAttempts))),
    title: cleanString(report.title),
    studentName: cleanString(report.studentName),
    date: cleanString(report.date),
    time: cleanString(report.time),
    figures,
    references,
    controlledVariables,
    setupDiagram,
    startedAt: cleanNumber(report.startedAt),
    timeSpentSeconds: cleanNumber(report.timeSpentSeconds),
    status: report.status === "Submitted" ? "Submitted" : "Draft",
    sections,
    tables: {
      rawData: normalizeTableList(report.tables?.rawData),
      processedData: normalizeTableList(report.tables?.processedData),
      dpRawData: normalizeTableList(report.tables?.dpRawData),
      dpProcessedData: normalizeTableList(report.tables?.dpProcessedData)
    },
    updatedAt: cleanString(report.updatedAt),
    submittedAt: cleanString(report.submittedAt)
  };
}

function tokenMatches(report, suppliedToken) {
  const expected = cleanString(report?.accessToken);
  const supplied = cleanString(suppliedToken);
  if (!expected || !supplied) {
    return false;
  }
  const expectedBuffer = Buffer.from(expected, "utf8");
  const suppliedBuffer = Buffer.from(supplied, "utf8");
  return expectedBuffer.length === suppliedBuffer.length && timingSafeEqual(expectedBuffer, suppliedBuffer);
}

function reportForClient(report) {
  const { accessToken: _accessToken, ...safeReport } = report;
  return safeReport;
}

function tableHasContent(table) {
  return LabFigures.hasTableContent(table);
}

function tableListHasContent(tableList) {
  return normalizeTableList(tableList).some((table) => tableHasContent(table));
}

function normalizeStoredRow(row) {
  const payload = row?.payload && typeof row.payload === "object" ? row.payload : {};
  const report = sanitizeReport({
    ...payload,
    id: row?.id || payload.id,
    teacherEmail: row?.teacher_email || payload.teacherEmail,
    title: row?.title || payload.title,
    studentName: row?.student_name || payload.studentName,
    date: row?.experiment_date || payload.date,
    status: row?.status || payload.status,
    updatedAt: row?.updated_at || payload.updatedAt,
    submittedAt: row?.submitted_at || payload.submittedAt
  });

  report.status = row?.status || payload.status || report.status;
  report.updatedAt = row?.updated_at || payload.updatedAt || "";
  report.submittedAt = row?.submitted_at || payload.submittedAt || "";

  return report;
}

function formatDuration(secondsInput) {
  const safeSeconds = Math.max(0, Math.round(cleanNumber(secondsInput)));
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const seconds = safeSeconds % 60;

  if (hours > 0) {
    return `${hours}h ${minutes}m ${seconds}s`;
  }
  if (minutes > 0) {
    return `${minutes}m ${seconds}s`;
  }
  return `${seconds}s`;
}

async function getReportById(reportId) {
  const id = cleanString(reportId);
  if (!id) {
    return null;
  }

  if (supabase) {
    const { data, error } = await supabase.from(SUPABASE_TABLE).select("*").eq("id", id).maybeSingle();
    if (error) {
      throw error;
    }
    return data ? normalizeStoredRow(data) : null;
  }

  return memoryStore.get(id) || null;
}

async function saveReport(report, targetStatus) {
  const status = targetStatus === "Submitted" ? "Submitted" : "Draft";
  const now = new Date().toISOString();
  const sanitized = sanitizeReport(report);
  const saved = {
    ...sanitized,
    status,
    updatedAt: now,
    submittedAt: status === "Submitted" ? now : sanitized.submittedAt || ""
  };

  if (supabase) {
    const row = {
      id: saved.id,
      teacher_email: saved.teacherEmail,
      title: saved.title,
      student_name: saved.studentName,
      experiment_date: saved.date,
      status: saved.status,
      payload: saved,
      updated_at: saved.updatedAt,
      submitted_at: saved.submittedAt || null
    };

    const { error } = await supabase.from(SUPABASE_TABLE).upsert(row, { onConflict: "id" });
    if (error) {
      throw error;
    }
  } else {
    memoryStore.set(saved.id, saved);
  }

  return saved;
}

function ensurePageSpace(doc, minHeight) {
  const bottomLimit = doc.page.height - doc.page.margins.bottom;
  if (doc.y + minHeight > bottomLimit) {
    doc.addPage();
  }
}

function measureTableGridHeight(doc, table) {
  const headers = table.headers.map((header) => cleanString(header));
  const dataRows = table.rows
    .map((row) => row.map((cell) => cleanString(String(cell))))
    .filter((row) => row.some((cell) => cell.length > 0));
  const columnCount = Math.max(headers.length, 1);
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const columnWidth = width / columnCount;
  const fontSize = columnCount <= 5 ? 10 : columnCount <= 7 ? 9 : 8;
  const horizontalPadding = 5;
  const verticalPadding = 7;
  const getRowHeight = (row, isHeader) => {
    doc.font(isHeader ? "Times-Bold" : "Times-Roman").fontSize(fontSize);
    const normalizedRow = Array.from({ length: columnCount }, (_, index) => row[index] || "");
    const cellHeights = normalizedRow.map((cell) =>
      doc.heightOfString(cell || " ", {
        width: columnWidth - horizontalPadding * 2,
        align: "left"
      })
    );
    return Math.max(...cellHeights, isHeader ? 20 : 24) + verticalPadding * 2;
  };

  return getRowHeight(headers, true) + dataRows.reduce((total, row) => total + getRowHeight(row, false), 0) + 12;
}

function measureSectionHeight(doc, number, section) {
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const textHeight = (text, { font = "Times-Roman", size = 12, lineGap = 0 } = {}) => {
    doc.font(font).fontSize(size);
    return doc.heightOfString(String(text || " "), { width, lineGap });
  };
  let height = textHeight(`${number}. ${section.label}`, { font: "Times-Bold", size: 13 }) + 14;

  if (section.type === "text") {
    return height + textHeight(section.value, { size: 12, lineGap: 4 }) + 18;
  }
  if (section.type === "variables") {
    section.parts.forEach((part) => {
      height += textHeight(part.label, { font: "Times-Bold", size: 12 }) + 6;
      height += textHeight(part.value, { size: 12, lineGap: 4 }) + 12;
    });
    if (section.controlledVariables.length) {
      height += textHeight("Controlled Variables", { font: "Times-Bold", size: 12 }) + 8;
      height += measureTableGridHeight(doc, controlledVariablesAsTable(section.controlledVariables));
    }
    return height + 8;
  }
  if (section.type === "references") {
    section.entries.forEach((entry, entryIndex) => {
      height += textHeight(`${entryIndex + 1}. ${entry.citation}`, { size: 12, lineGap: 4 }) + 6;
      if (entry.url) height += textHeight(entry.url, { size: 11, lineGap: 2 }) + 8;
      height += 8;
    });
    return height + 6;
  }
  if (section.type === "structuredText") {
    section.parts.forEach((part) => {
      height += textHeight(part.label, { font: "Times-Bold", size: 12 }) + 6;
      height += textHeight(part.value, { size: 12, lineGap: 4 }) + 12;
    });
    return height + 8;
  }

  if (section.notes) {
    height += textHeight(section.notes, { size: 12, lineGap: 4 }) + 12;
  }
  if (section.sampleCalculations) {
    height += textHeight("Sample Calculations", { font: "Times-Bold", size: 12 }) + 6;
    height += textHeight(section.sampleCalculations, { size: 12, lineGap: 4 }) + 12;
  }

  const contentTables = normalizeTableList(section.tables).filter((table) => tableHasContent(table));
  contentTables.forEach((table) => {
    if (table.title) {
      height += textHeight(String(table.title), { font: "Times-Bold", size: 11 }) + 5;
    }
    if (contentTables.length > 1) {
      height += textHeight("Table 1", { font: "Times-Bold", size: 11 }) + 6;
    }
    height += measureTableGridHeight(doc, table);
  });

  (section.figures || []).forEach((figure, figureIndex) => {
    const image = doc.openImage(Buffer.from(figure.dataUrl.split(",")[1], "base64"));
    const scale = Math.min(width / image.width, 300 / image.height);
    const title = `Figure ${figureIndex + 1}${figure.title ? `. ${figure.title}` : ""}`;
    height += textHeight(title, { font: "Times-Bold", size: 12 }) + image.height * scale + 24;
    if (figure.description) {
      height += textHeight(figure.description, { size: 11, lineGap: 3 }) + 12;
    }
  });

  return height + 12;
}

function startSectionOnWholePageWhenPossible(doc, number, section) {
  const bottomLimit = doc.page.height - doc.page.margins.bottom;
  const requiredHeight = measureSectionHeight(doc, number, section);
  if (doc.y > doc.page.margins.top + 1 && doc.y + requiredHeight > bottomLimit) {
    doc.addPage();
  }
}

function drawSectionHeading(doc, number, label) {
  ensurePageSpace(doc, 48);
  doc.font("Times-Bold").fontSize(13).fillColor("#124232").text(`${number}. ${label}`);
  doc.moveDown(0.3);
}

function drawTableGrid(doc, table) {
  const headers = table.headers.map((header) => cleanString(header));
  const dataRows = table.rows
    .map((row) => row.map((cell) => cleanString(String(cell))))
    .filter((row) => row.some((cell) => cell.length > 0));

  const columnCount = Math.max(headers.length, 1);
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const columnWidth = width / columnCount;
  const left = doc.page.margins.left;
  const fontSize = columnCount <= 5 ? 10 : columnCount <= 7 ? 9 : 8;
  const horizontalPadding = 5;
  const verticalPadding = 7;
  let currentY = doc.y;
  const bottomLimit = () => doc.page.height - doc.page.margins.bottom - 22;

  const getRowHeight = (row, isHeader) => {
    doc.font(isHeader ? "Times-Bold" : "Times-Roman").fontSize(fontSize);
    const cellHeights = row.map((cell) =>
      doc.heightOfString(cell || " ", {
        width: columnWidth - horizontalPadding * 2,
        align: "left"
      })
    );
    return Math.max(...cellHeights, isHeader ? 20 : 24) + verticalPadding * 2;
  };

  const drawRow = (row, isHeader) => {
    const normalizedRow = Array.from({ length: columnCount }, (_, index) => row[index] || "");
    const rowHeight = getRowHeight(normalizedRow, isHeader);

    normalizedRow.forEach((cell, cellIndex) => {
      const x = left + columnWidth * cellIndex;
      doc.save();
      doc.rect(x, currentY, columnWidth, rowHeight).lineWidth(0.8);
      if (isHeader) {
        doc.fillAndStroke("#e8f1ec", "#3f6b58");
      } else {
        doc.strokeColor("#3f6b58").stroke();
      }
      doc.restore();
      doc
        .font(isHeader ? "Times-Bold" : "Times-Roman")
        .fontSize(fontSize)
        .fillColor(isHeader ? "#143427" : "#111111")
        .text(cell || " ", x + horizontalPadding, currentY + verticalPadding, {
          width: columnWidth - horizontalPadding * 2,
          align: "left"
        });
    });

    currentY += rowHeight;
  };

  const normalizedHeaders = Array.from({ length: columnCount }, (_, index) => headers[index] || "");
  const firstBodyRow = dataRows[0] || Array(columnCount).fill("");
  const headerHeight = getRowHeight(normalizedHeaders, true);
  const fullPageHeight = bottomLimit() - doc.page.margins.top;
  const bodyRowsToCheck = dataRows.length ? dataRows : [firstBodyRow];
  if (bodyRowsToCheck.some(row => headerHeight + getRowHeight(row, false) > fullPageHeight)) {
    doc.destroy();
    throw invalidReport("A table row is too tall to fit on one PDF page. Shorten the table cells or move detailed explanations to the notes, then download again.");
  }
  const openingHeight = headerHeight + getRowHeight(firstBodyRow, false);
  if (currentY + openingHeight > bottomLimit()) {
    doc.addPage();
    currentY = doc.y;
  }
  drawRow(normalizedHeaders, true);

  dataRows.forEach((row) => {
    const normalizedRow = Array.from({ length: columnCount }, (_, index) => row[index] || "");
    const rowHeight = getRowHeight(normalizedRow, false);
    if (currentY + rowHeight > bottomLimit()) {
      doc.addPage();
      currentY = doc.y;
      drawRow(normalizedHeaders, true);
    }
    drawRow(normalizedRow, false);
  });

  doc.x = left;
  doc.y = currentY + 12;
}

function drawTextSection(doc, number, label, text) {
  drawSectionHeading(doc, number, label);
  doc
    .font("Times-Roman")
    .fontSize(12)
    .fillColor("#111111")
    .text(text, { align: label === "Materials" ? "left" : "justify", lineGap: 4 });
  doc.moveDown(0.7);
}

function drawStructuredTextSection(doc, number, label, parts) {
  drawSectionHeading(doc, number, label);
  parts.forEach((part) => {
    doc.font("Times-Bold").fontSize(12).fillColor("#124232").text(part.label);
    doc.moveDown(0.2);
    doc
      .font("Times-Roman")
      .fontSize(12)
      .fillColor("#111111")
      .text(part.value, { align: "justify", lineGap: 4 });
    doc.moveDown(0.6);
  });
}

function drawVariablesSection(doc, number, label, parts, controlledVariables) {
  drawSectionHeading(doc, number, label);
  parts.forEach((part) => {
    doc.font("Times-Bold").fontSize(12).fillColor("#124232").text(part.label);
    doc.moveDown(0.2);
    doc
      .font("Times-Roman")
      .fontSize(12)
      .fillColor("#111111")
      .text(part.value, { align: "justify", lineGap: 4 });
    doc.moveDown(0.6);
  });
  if (controlledVariables.length) {
    doc.font("Times-Bold").fontSize(12).fillColor("#124232").text("Controlled Variables");
    doc.moveDown(0.3);
    drawTableGrid(doc, controlledVariablesAsTable(controlledVariables));
  }
}

function drawReferencesSection(doc, number, label, entries) {
  drawSectionHeading(doc, number, label);
  entries.forEach((entry, entryIndex) => {
    doc
      .font("Times-Roman")
      .fontSize(12)
      .fillColor("#111111")
      .text(`${entryIndex + 1}. ${entry.citation}`, { align: "left", lineGap: 4 });
    if (entry.url) {
      doc.moveDown(0.15);
      doc
        .font("Times-Roman")
        .fontSize(11)
        .fillColor("#086c6c")
        .text(entry.url, { link: entry.url, underline: true, lineGap: 2 });
    }
    doc.moveDown(0.7);
  });
}

function drawDataSection(doc, number, label, notes, sampleCalculations, tables) {
  drawSectionHeading(doc, number, label);

  if (notes) {
    doc
      .font("Times-Roman")
      .fontSize(12)
      .fillColor("#111111")
      .text(notes, { align: "justify", lineGap: 4 });
    doc.moveDown(0.5);
  }

  if (sampleCalculations) {
    doc.font("Times-Bold").fontSize(12).fillColor("#124232").text("Sample Calculations");
    doc.moveDown(0.2);
    doc
      .font("Times-Roman")
      .fontSize(12)
      .fillColor("#111111")
      .text(sampleCalculations, { align: "justify", lineGap: 4 });
    doc.moveDown(0.5);
  }

  const contentTables = normalizeTableList(tables).filter((table) => tableHasContent(table));
  if (contentTables.length > 0) {
    contentTables.forEach((table, index) => {
      ensurePageSpace(doc, 100);
      if (table.title) {
        doc.font("Times-Bold").fontSize(11).fillColor("#124232").text(String(table.title));
        doc.moveDown(0.2);
      }
      if (contentTables.length > 1) {
        doc.font("Times-Bold").fontSize(11).fillColor("#124232").text(`Table ${index + 1}`);
        doc.moveDown(0.25);
      }
      drawTableGrid(doc, table);
    });
  }
}

function buildSectionsForPdf(report) {
  const ordered = [];
  const active = Array.isArray(report.activeSections?.[report.program])
    ? report.activeSections[report.program]
    : programSections[report.program];

  sectionOrder.forEach((section) => {
    if (section.program !== report.program || !active.includes(section.key) || (section.dpOnly && report.studentProgramme !== "DP")) return;
    if (section.type === "background") {
      const labels = {
        backgroundPurpose: "Paragraph 1 - Purpose of the Investigation",
        backgroundScience: "Paragraph 2 - Scientific Information About the IV, DV, and Their Relationship"
      };
      const parts = section.fieldKeys
        .map((key) => ({ label: labels[key], value: cleanMultiline(report.sections[key]) }))
        .filter((part) => part.value);
      const legacyValue = cleanMultiline(report.sections.backgroundInformation);
      if (legacyValue && parts.length === 0) {
        parts.push({ label: labels.backgroundPurpose, value: legacyValue });
      }
      if (parts.length) {
        ordered.push({
          type: "text",
          label: section.label,
          value: parts.map((part) => part.value).join("\n\n")
        });
      }
      return;
    }
    if (section.type === "variables") {
      const labels = {
        independentVariable: "Independent Variable",
        dependentVariable: "Dependent Variable"
      };
      const parts = ["independentVariable", "dependentVariable"]
        .map((key) => ({ label: labels[key], value: cleanMultiline(report.sections[key]) }))
        .filter((part) => part.value);
      const controlledVariables = normalizeControlledVariables(
        report.controlledVariables,
        report.sections.controlledVariables
      );
      const legacyValue = cleanMultiline(report.sections.variables);
      if (legacyValue && parts.length === 0 && controlledVariables.length === 0) {
        ordered.push({ type: "text", label: section.label, value: legacyValue });
        return;
      }
      if (parts.length || controlledVariables.length) {
        ordered.push({
          type: "variables",
          label: section.label,
          parts,
          controlledVariables
        });
      }
      return;
    }
    if (section.type === "diagram") {
      if (report.setupDiagram?.dataUrl) {
        ordered.push({
          type: "data",
          label: section.label,
          notes: "",
          sampleCalculations: "",
          tables: [],
          figures: [report.setupDiagram]
        });
      }
      return;
    }
    if (["references", "dpReferences"].includes(section.key)) {
      const entries = normalizeReferences(report.references, report.sections[section.key]);
      if (entries.length) {
        ordered.push({ type: "references", label: section.label, entries });
      }
      return;
    }
    if (section.type === "text") {
      const text = cleanMultiline(report.sections[section.key]);
      if (text.length > 0) {
        ordered.push({ type: "text", label: section.label, value: text });
      }
      return;
    }

    if (section.type === "data") {
      const notes = cleanMultiline(report.sections[section.noteKey]);
      const sampleCalculations = cleanMultiline(report.sections[section.sampleCalculationsKey]);
      const tables = normalizeTableList(report.tables?.[section.key]);
      const figures = section.key === "processedData" ? (report.figures || []).filter(figure => figure.dataUrl) : [];
      if (notes.length > 0 || sampleCalculations.length > 0 || tableListHasContent(tables) || figures.length) {
        ordered.push({ type: "data", label: section.label, notes, sampleCalculations, tables, figures });
      }
    }
  });

  return ordered;
}

function generatePdf(report) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "LETTER",
      margins: { top: 72, right: 72, bottom: 72, left: 72 }
    });
    const chunks = [];

    // Embed the same Unicode serif fonts used by the browser PDF exporter.
    doc.registerFont("Times-Roman", Buffer.from(ReportFonts.normal, "base64"));
    doc.registerFont("Times-Bold", Buffer.from(ReportFonts.bold, "base64"));

    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    doc.font("Times-Bold").fontSize(20).fillColor("#123f2e").text(report.title, {
      align: "center"
    });
    doc.moveDown(0.3);
    doc.font("Times-Roman").fontSize(12).fillColor("#111111").text(`Teacher: ${report.teacher || "Not specified"}`, {
      align: "center"
    });
    doc.font("Times-Roman").fontSize(12).fillColor("#111111").text(`Student: ${report.studentName}`, {
      align: "center"
    });
    doc.text(`Date: ${report.date}`, { align: "center" });
    doc.text(`Time: ${report.time}`, { align: "center" });
    doc.text(`Programme: ${report.studentProgramme || report.program.toUpperCase()} | Class Code: ${report.classCode}`, { align: "center" });
    doc.fontSize(10).fillColor("#4b5563").text(`Copy and Paste Attempts: ${report.blockedAttempts}`, { align: "center" });
    doc.text(`Time Spent: ${formatDuration(report.timeSpentSeconds)}`, { align: "center" });
    doc.moveDown(1);

    const printableSections = buildSectionsForPdf(report);
    if (printableSections.length === 0) {
      doc.font("Times-Italic").fontSize(12).fillColor("#333333").text("No sections with content.", {
        align: "left"
      });
    } else {
      printableSections.forEach((section, index) => {
        const number = index + 1;
        startSectionOnWholePageWhenPossible(doc, number, section);
        if (section.type === "text") {
          drawTextSection(doc, number, section.label, section.value);
        } else if (section.type === "variables") {
          drawVariablesSection(doc, number, section.label, section.parts, section.controlledVariables);
        } else if (section.type === "references") {
          drawReferencesSection(doc, number, section.label, section.entries);
        } else if (section.type === "structuredText") {
          drawStructuredTextSection(doc, number, section.label, section.parts);
        } else {
          drawDataSection(
            doc,
            number,
            section.label,
            section.notes,
            section.sampleCalculations,
            section.tables
          );
          (section.figures || []).forEach((figure, figureIndex) => {
            const image = doc.openImage(Buffer.from(figure.dataUrl.split(",")[1], "base64"));
            const widthLimit = doc.page.width - 144;
            const scale = Math.min(widthLimit / image.width, 300 / image.height);
            const width = image.width * scale;
            const height = image.height * scale;
            const title = `Figure ${figureIndex + 1}${figure.title ? `. ${figure.title}` : ""}`;
            doc.font("Times-Bold").fontSize(12);
            ensurePageSpace(doc, height + doc.heightOfString(title, { width: widthLimit }) + 30);
            doc.fillColor("#124232").text(title, 72, doc.y, { width: widthLimit });
            doc.moveDown(0.4);
            const imageY = doc.y;
            doc.image(image, (doc.page.width - width) / 2, imageY, { width, height });
            doc.y = imageY + height + 12;
            if (figure.description) doc.font("Times-Roman").fontSize(11).fillColor("#111111").text(figure.description, 72, doc.y, { width: widthLimit, lineGap: 3 });
            doc.moveDown(0.8);
          });
        }
      });
    }

    doc.end();
  });
}

function safeFileName(baseName) {
  const cleaned = cleanString(baseName)
    .replace(/[^a-zA-Z0-9-_ ]+/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();

  return cleaned || "lab-report";
}

function validateSubmitFields(report) {
  if (!report.classCode) {
    return "Class Code is required.";
  }
  if (!report.title) {
    return "Title of Experiment is required.";
  }
  if (!report.studentName) {
    return "Student Name is required.";
  }
  if (!report.date) {
    return "Date is required.";
  }
  if (!report.time) {
    return "Time is required.";
  }
  return "";
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, database: supabase ? "supabase" : "memory" });
});

app.get("/api/report/:id", async (req, res) => {
  try {
    const report = await getReportById(req.params.id);
    if (!report) {
      return res.status(404).json({ error: "Report not found." });
    }
    if (!tokenMatches(report, req.get("x-report-token"))) {
      return res.status(404).json({ error: "Report not found." });
    }
    return res.json({ report: reportForClient(report) });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Failed to load report." });
  }
});

app.post("/api/draft", async (req, res) => {
  try {
    const report = sanitizeReport(req.body?.report);
    if (!report.accessToken) {
      return res.status(400).json({ error: "Report access token is required." });
    }
    const existing = await getReportById(report.id);
    if (existing?.accessToken && !tokenMatches(existing, report.accessToken)) {
      return res.status(404).json({ error: "Report not found." });
    }
    if (existing?.status === "Submitted") {
      return res.status(409).json({ error: "Report already submitted and locked." });
    }

    const saved = await saveReport(report, "Draft");
    return res.json({
      ok: true,
      reportId: saved.id,
      status: saved.status,
      updatedAt: saved.updatedAt
    });
  } catch (error) {
    return res.status(error.status || 500).json({ error: error.message || "Failed to save draft." });
  }
});

app.post("/api/submit", async (req, res) => {
  try {
    const report = sanitizeReport(req.body?.report);
    if (!report.accessToken) {
      return res.status(400).json({ error: "Report access token is required." });
    }
    const validationMessage = validateSubmitFields(report);
    if (validationMessage) {
      return res.status(400).json({ error: validationMessage });
    }

    const existing = await getReportById(report.id);
    if (existing?.accessToken && !tokenMatches(existing, report.accessToken)) {
      return res.status(404).json({ error: "Report not found." });
    }
    if (existing?.status === "Submitted") {
      return res.status(409).json({ error: "Report already submitted and locked." });
    }

    const finalReport = { ...report, status: "Submitted" };
    const pdfBuffer = await generatePdf(finalReport);
    await saveReport(finalReport, "Submitted");

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${safeFileName(finalReport.title)}.pdf"`);
    return res.send(pdfBuffer);
  } catch (error) {
    return res.status(error.status || 500).json({ error: error.message || "Failed to submit report." });
  }
});

app.use((req, res) => {
  if (req.path.startsWith("/api/")) {
    return res.status(404).json({ error: "Not found." });
  }
  return res.sendFile(path.join(__dirname, "index.html"));
});

// Parser and static-file errors must not reveal stack traces or local paths.
app.use((error, _req, res, _next) => {
  if (error.type === "entity.too.large") {
    return res.status(413).json({ error: "The report is too large. The maximum request size is 4 MB." });
  }
  if (error.type === "entity.parse.failed") {
    return res.status(400).json({ error: "The report must contain valid JSON." });
  }
  const status = error.status >= 400 && error.status < 500 ? error.status : 500;
  return res.status(status).json({ error: status === 404 ? "Not found." : "The request could not be processed." });
});

app.listen(PORT, () => {
  process.stdout.write(`Lab notebook server running on http://localhost:${PORT}\n`);
});
