const path = require("path");
const { randomUUID } = require("crypto");
const express = require("express");
const dotenv = require("dotenv");
const PDFDocument = require("pdfkit");
const { createClient } = require("@supabase/supabase-js");

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 3000);
const SUPABASE_TABLE = process.env.SUPABASE_TABLE || "lab_reports";

const sectionOrder = [
  { type: "text", key: "researchQuestion", label: "Research Question" },
  { type: "text", key: "backgroundInformation", label: "Background Information" },
  { type: "text", key: "variables", label: "Variables" },
  { type: "text", key: "hypothesis", label: "Hypothesis" },
  { type: "text", key: "materials", label: "Materials" },
  { type: "text", key: "procedure", label: "Procedure" },
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
  { type: "text", key: "references", label: "References (APA 7)" }
];

const supabase =
  process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY
    ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY)
    : null;

const memoryStore = new Map();

app.use(express.json({ limit: "4mb" }));
app.use(express.static(path.join(__dirname, "public")));

function cleanString(value) {
  return typeof value === "string" ? value.trim() : "";
}

function cleanMultiline(value) {
  return typeof value === "string" ? value.trim() : "";
}

function cleanNumber(value) {
  const parsed = Number(value);
  if (Number.isNaN(parsed)) {
    return 0;
  }
  return parsed;
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

  let headers = Array.isArray(table.headers)
    ? table.headers.map((header) => cleanString(String(header || ""))).slice(0, 12)
    : [];

  let rows = Array.isArray(table.rows) ? table.rows.slice(0, 300) : [];
  const maxRowColumns = rows.reduce((max, row) => {
    if (!Array.isArray(row)) {
      return max;
    }
    return Math.max(max, row.length);
  }, 0);

  const width = Math.max(headers.length, maxRowColumns, 2);
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
  const sections = {};

  sectionOrder.forEach((section) => {
    if (section.type === "text") {
      sections[section.key] = cleanMultiline(report.sections?.[section.key]);
      return;
    }
    if (section.type === "data" && section.noteKey) {
      sections[section.noteKey] = cleanMultiline(report.sections?.[section.noteKey]);
      if (section.sampleCalculationsKey) {
        sections[section.sampleCalculationsKey] = cleanMultiline(report.sections?.[section.sampleCalculationsKey]);
      }
    }
  });

  return {
    id: cleanString(report.id) || randomUUID(),
    teacherEmail: cleanString(report.teacherEmail),
    teacher: cleanString(report.teacher),
    title: cleanString(report.title),
    studentName: cleanString(report.studentName),
    date: cleanString(report.date),
    startedAt: cleanNumber(report.startedAt),
    timeSpentSeconds: cleanNumber(report.timeSpentSeconds),
    status: report.status === "Submitted" ? "Submitted" : "Draft",
    sections,
    tables: {
      rawData: normalizeTableList(report.tables?.rawData),
      processedData: normalizeTableList(report.tables?.processedData)
    },
    updatedAt: cleanString(report.updatedAt),
    submittedAt: cleanString(report.submittedAt)
  };
}

function tableHasContent(table) {
  if (!table) {
    return false;
  }

  const hasTitle = cleanString(String(table.title || "")).length > 0;
  const hasDataCells = table.rows.some((row) =>
    row.some((cell) => cleanString(String(cell)).length > 0)
  );

  const hasCustomHeaders = table.headers.some((header, index) => {
    const fallback = `Column ${index + 1}`;
    return cleanString(header) !== "" && cleanString(header) !== fallback;
  });

  return hasTitle || hasDataCells || hasCustomHeaders;
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

  const rows = [headers, ...dataRows];
  const columnCount = Math.max(headers.length, 1);
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const columnWidth = width / columnCount;
  const left = doc.page.margins.left;
  let currentY = doc.y;

  rows.forEach((row, rowIndex) => {
    const normalizedRow = Array.from({ length: columnCount }, (_, index) => row[index] || "");
    const cellHeights = normalizedRow.map((cell) =>
      doc.heightOfString(cell || " ", { width: columnWidth - 10, align: "left" })
    );
    const rowHeight = Math.max(...cellHeights, 16) + 10;
    const bottomLimit = doc.page.height - doc.page.margins.bottom - 22;

    if (currentY + rowHeight > bottomLimit) {
      doc.addPage();
      currentY = doc.y;
    }

    normalizedRow.forEach((cell, cellIndex) => {
      const x = left + columnWidth * cellIndex;
      doc.rect(x, currentY, columnWidth, rowHeight).lineWidth(0.8).strokeColor("#3f6b58").stroke();
      doc
        .font(rowIndex === 0 ? "Times-Bold" : "Times-Roman")
        .fontSize(11)
        .fillColor("#111111")
        .text(cell || " ", x + 5, currentY + 5, { width: columnWidth - 10, align: "left" });
    });

    currentY += rowHeight;
  });

  doc.y = currentY + 12;
}

function drawTextSection(doc, number, label, text) {
  drawSectionHeading(doc, number, label);
  doc
    .font("Times-Roman")
    .fontSize(12)
    .fillColor("#111111")
    .text(text, { align: "justify", lineGap: 4 });
  doc.moveDown(0.7);
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
  } else if (!notes && !sampleCalculations) {
    doc
      .font("Times-Italic")
      .fontSize(11)
      .fillColor("#4a4a4a")
      .text("No table data provided for this section.");
    doc.moveDown(0.6);
  }
}

function buildSectionsForPdf(report) {
  const ordered = [];

  sectionOrder.forEach((section) => {
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
      if (notes.length > 0 || sampleCalculations.length > 0 || tableListHasContent(tables)) {
        ordered.push({ type: "data", label: section.label, notes, sampleCalculations, tables });
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
        if (section.type === "text") {
          drawTextSection(doc, number, section.label, section.value);
        } else {
          drawDataSection(
            doc,
            number,
            section.label,
            section.notes,
            section.sampleCalculations,
            section.tables
          );
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
  if (!report.title) {
    return "Title of Experiment is required.";
  }
  if (!report.studentName) {
    return "Student Name is required.";
  }
  if (!report.date) {
    return "Date is required.";
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
    return res.json({ report });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Failed to load report." });
  }
});

app.post("/api/draft", async (req, res) => {
  try {
    const report = sanitizeReport(req.body?.report);
    const existing = await getReportById(report.id);
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
    return res.status(500).json({ error: error.message || "Failed to save draft." });
  }
});

app.post("/api/submit", async (req, res) => {
  try {
    const report = sanitizeReport(req.body?.report);
    const validationMessage = validateSubmitFields(report);
    if (validationMessage) {
      return res.status(400).json({ error: validationMessage });
    }

    const existing = await getReportById(report.id);
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
    return res.status(500).json({ error: error.message || "Failed to submit report." });
  }
});

app.use((req, res) => {
  if (req.path.startsWith("/api/")) {
    return res.status(404).json({ error: "Not found." });
  }
  return res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, () => {
  process.stdout.write(`Lab notebook server running on http://localhost:${PORT}\n`);
});
