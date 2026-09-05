/* Shared limits for locally saved and server-validated graph images. */
(function (root) {
  const MAX_COUNT = 6;
  const MAX_IMAGE_LENGTH = 350000;
  const MAX_TOTAL_LENGTH = 1800000;
  function normalize(value) {
    if (value == null) return [];
    if (!Array.isArray(value) || value.length > MAX_COUNT) throw new Error("A report supports up to 6 graph images.");
    let total = 0;
    return value.map((figure) => {
      const dataUrl = typeof figure?.dataUrl === "string" ? figure.dataUrl : "";
      total += dataUrl.length;
      if (dataUrl.length > MAX_IMAGE_LENGTH || total > MAX_TOTAL_LENGTH ||
          (dataUrl && !/^data:image\/(png|jpeg);base64,[A-Za-z0-9+/]+={0,2}$/.test(dataUrl))) {
        throw new Error("Graph images must be valid PNG or JPEG images within the report size limit.");
      }
      return {
        dataUrl,
        title: String(figure.title || "").trim().slice(0, 160),
        description: String(figure.description || "").trim().slice(0, 2000)
      };
    });
  }
  function hasTableContent(table) {
    if (!table || !Array.isArray(table.headers) || !Array.isArray(table.rows)) return false;
    const clean = value => String(value ?? "").trim();
    if (clean(table.title)) return true;
    const customHeader = table.headers.some((value, column) => {
      const text = clean(value);
      return text && !/^Column\s+\d+$/i.test(text) && !(column === 0 && /^Trial$/i.test(text));
    });
    const data = table.rows.some(row => Array.isArray(row) && row.some((value, column) => {
      const text = clean(value);
      return text && !(column === 0 && /^Trial\s+\d+$/i.test(text));
    }));
    return customHeader || data;
  }
  function numberedMaterials(value) {
    return String(value || "").split(/\r?\n/)
      .map(line => line.replace(/^\s*(?:\d+[.)]|[-*•])\s*/, "").trim())
      .filter(Boolean).map((line, index) => `${index + 1}. ${line}`).join("\n");
  }
  const api = { MAX_COUNT, MAX_IMAGE_LENGTH, MAX_TOTAL_LENGTH, normalize, hasTableContent, numberedMaterials };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.LabFigures = api;
})(typeof window !== "undefined" ? window : this);
