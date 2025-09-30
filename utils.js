/* utils.js
 * Utility functions for Excel processing
 */

// ===== Configuration =====
const COL_X = 24;                 // "X" = 24 (1-based). En AOA es índice 23.
const TREAT_PARENS_AS_NEGATIVE = false; // true => "(1,23)" -> -1.23 ; false => 1.23

// ===== Helper Functions =====

// Normaliza valores: string recortado; "nan" → vacío
function valOrEmpty(v) {
  if (v === undefined || v === null) return "";
  const s = String(v).trim();
  return s.toLowerCase() === "nan" ? "" : s;
}

function pickFirstNonEmpty(arr) {
  for (const v of arr) if (v !== "") return v;
  return null;
}

// Encuentra índice de fila que contenga "text" (case-insensitive) partiendo de findColStart (1-based)
// Busca en esa fila desde esa columna hacia la derecha hasta el final.
// startFromIndex: índice de fila (0-based) desde donde iniciar la búsqueda (default 0).
function findRowIndexByText(aoa, text, findColStart = 1, startFromIndex = 0) {
  const needle = String(text ?? "").trim().toLowerCase();
  if (!needle) return -1;
  const startCol = Math.max(1, findColStart) - 1;

  for (let r = startFromIndex; r < aoa.length; r++) {
    const row = Array.isArray(aoa[r]) ? aoa[r] : [];
    for (let c = startCol; c < row.length; c++) {
      const cell = row[c];
      if (cell == null) continue;
      const s = String(cell).trim().toLowerCase();
      if (s.includes(needle)) return r;
    }
  }
  return -1;
}

// Asegura el rango !ref de la hoja
function ensureRef(ws, aoa) {
  if (ws["!ref"]) return;
  const maxRow = Math.max(0, aoa.length - 1);
  const maxCol = Math.max(0, (aoa[0]?.length || 1) - 1);
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
}

// ===== Number Parsing Functions =====

// Conversión para una fila
function convertColumnXInRow(row) {
  const idx = COL_X - 1; // 0-based
  const raw = row[idx];
  const num = parseSingleEuroNumber(raw, { parensAsNegative: TREAT_PARENS_AS_NEGATIVE });
  row[idx] = Number.isFinite(num) ? num : ""; // si no se pudo parsear, dejar vacío
}

// Parser para un solo número en formato EU
function parseSingleEuroNumber(value, opts = {}) {
  const { parensAsNegative = false } = opts;
  if (value === undefined || value === null) return NaN;

  // String limpieza básica
  let s = String(value).trim();
  if (!s) return NaN;

  // Detectar paréntesis (contabilidad)
  const hasParens = /^\(.*\)$/.test(s);
  if (hasParens) s = s.replace(/[()]/g, "");

  // Quitar espacios
  s = s.replace(/\s+/g, "");

  // Si hay coma y punto, asumimos que el ÚLTIMO símbolo es el separador decimal.
  const lastComma = s.lastIndexOf(",");
  const lastDot   = s.lastIndexOf(".");
  const decimalSep = lastComma > lastDot ? "," : (lastDot >= 0 ? "." : ",");

  if (decimalSep === ",") {
    // puntos como miles -> fuera
    s = s.replace(/\./g, "");
    // la coma decimal -> punto
    const parts = s.split(",");
    if (parts.length > 1) {
      const dec = parts.pop();
      s = parts.join("") + "." + dec;
    } else {
      s = s.replace(",", ".");
    }
  } else {
    // comas como miles -> fuera
    s = s.replace(/,/g, "");
    // si hubiera múltiples puntos, dejamos solo el último como decimal
    const parts = s.split(".");
    if (parts.length > 2) {
      const dec = parts.pop();
      s = parts.join("") + "." + dec;
    }
  }

  let num = parseFloat(s);
  if (!Number.isFinite(num)) return NaN;

  if (hasParens && parensAsNegative) num = -num;
  return num;
}

// ===== Row Processing Functions =====

function shouldSkipRow(row, keywords = [], findColStart = 1) {
  if (!Array.isArray(keywords) || keywords.length === 0) return false;
  const startCol = Math.max(1, findColStart) - 1;
  const needles = keywords
    .map(k => String(k || "").trim().toLowerCase())
    .filter(Boolean);

  if (needles.length === 0) return false;

  for (let c = startCol; c < row.length; c++) {
    const cell = row[c];
    if (cell == null) continue;
    const s = String(cell).trim().toLowerCase();
    if (!s) continue;
    for (const n of needles) {
      if (s.includes(n)) return true;
    }
  }
  return false;
}

// Export functions for use in other modules
window.Utils = {
  valOrEmpty,
  pickFirstNonEmpty,
  findRowIndexByText,
  ensureRef,
  convertColumnXInRow,
  parseSingleEuroNumber,
  shouldSkipRow
};
