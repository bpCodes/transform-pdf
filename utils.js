/* utils.js
 * Utility functions for Excel processing
 */

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

// ===== Number Parsing =====

// Parser para un solo número en formato EU
function parseSingleEuroNumber(value, opts = {}) {
  const { parensAsNegative = false } = opts;
  if (value === undefined || value === null) return NaN;

  let s = String(value).trim();
  if (!s) return NaN;

  const hasParens = /^\(.*\)$/.test(s);
  if (hasParens) s = s.replace(/[()]/g, "");

  s = s.replace(/\s+/g, "");

  const lastComma = s.lastIndexOf(",");
  const lastDot   = s.lastIndexOf(".");
  const decimalSep = lastComma > lastDot ? "," : (lastDot >= 0 ? "." : ",");

  if (decimalSep === ",") {
    s = s.replace(/\./g, "");
    const parts = s.split(",");
    if (parts.length > 1) {
      const dec = parts.pop();
      s = parts.join("") + "." + dec;
    } else {
      s = s.replace(",", ".");
    }
  } else {
    s = s.replace(/,/g, "");
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

// ===== Cell Address Utilities =====

// Converts A1 notation ("J22") to 0-based {row, col}
function parseCellAddress(addr) {
  const match = String(addr).trim().match(/^([A-Za-z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell address: ${addr}`);
  const colStr = match[1].toUpperCase();
  const rowNum = parseInt(match[2], 10);
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { row: rowNum - 1, col: col - 1 };
}

// Converts 0-based {row, col} back to A1 notation
function encodeCellAddress(row, col) {
  let c = col + 1;
  let colStr = "";
  while (c > 0) {
    const rem = (c - 1) % 26;
    colStr = String.fromCharCode(65 + rem) + colStr;
    c = Math.floor((c - 1) / 26);
  }
  return colStr + (row + 1);
}

// Finds the first cell containing labelText (case-insensitive substring); returns {row, col} or null
function findLabelCell(aoa, labelText) {
  const needle = String(labelText).trim().toLowerCase();
  for (let r = 0; r < aoa.length; r++) {
    const row = Array.isArray(aoa[r]) ? aoa[r] : [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] == null) continue;
      if (String(row[c]).trim().toLowerCase().includes(needle)) return { row: r, col: c };
    }
  }
  return null;
}

// Sets aoa[row][col] = value, extending rows/cols as needed
function setCell(aoa, row, col, value) {
  while (aoa.length <= row) aoa.push([]);
  while (aoa[row].length <= col) aoa[row].push(undefined);
  aoa[row][col] = value;
}

// ===== Row Processing =====

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
  parseSingleEuroNumber,
  shouldSkipRow,
  parseCellAddress,
  encodeCellAddress,
  findLabelCell,
  setCell
};
