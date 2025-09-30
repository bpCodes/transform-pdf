/* processor.js
 * Main Excel processing logic
 */

// ===== Configuration =====
const CONFIG = {
  SHEET_NAME: "",           // "" → primera hoja
  START_TEXT: "1 - Pumps",      // texto que marca inicio del rango (match insensible a may/min)
  END_TEXT: "2 - Pressure",     // texto que marca fin del rango (la fila donde se encuentra NO se procesa)
  FIND_COL_START: 3,        // desde qué columna (1-based) empieza a buscar START/END en cada fila
  VALUE_COL_START: 3       , // primera de las 3 columnas a evaluar/mover (1=A, 2=B, 3=C, etc.)
  SKIP_KEYWORDS: ["Total"]  // add more like ["total", "subtotal", "grand total"]
};

// ===== Main Processing Function =====
async function processFile(file, cfg = CONFIG) {
  console.log("process");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const sheetName = cfg.SHEET_NAME || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("No se encontró la hoja especificada.");

  // AOA (Array of Arrays) para manipular filas
  const aoa = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    blankrows: true,
    raw: true
  });

  // Detectar rango por START_TEXT → END_TEXT
  const startIdx = Utils.findRowIndexByText(aoa, cfg.START_TEXT, cfg.FIND_COL_START);
  if (startIdx === -1) throw new Error(`No se encontró START_TEXT: "${cfg.START_TEXT}"`);
  const endIdx = Utils.findRowIndexByText(aoa, cfg.END_TEXT, cfg.FIND_COL_START, startIdx + 1);
  // Si no hay END_TEXT, se procesa hasta el final
  const processEnd = endIdx === -1 ? aoa.length - 1 : endIdx - 1;

  // Procesar HEAD → FOOT dentro del rango [startIdx .. processEnd]
  const valueStart = Math.max(1, cfg.VALUE_COL_START) - 1; // a 0-based
  const newAOA = [];
  let removed = 0, consolidated = 0;

  for (let i = 0; i < aoa.length; i++) {
    const rowNumber = i + 1; // 1-based referencia
    const row = Array.isArray(aoa[i]) ? [...aoa[i]] : [];

    if (i >= startIdx && i <= (processEnd)) {
      if (Utils.shouldSkipRow(row, CONFIG.SKIP_KEYWORDS, CONFIG.FIND_COL_START)) {
        row[valueStart] = cfg.END_TEXT;
        newAOA.push(row);   // keep the row untouched
        continue;           // skip all processing (no consolidation, no X parsing)
      }
      // Tomar 3 celdas consecutivas: valueStart, valueStart+1, valueStart+2
      const v0 = Utils.valOrEmpty(row[valueStart]);
      const v1 = Utils.valOrEmpty(row[valueStart + 1]);
      const v2 = Utils.valOrEmpty(row[valueStart + 2]);

      const first = Utils.pickFirstNonEmpty([v0, v1, v2]);

      if (first === null) {
        // eliminar fila (no se agrega a newAOA)
        removed++;
        continue;
      } else {
        // mover a la primera columna del trío y vaciar las otras dos
        if (row[valueStart] !== first || v1 !== "" || v2 !== "") consolidated++;
        row[valueStart] = first;
        row[valueStart + 1] = "";
        row[valueStart + 2] = "";
      }
    }
    Utils.convertColumnXInRow(row);
    if (i === 21) { // fila 22 (0-based)
      const j = 9;  // columna J (0-based)
      const prev = row[j] == null ? "" : String(row[j]).trim();
      if (!/\s*m$/i.test(prev)) row[j] = prev ? `${prev} m` : " m";
    }
    newAOA.push(row);
  }

  // Reconstruir hoja
  const newWS = XLSX.utils.aoa_to_sheet(newAOA);
  Utils.ensureRef(newWS, newAOA);

  // Reemplazar y descargar
  wb.Sheets[sheetName] = newWS;
  const outName = file.name.replace(/(\.xlsx?|\.xlsb|\.xlsm)$/i, "") + "_clean.xlsx";
  XLSX.writeFile(wb, outName, { compression: true });

  console.info(`Filas consolidadas: ${consolidated}, filas eliminadas: ${removed}`);
  return outName;
}

// Export for use in other modules
window.Processor = {
  processFile,
  CONFIG
};
