/* processor.js
 * Main Excel processing logic
 */

const TEMPLATE_NAME = "hennecke-2025-02";
const SKIP_KEYWORDS = ["Total"];

// ===== Template Loading =====

async function loadTemplate(name) {
  const resp = await fetch(`./templates/${name}.json`);
  if (!resp.ok) throw new Error(`Template not found: ${name}`);
  return resp.json();
}

// ===== Field Handlers =====

function applyScalarField(aoa, field) {
  const labelCell = Utils.findLabelCell(aoa, field.label);
  if (!labelCell) throw new Error(`Field "${field.name}": label "${field.label}" not found in input`);

  const row = aoa[labelCell.row];
  let srcCol = null;
  let srcVal = null;
  for (let c = labelCell.col + 1; c < row.length; c++) {
    const v = Utils.valOrEmpty(row[c]);
    // skip intermediate sub-labels (cells that end with ":")
    if (v !== "" && !v.endsWith(":")) { srcCol = c; srcVal = v; break; }
  }
  if (srcVal === null) throw new Error(`Field "${field.name}": no value found to the right of label "${field.label}"`);

  let val = srcVal;
  if (field.number_format === "eu") {
    const num = Utils.parseSingleEuroNumber(srcVal);
    if (Number.isFinite(num)) val = num;
  }

  if (field.suffix) {
    const s = String(val).trim();
    val = s.endsWith(field.suffix) ? s : s + field.suffix;
  }

  const target = Utils.parseCellAddress(field.target_cell);
  Utils.setCell(aoa, target.row, target.col, val);

  if (srcCol !== target.col || labelCell.row !== target.row) {
    aoa[labelCell.row][srcCol] = "";
  }
}

function applyTableColumnField(aoa, field) {
  const startIdx = Utils.findRowIndexByText(aoa, field.section_start, 1);
  if (startIdx === -1) throw new Error(`Field "${field.name}": section_start "${field.section_start}" not found`);

  const stopIdx = Utils.findRowIndexByText(aoa, field.section_stop, 1, startIdx + 1);
  if (stopIdx === -1) throw new Error(`Field "${field.name}": section_stop "${field.section_stop}" not found`);

  // Find value_header column: search 5 rows before and after section start
  let headerCol = null;
  for (let r = Math.max(0, startIdx - 5); r <= Math.min(startIdx + 5, aoa.length - 1); r++) {
    const row = Array.isArray(aoa[r]) ? aoa[r] : [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] == null) continue;
      if (String(row[c]).trim().toLowerCase().includes(field.value_header.toLowerCase())) {
        headerCol = c;
        break;
      }
    }
    if (headerCol !== null) break;
  }
  if (headerCol === null) throw new Error(`Field "${field.name}": value_header "${field.value_header}" not found near section start`);

  // Resolve target column from letter (e.g. "X" → 23)
  const targetCol = Utils.parseCellAddress(field.target_col + "1").col;

  for (let r = startIdx + 1; r < stopIdx; r++) {
    const row = aoa[r];
    if (!Array.isArray(row)) continue;
    if (Utils.shouldSkipRow(row, SKIP_KEYWORDS, 1)) continue;

    const raw = row[headerCol];
    if (raw == null || Utils.valOrEmpty(raw) === "") continue;

    let val = raw;
    if (field.number_format === "eu") {
      const num = Utils.parseSingleEuroNumber(raw);
      if (Number.isFinite(num)) val = num;
    }

    while (row.length <= targetCol) row.push(undefined);
    row[targetCol] = val;

    if (headerCol !== targetCol) row[headerCol] = "";
  }
}

function applyLabeledInSectionField(aoa, field) {
  const startIdx = Utils.findRowIndexByText(aoa, field.section_start, 1);
  if (startIdx === -1) throw new Error(`Field "${field.name}": section_start "${field.section_start}" not found`);

  const stopIdx = Utils.findRowIndexByText(aoa, field.section_stop, 1, startIdx + 1);
  if (stopIdx === -1) throw new Error(`Field "${field.name}": section_stop "${field.section_stop}" not found`);

  // Find value_header column anywhere in the sheet (headers may be in a shared header row)
  const headerCell = Utils.findLabelCell(aoa, field.value_header);
  if (!headerCell) throw new Error(`Field "${field.name}": value_header "${field.value_header}" not found in sheet`);
  const headerCol = headerCell.col;

  // Find labeled row within the section
  const labelIdx = Utils.findRowIndexByText(aoa, field.label, 1, startIdx + 1);
  if (labelIdx === -1 || labelIdx >= stopIdx) throw new Error(`Field "${field.name}": label "${field.label}" not found in section "${field.section_start}"`);

  const targetCol = Utils.parseCellAddress(field.target_col + "1").col;
  const row = aoa[labelIdx];

  const raw = row[headerCol];
  let val = raw;
  if (field.number_format === "eu") {
    const num = Utils.parseSingleEuroNumber(raw);
    if (Number.isFinite(num)) val = num;
  }

  while (row.length <= targetCol) row.push(undefined);
  row[targetCol] = val;

  // Clear source only if it differs from target
  if (headerCol !== targetCol) row[headerCol] = "";
}

async function applyTemplateFields(aoa, templateName) {
  const template = await loadTemplate(templateName);
  for (const field of template.fields) {
    if (field.kind === "scalar") applyScalarField(aoa, field);
    else if (field.kind === "table_column") applyTableColumnField(aoa, field);
    else if (field.kind === "labeled_in_section") applyLabeledInSectionField(aoa, field);
  }
}

// ===== Main Processing Function =====
async function processFile(file) {
  console.log("process");
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("No se encontró la hoja especificada.");

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, raw: true });

  // Copy all rows (template handlers do all transformations)
  const newAOA = aoa.map(r => Array.isArray(r) ? [...r] : []);

  // Apply template-driven field transformations
  await applyTemplateFields(newAOA, TEMPLATE_NAME);

  // Rebuild sheet
  const newWS = XLSX.utils.aoa_to_sheet(newAOA);
  Utils.ensureRef(newWS, newAOA);

  wb.Sheets[wb.SheetNames[0]] = newWS;
  const outName = file.name.replace(/(\.xlsx?|\.xlsb|\.xlsm)$/i, "") + "_clean.xlsx";
  XLSX.writeFile(wb, outName, { compression: true });

  return outName;
}

// Export for use in other modules
window.Processor = {
  processFile
};
