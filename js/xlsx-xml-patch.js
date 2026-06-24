import JSZip from "jszip";
import { colLettersToNumber, colNumberToLetters, parseRangeRef } from "./xlsx-io.js";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** @param {string} cellRef e.g. C8 */
function colIndexFromCellRef(cellRef) {
  const m = cellRef.match(/^([A-Za-z]+)/);
  return m ? colLettersToNumber(m[1]) - 1 : 0;
}

/** @param {number} r0 @param {number} c0 */
function cellRefFromCoords(r0, c0) {
  return `${colNumberToLetters(c0 + 1)}${r0 + 1}`;
}

/**
 * @param {JSZip} zip
 * @returns {Promise<Map<string, string>>}
 */
async function buildSheetPathMap(zip) {
  const workbookFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookFile || !relsFile) {
    throw new Error("Not a valid .xlsx workbook.");
  }

  const parser = new DOMParser();
  const wbDoc = parser.parseFromString(await workbookFile.async("string"), "application/xml");
  const relsDoc = parser.parseFromString(await relsFile.async("string"), "application/xml");

  /** @type {Map<string, string>} */
  const relMap = new Map();
  for (const rel of relsDoc.getElementsByTagName("Relationship")) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) relMap.set(id, target);
  }

  /** @type {Map<string, string>} */
  const sheetPaths = new Map();
  for (const sheet of wbDoc.getElementsByTagName("sheet")) {
    const name = sheet.getAttribute("name");
    const rid =
      sheet.getAttribute("r:id") ||
      sheet.getAttributeNS(REL_NS, "id");
    const target = rid ? relMap.get(rid) : null;
    if (!name || !target) continue;
    const normalized = target.replace(/^\//, "");
    sheetPaths.set(name, normalized.startsWith("xl/") ? normalized : `xl/${normalized}`);
  }

  return sheetPaths;
}

/**
 * @param {Document} doc
 * @param {Element} cellEl
 * @param {number} value
 */
function patchCellElement(doc, cellEl, value) {
  for (const tag of ["f", "is"]) {
    Array.from(cellEl.getElementsByTagName(tag)).forEach((el) => el.remove());
  }
  cellEl.removeAttribute("t");

  let vEl = cellEl.getElementsByTagName("v")[0];
  if (!vEl) {
    vEl = doc.createElementNS(MAIN_NS, "v");
    cellEl.appendChild(vEl);
  }
  vEl.textContent = String(Math.max(0, Math.round(value)));
}

/**
 * @param {Document} doc
 * @param {string} cellRef
 * @param {number} value
 */
function createNumericCell(doc, cellRef, value) {
  const cell = doc.createElementNS(MAIN_NS, "c");
  cell.setAttribute("r", cellRef);
  const vEl = doc.createElementNS(MAIN_NS, "v");
  vEl.textContent = String(Math.max(0, Math.round(value)));
  cell.appendChild(vEl);
  return cell;
}

/**
 * @param {Element} row
 * @param {string} cellRef
 */
function findCellInRow(row, cellRef) {
  for (const cell of row.getElementsByTagName("c")) {
    if (cell.getAttribute("r") === cellRef) return cell;
  }
  return null;
}

/**
 * @param {Document} doc
 * @param {Element} row
 * @param {Element} cellEl
 * @param {number} colIndex
 */
function insertCellInRow(row, cellEl, colIndex) {
  const cells = row.getElementsByTagName("c");
  for (let i = 0; i < cells.length; i++) {
    const existing = cells[i];
    const ref = existing.getAttribute("r");
    if (!ref) continue;
    if (colIndex < colIndexFromCellRef(ref)) {
      row.insertBefore(cellEl, existing);
      return;
    }
  }
  row.appendChild(cellEl);
}

/**
 * @param {Document} doc
 * @param {Element} sheetData
 * @param {number} rowNum 1-based
 */
function findOrCreateRow(doc, sheetData, rowNum) {
  for (const row of sheetData.getElementsByTagName("row")) {
    if (Number(row.getAttribute("r")) === rowNum) return row;
  }

  const row = doc.createElementNS(MAIN_NS, "row");
  row.setAttribute("r", String(rowNum));

  const rows = sheetData.getElementsByTagName("row");
  for (let i = 0; i < rows.length; i++) {
    const existingNum = Number(rows[i].getAttribute("r"));
    if (rowNum < existingNum) {
      sheetData.insertBefore(row, rows[i]);
      return row;
    }
  }

  sheetData.appendChild(row);
  return row;
}

/**
 * Patch numeric values in worksheet XML, preserving styles (`s`) and all other markup.
 * @param {string} xml
 * @param {string} rangeRef
 * @param {number[]} values
 */
export function patchSheetXml(xml, rangeRef, values) {
  const { r0, c0, c1 } = parseRangeRef(rangeRef);
  const rowNum = r0 + 1;
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");

  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("Failed to parse worksheet XML.");
  }

  const sheetData = doc.getElementsByTagName("sheetData")[0];
  if (!sheetData) throw new Error("Worksheet has no sheetData.");

  const row = findOrCreateRow(doc, sheetData, rowNum);
  const width = c1 - c0 + 1;

  for (let i = 0; i < width; i++) {
    const col = c0 + i;
    const cellRef = cellRefFromCoords(r0, col);
    const val = values[i] ?? 0;
    const existing = findCellInRow(row, cellRef);

    if (existing) {
      patchCellElement(doc, existing, val);
    } else {
      insertCellInRow(row, createNumericCell(doc, cellRef, val), col);
    }
  }

  return new XMLSerializer().serializeToString(doc);
}

/**
 * Patch the original .xlsx bytes in-place (ZIP/XML level). Preserves formatting.
 * @param {ArrayBuffer} originalBuffer
 * @param {{ sheet: string, rangeRef: string, values: number[] }[]} updates
 */
export async function patchWorkbookBuffer(originalBuffer, updates) {
  const zip = await JSZip.loadAsync(originalBuffer);
  const sheetPaths = await buildSheetPathMap(zip);

  for (const { sheet, rangeRef, values } of updates) {
    const path = sheetPaths.get(sheet);
    if (!path) throw new Error(`Sheet not found: ${sheet}`);

    const file = zip.file(path);
    if (!file) throw new Error(`Worksheet file missing: ${path}`);

    const patched = patchSheetXml(await file.async("string"), rangeRef, values);
    zip.file(path, patched);
  }

  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}
