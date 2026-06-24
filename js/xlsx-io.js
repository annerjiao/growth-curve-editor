import * as XLSX from "xlsx";
import { patchWorkbookBuffer } from "./xlsx-xml-patch.js";

/** Preserve formulae, number formats, and sheet structure when reading. */
const READ_OPTS = {
  type: "array",
  cellDates: true,
  cellFormula: true,
  cellNF: true,
  sheetStubs: true,
};

/** Write options — fallback export only (rebuilds workbook, drops formatting). */
const WRITE_OPTS = {
  bookType: "xlsx",
  type: "array",
  cellDates: true,
};

/** @param {string} letters */
export function colLettersToNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

/** @param {number} n 1-based */
export function colNumberToLetters(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Parse A1 or A1:B2 into { r0,c0,r1,c1 } (0-based, inclusive).
 * @param {string} ref
 */
export function parseRangeRef(ref) {
  const m = ref.trim().match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
  if (!m) throw new Error(`Invalid range: ${ref}`);
  const c0 = colLettersToNumber(m[1]) - 1;
  const r0 = Number(m[2]) - 1;
  const c1 = m[3] ? colLettersToNumber(m[3]) - 1 : c0;
  const r1 = m[4] ? Number(m[4]) - 1 : r0;
  return {
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
}

/**
 * @param {ArrayBuffer} buffer
 */
export function loadWorkbook(buffer) {
  return XLSX.read(buffer, READ_OPTS);
}

/** @param {import('xlsx').WorkBook} wb */
export function listSheetNames(wb) {
  return wb.SheetNames.slice();
}

/**
 * Read numeric flow values from a horizontal range on one row.
 * @param {import('xlsx').WorkBook} wb
 * @param {string} sheetName
 * @param {string} rangeRef e.g. C8:AN8
 */
export function readFlowRange(wb, sheetName, rangeRef) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  const { r0, c0, c1 } = parseRangeRef(rangeRef);
  const values = [];
  for (let c = c0; c <= c1; c++) {
    const addr = XLSX.utils.encode_cell({ r: r0, c });
    const cell = ws[addr];
    const raw = cell?.v;
    values.push(typeof raw === "number" ? raw : Number(raw) || 0);
  }
  return {
    values,
    row: r0 + 1,
    startCol: colNumberToLetters(c0 + 1),
    endCol: colNumberToLetters(c1 + 1),
    rangeRef,
  };
}

/**
 * Patch only the mapped driver cells — leave every other cell untouched.
 * @param {import('xlsx').WorkSheet} ws
 * @param {number} r0
 * @param {number} c0
 * @param {number} c1
 * @param {number[]} values
 */
function patchFlowCells(ws, r0, c0, c1, values) {
  const width = c1 - c0 + 1;
  for (let i = 0; i < width; i++) {
    const addr = XLSX.utils.encode_cell({ r: r0, c: c0 + i });
    const val = Math.max(0, Math.round(values[i] ?? 0));
    const existing = ws[addr];

    if (existing) {
      existing.t = "n";
      existing.v = val;
      delete existing.f;
      delete existing.F;
      delete existing.w;
    } else {
      ws[addr] = { t: "n", v: val };
    }
  }

  if (!ws["!ref"]) {
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: r0, c: c0 },
      e: { r: r0, c: c1 },
    });
  } else {
    const cur = XLSX.utils.decode_range(ws["!ref"]);
    cur.e.r = Math.max(cur.e.r, r0);
    cur.e.c = Math.max(cur.e.c, c1);
    ws["!ref"] = XLSX.utils.encode_range(cur);
  }
}

/**
 * Write flow values back to workbook (same range width).
 * @param {import('xlsx').WorkBook} wb
 * @param {string} sheetName
 * @param {string} rangeRef
 * @param {number[]} values
 */
export function writeFlowRange(wb, sheetName, rangeRef, values) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  const { r0, c0, c1 } = parseRangeRef(rangeRef);
  patchFlowCells(ws, r0, c0, c1, values);
}

/**
 * SheetJS drops formula cells that have no cached value on write.
 * Ensure every formula cell has a value so Excel shows content after open.
 * @param {import('xlsx').WorkBook} wb
 */
export function prepareWorkbookForWrite(wb) {
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    for (const key of Object.keys(ws)) {
      if (key[0] === "!") continue;
      const cell = ws[key];
      if (!cell?.f || cell.v !== undefined) continue;
      cell.t = cell.t || "n";
      cell.v = 0;
    }
  }
}

/**
 * Re-read the original upload and patch only the mapped ranges.
 * @param {ArrayBuffer} originalBuffer
 * @param {{ sheet: string, rangeRef: string, values: number[] }[]} updates
 */
export function buildUpdatedWorkbook(originalBuffer, updates) {
  const wb = loadWorkbook(originalBuffer);
  for (const { sheet, rangeRef, values } of updates) {
    writeFlowRange(wb, sheet, rangeRef, values);
  }
  prepareWorkbookForWrite(wb);
  return wb;
}

/**
 * @param {ArrayBuffer} buffer
 * @param {string} [filename]
 */
function downloadArrayBuffer(buffer, filename = "model-updated.xlsx") {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/**
 * @param {import('xlsx').WorkBook} wb
 * @param {string} [filename]
 */
export function downloadWorkbook(wb, filename = "model-updated.xlsx") {
  const out = XLSX.write(wb, WRITE_OPTS);
  downloadArrayBuffer(out, filename);
}

/**
 * Export from the original file bytes — XML patch preserves formatting.
 * @param {ArrayBuffer} originalBuffer
 * @param {{ sheet: string, rangeRef: string, values: number[] }[]} updates
 * @param {string} [filename]
 */
export async function downloadUpdatedWorkbook(originalBuffer, updates, filename) {
  try {
    const out = await patchWorkbookBuffer(originalBuffer, updates);
    downloadArrayBuffer(out, filename);
  } catch (err) {
    console.warn("XML patch failed, falling back to SheetJS export:", err);
    const wb = buildUpdatedWorkbook(originalBuffer, updates);
    downloadWorkbook(wb, filename);
  }
}

/**
 * @param {import('xlsx').WorkBook} wb
 * @param {string} sheetName
 * @param {number} maxRows
 * @param {number} maxCols
 */
export function getSheetPreview(wb, sheetName, maxRows = 12, maxCols = 16) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return { rows: [] };
  const ref = ws["!ref"];
  if (!ref) return { rows: [] };
  const range = XLSX.utils.decode_range(ref);
  const rows = [];
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + maxRows - 1); r++) {
    const row = [];
    for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + maxCols - 1); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell?.w ?? cell?.v ?? "");
    }
    rows.push(row);
  }
  return {
    rows,
    startRow: range.s.r + 1,
    startCol: colNumberToLetters(range.s.c + 1),
  };
}
