import * as XLSX from "xlsx";
import { colNumberToLetters, parseRangeRef } from "./xlsx-io.js";

/** @typedef {{ row: number, startCol: number, endCol: number, count: number }} MonthHeader */

/** @typedef {{ sheet: string, rangeRef: string, label: string, metricName: string, scenarioLabel: string, metricLabel: string, score: number, confidence: "high" | "medium" | "low", reasons: string[], values: number[], formulaPct: number, hasFormulas: boolean, periodCount: number }} DriverCandidate */

const METRIC_PATTERNS = [
  { re: /\bnew\s+user\s+signups?\b/i, label: "New user signups", weight: 16 },
  { re: /\b(new\s+)?signups?\s*(\/|per)\s*(month|mo)\b/i, label: "Signups", weight: 15 },
  { re: /\b(new\s+)?signups?\b/i, label: "Signups", weight: 14 },
  { re: /\busers?\s*(\/|per)\s*(month|mo)\b/i, label: "Users", weight: 15 },
  { re: /\busers?\b/i, label: "Users", weight: 13 },
  { re: /\b(customers?|clients?)\b/i, label: "Customers", weight: 11 },
  { re: /\b(mau|dau|wau)\b/i, label: "Users", weight: 12 },
  { re: /\b(subscribers?|members?)\b/i, label: "Subscribers", weight: 10 },
  { re: /\b(accounts?|seats?)\b/i, label: "Accounts", weight: 9 },
  { re: /\b(installs?|downloads?)\b/i, label: "Installs", weight: 8 },
  { re: /\b(revenue|mrr|arr|gmv|sales)\b/i, label: "Revenue", weight: 6 },
  { re: /\b(units?|orders?|transactions?)\b/i, label: "Units", weight: 5 },
  { re: /\b(growth|driver|input|assumption)\b/i, label: "Driver", weight: 4 },
];

const MONTH_RES = [
  /^m\s*(\d{1,3})$/i,
  /^month\s*(\d{1,3})$/i,
  /^period\s*(\d{1,3})$/i,
  /^p\s*(\d{1,3})$/i,
  /^y\s*(\d)\s*m\s*(\d{1,2})$/i,
];

const MONTH_NAMES =
  /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(?:uary|ch|il|e|ust|ember|t|ober|ember)?$/i;

/** @param {unknown} v */
function cellText(v) {
  if (v == null) return "";
  if (v instanceof Date) return v.toLocaleDateString("en-US", { month: "short" });
  return String(v).trim();
}

/** @param {string} text */
function parseMonthIndex(text) {
  const t = text.trim();
  for (const re of MONTH_RES) {
    const m = t.match(re);
    if (m) {
      if (m[2] != null) return Number(m[2]);
      return Number(m[1]);
    }
  }
  if (MONTH_NAMES.test(t)) return 1;
  if (/^\d{1,2}$/.test(t)) return Number(t);
  return null;
}

/**
 * @param {import('xlsx').WorkSheet} ws
 * @param {number} maxRows
 * @param {number} maxCols
 */
function sheetBounds(ws, maxRows = 200, maxCols = 120) {
  const ref = ws["!ref"];
  if (!ref) return { rows: 0, cols: 0 };
  const range = XLSX.utils.decode_range(ref);
  return {
    rows: Math.min(range.e.r + 1, maxRows),
    cols: Math.min(range.e.c + 1, maxCols),
  };
}

/** @param {import('xlsx').WorkSheet} ws @param {number} r @param {number} c */
function getCell(ws, r, c) {
  return ws[XLSX.utils.encode_cell({ r, c })];
}

/** @param {import('xlsx').WorkSheet} ws @param {number} r @param {number} startCol @param {number} endCol */
function rowLabelText(ws, r, startCol) {
  const parts = [];
  for (let c = 0; c < Math.min(startCol, 6); c++) {
    const cell = getCell(ws, r, c);
    const t = cellText(cell?.w ?? cell?.v);
    if (t) parts.push(t);
  }
  return parts.join(" ").trim();
}

/** First descriptive text cell in a row (section headers). */
function rowHeaderText(ws, r, maxCol = 10) {
  for (let c = 0; c < maxCol; c++) {
    const cell = getCell(ws, r, c);
    const t = cellText(cell?.w ?? cell?.v);
    if (!t || /^\d+$/.test(t)) continue;
    if (t.length >= 4) return t;
  }
  return "";
}

/** @param {string} text */
function isScenarioLike(text) {
  const t = text.trim();
  return (
    /^(active\s+case|\d+[.)]\s*)/i.test(t) ||
    /\b(base|upside|downside|bear|bull)\b.*\bcase\b/i.test(t) ||
    (/^\d+[.)]\s*\w+/i.test(t) && t.length < 50)
  );
}

/** @param {string} text */
function cleanScenarioLabel(text) {
  return text.replace(/^\d+[.)]\s*/, "").replace(/\s+/g, " ").trim();
}

/** Count cells with real numeric values (blank/empty string cells excluded). */
function countNumericCells(ws, r, startCol, endCol) {
  let count = 0;
  for (let c = startCol; c <= endCol; c++) {
    const cell = getCell(ws, r, c);
    if (!cell || cell.v == null || cell.v === "") continue;
    const raw = cell.v;
    const num =
      typeof raw === "number"
        ? raw
        : raw instanceof Date
          ? null
          : Number(String(raw).replace(/[,$%]/g, ""));
    if (Number.isFinite(num)) count++;
  }
  return count;
}

/**
 * Look above the data row for a section metric name (e.g. "New user signups / month").
 * @param {import('xlsx').WorkSheet} ws
 * @param {number} dataRow
 * @param {number} startCol
 */
function findMetricNameAbove(ws, dataRow, startCol) {
  let best = { text: "", score: 0 };

  for (let r = dataRow - 1; r >= Math.max(0, dataRow - 15); r--) {
    const headerText = rowHeaderText(ws, r, Math.max(startCol, 10));
    if (!headerText || headerText.length < 4) continue;

    if (countNumericCells(ws, r, startCol, Math.min(startCol + 11, startCol + 35)) >= 4) {
      continue;
    }

    if (
      isScenarioLike(headerText) &&
      !/\//.test(headerText) &&
      !/per\s+(month|mo|year)/i.test(headerText)
    ) {
      continue;
    }

    let score = scoreLabel(headerText).score;
    if (/\//.test(headerText)) score += 8;
    if (/per\s+(month|mo|year|week|day)/i.test(headerText)) score += 10;
    if (/\b(month|monthly|signups?|users?|revenue|customers?)\b/i.test(headerText)) {
      score += 6;
    }
    if (headerText.length >= 12) score += 3;

    if (score > best.score) best = { text: headerText, score };
  }

  return best.text;
}

/**
 * @param {import('xlsx').WorkSheet} ws
 * @param {number} dataRow
 * @param {number} startCol
 */
function resolveRowLabels(ws, dataRow, startCol) {
  const rawScenario = rowLabelText(ws, dataRow, startCol);
  const scenarioLabel = cleanScenarioLabel(rawScenario);
  const metricName = findMetricNameAbove(ws, dataRow, startCol);
  const combined = [metricName, scenarioLabel].filter(Boolean).join(" · ");
  const label = combined || scenarioLabel || metricName || `Row ${dataRow + 1}`;
  const labelScore = scoreLabel(metricName || scenarioLabel || rawScenario);

  return {
    label,
    metricName,
    scenarioLabel,
    metricLabel: metricName || labelScore.metricLabel || "Users",
    labelScore: labelScore.score,
  };
}

/** @param {import('xlsx').WorkSheet} ws @param {number} r @param {number} maxCols */
function findMonthHeadersInRow(ws, r, maxCols) {
  /** @type {MonthHeader[]} */
  const headers = [];
  let c = 0;
  while (c < maxCols) {
    const cell = getCell(ws, r, c);
    const text = cellText(cell?.w ?? cell?.v);
    const idx = parseMonthIndex(text);
    if (idx == null) {
      c++;
      continue;
    }

    const startCol = c;
    let prev = idx;
    let count = 1;
    c++;

    while (c < maxCols) {
      const nextCell = getCell(ws, r, c);
      const nextText = cellText(nextCell?.w ?? nextCell?.v);
      const nextIdx = parseMonthIndex(nextText);
      if (nextIdx == null) break;
      if (nextIdx <= prev && count > 1) break;
      prev = nextIdx;
      count++;
      c++;
    }

    if (count >= 6) {
      headers.push({ row: r, startCol, endCol: c - 1, count });
    }
  }
  return headers;
}

/** @param {import('xlsx').WorkSheet} ws @param {number} maxRows @param {number} maxCols */
function findAllMonthHeaders(ws, maxRows, maxCols) {
  /** @type {MonthHeader[]} */
  const all = [];
  for (let r = 0; r < maxRows; r++) {
    all.push(...findMonthHeadersInRow(ws, r, maxCols));
  }
  return all;
}

/**
 * @param {import('xlsx').WorkSheet} ws
 * @param {number} r
 * @param {number} startCol
 * @param {number} endCol
 */
function extractNumericRun(ws, r, startCol, endCol) {
  /** @type {number[]} */
  const values = [];
  let formulaCount = 0;
  let numericCount = 0;

  for (let c = startCol; c <= endCol; c++) {
    const cell = getCell(ws, r, c);
    if (cell?.f) formulaCount++;
    const raw = cell?.v;
    const num =
      typeof raw === "number"
        ? raw
        : raw instanceof Date
          ? null
          : Number(String(raw ?? "").replace(/[,$%]/g, ""));
    if (Number.isFinite(num)) {
      values.push(num);
      numericCount++;
    } else {
      values.push(NaN);
    }
  }

  return { values, formulaCount, numericCount };
}

/** @param {number[]} values */
function scoreVariation(values) {
  const nums = values.filter((v) => Number.isFinite(v));
  if (nums.length < 6) return 0;
  const nonZero = nums.filter((v) => v !== 0);
  if (nonZero.length < 3) return 0;

  const min = Math.min(...nums);
  const max = Math.max(...nums);
  if (min === max) return 0;

  const range = max - min;
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  const variance = nums.reduce((s, v) => s + (v - mean) ** 2, 0) / nums.length;
  const cv = mean !== 0 ? Math.sqrt(variance) / Math.abs(mean) : range > 0 ? 1 : 0;

  let score = 0;
  if (range > 0) score += 8;
  if (cv > 0.15) score += 6;
  if (nonZero.length / nums.length > 0.5) score += 6;
  return Math.min(20, score);
}

/**
 * Hardcoded driver row with flat/unfilled values (common in starter templates).
 * @param {number[]} values
 * @param {number} formulaCount
 * @param {number} labelScore
 */
function isPlaceholderDriverRow(values, formulaCount, labelScore) {
  if (formulaCount > 0 || labelScore < 8) return false;
  const nums = values.filter((v) => Number.isFinite(v));
  if (nums.length < 6) return false;
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  if (min === max) return true;
  const spread = max - min;
  const ref = Math.max(Math.abs(max), Math.abs(min), 1);
  return spread / ref < 0.02;
}

/** @param {number[]} values */
function scoreGrowthShape(values) {
  const nums = values.filter((v) => Number.isFinite(v));
  if (nums.length < 6) return 0;

  if (nums.some((v) => v < 0)) return -5;

  const diffs = nums.slice(1).map((v, i) => v - nums[i]);
  const nonDecreasing = diffs.filter((d) => d >= -0.01).length / diffs.length;
  const nonNegativeFlow = nums.filter((v) => v >= 0).length / nums.length;

  let score = 0;
  if (nonDecreasing > 0.75) score += 10;
  else if (nonNegativeFlow > 0.9) score += 6;

  const totalGrowth = nums.at(-1) - nums[0];
  if (totalGrowth > 0) score += 5;

  return Math.min(15, score);
}

/** @param {string} labelText */
function scoreLabel(labelText) {
  if (!labelText) return { score: 0, metricLabel: "" };
  let best = { score: 0, metricLabel: "" };
  for (const { re, label, weight } of METRIC_PATTERNS) {
    if (re.test(labelText) && weight > best.score) {
      best = { score: weight, metricLabel: label };
    }
  }
  return best;
}

/** @param {number} formulaCount @param {number} numericCount */
function scoreInputCells(formulaCount, numericCount) {
  if (numericCount === 0) return -20;
  const pct = formulaCount / numericCount;
  if (pct === 0) return 20;
  if (pct < 0.2) return 8;
  if (pct < 0.5) return -5;
  return -15;
}

/** @param {number} count */
function scorePeriodCount(count) {
  if (count >= 12 && count <= 60) return 10;
  if (count >= 6 && count <= 84) return 6;
  if (count >= 3) return 2;
  return 0;
}

/**
 * @param {import('xlsx').WorkBook} wb
 * @param {string} [preferredSheet]
 * @returns {DriverCandidate[]}
 */
export function detectDriverCandidates(wb, preferredSheet) {
  /** @type {DriverCandidate[]} */
  const candidates = [];
  const sheets = preferredSheet ? [preferredSheet] : wb.SheetNames;

  for (const sheetName of sheets) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const { rows, cols } = sheetBounds(ws);
    const headers = findAllMonthHeaders(ws, rows, cols);

    /** @type {{ r: number, startCol: number, endCol: number }[]} */
    const spans = headers.map((h) => ({
      r: h.row,
      startCol: h.startCol,
      endCol: h.endCol,
    }));

    if (spans.length === 0) {
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols - 5; c++) {
          const run = extractNumericRun(ws, r, c, Math.min(c + 35, cols - 1));
          const valid = run.values.filter((v) => Number.isFinite(v));
          if (valid.length >= 6 && scoreVariation(valid) >= 8) {
            spans.push({ r, startCol: c, endCol: c + valid.length - 1 });
          }
        }
      }
    }

    const seen = new Set();

    for (const header of headers) {
      for (let r = header.row + 1; r <= Math.min(header.row + 25, rows - 1); r++) {
        const key = `${r}:${header.startCol}:${header.endCol}`;
        if (seen.has(key)) continue;
        seen.add(key);

        const { values, formulaCount, numericCount } = extractNumericRun(
          ws,
          r,
          header.startCol,
          header.endCol
        );
        const valid = values.filter((v) => Number.isFinite(v));
        if (valid.length < 6) continue;

        const resolved = resolveRowLabels(ws, r, header.startCol);
        let varScore = scoreVariation(valid);
        const isPlaceholder =
          varScore < 4 && isPlaceholderDriverRow(valid, formulaCount, resolved.labelScore);
        if (varScore < 4 && !isPlaceholder) continue;
        if (isPlaceholder) varScore = 5;

        const shapeScore = scoreGrowthShape(valid);
        const inputScore = scoreInputCells(formulaCount, numericCount);
        const periodScore = scorePeriodCount(valid.length);
        const headerScore = 25;

        const total =
          resolved.labelScore +
          varScore +
          shapeScore +
          inputScore +
          periodScore +
          headerScore;

        /** @type {string[]} */
        const reasons = [];
        if (resolved.metricName) {
          reasons.push(`Metric: "${resolved.metricName.slice(0, 50)}"`);
        }
        if (resolved.scenarioLabel) {
          reasons.push(`Scenario: "${resolved.scenarioLabel.slice(0, 40)}"`);
        } else if (resolved.labelScore > 0) {
          reasons.push(`Label: "${resolved.label.slice(0, 40)}"`);
        }
        reasons.push(`${valid.length} monthly periods detected`);
        if (formulaCount === 0) reasons.push("Hardcoded values (safe to edit)");
        else if (formulaCount / numericCount > 0.5) {
          reasons.push("Mostly formulas — may break if overwritten");
        } else {
          reasons.push("Mix of values and formulas");
        }
        if (shapeScore >= 10) reasons.push("Growth pattern detected");
        if (isPlaceholder) reasons.push("Flat placeholder values — ready to edit");

        const rangeRef = `${colNumberToLetters(header.startCol + 1)}${r + 1}:${colNumberToLetters(header.endCol + 1)}${r + 1}`;

        candidates.push({
          sheet: sheetName,
          rangeRef,
          label: resolved.label,
          metricName: resolved.metricName,
          scenarioLabel: resolved.scenarioLabel,
          metricLabel: resolved.metricLabel,
          score: total,
          confidence: total >= 55 ? "high" : total >= 35 ? "medium" : "low",
          reasons,
          values: valid,
          formulaPct: numericCount ? formulaCount / numericCount : 0,
          hasFormulas: formulaCount > 0,
          periodCount: valid.length,
        });
      }
    }

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols - 5; c++) {
        const endCol = Math.min(c + 59, cols - 1);
        const { values, formulaCount, numericCount } = extractNumericRun(ws, r, c, endCol);
        const firstNaN = values.findIndex((v) => !Number.isFinite(v));
        const runEnd = firstNaN === -1 ? endCol : c + firstNaN - 1;
        if (runEnd < c + 5) continue;

        const key = `${r}:${c}:${runEnd}`;
        if (seen.has(key)) continue;

        const slice = values.slice(0, runEnd - c + 1).filter((v) => Number.isFinite(v));
        if (slice.length < 6) continue;

        const varScore = scoreVariation(slice);
        if (varScore < 6) continue;

        seen.add(key);
        const resolved = resolveRowLabels(ws, r, c);
        const isPlaceholder =
          varScore < 6 && isPlaceholderDriverRow(slice, formulaCount, resolved.labelScore);

        if (resolved.labelScore === 0 && varScore < 10 && !isPlaceholder) continue;

        const effectiveVarScore = isPlaceholder ? 5 : varScore;

        const shapeScore = scoreGrowthShape(slice);
        const inputScore = scoreInputCells(formulaCount, numericCount);
        const periodScore = scorePeriodCount(slice.length);

        const total =
          resolved.labelScore + effectiveVarScore + shapeScore + inputScore + periodScore;

        if (total < 20) continue;

        const rangeRef = `${colNumberToLetters(c + 1)}${r + 1}:${colNumberToLetters(runEnd + 1)}${r + 1}`;

        /** @type {string[]} */
        const reasons = [];
        if (resolved.metricName) {
          reasons.push(`Metric: "${resolved.metricName.slice(0, 50)}"`);
        }
        if (resolved.scenarioLabel) {
          reasons.push(`Scenario: "${resolved.scenarioLabel.slice(0, 40)}"`);
        } else if (resolved.label) {
          reasons.push(`Label: "${resolved.label.slice(0, 40)}"`);
        }
        reasons.push(`${slice.length} changing values in a row`);
        if (isPlaceholder) reasons.push("Flat placeholder values — ready to edit");
        if (formulaCount === 0) reasons.push("Hardcoded values (safe to edit)");
        else if (formulaCount / numericCount > 0.5) {
          reasons.push("Mostly formulas — may break if overwritten");
        }

        candidates.push({
          sheet: sheetName,
          rangeRef,
          label: resolved.label,
          metricName: resolved.metricName,
          scenarioLabel: resolved.scenarioLabel,
          metricLabel: resolved.metricLabel,
          score: total,
          confidence: total >= 50 ? "high" : total >= 30 ? "medium" : "low",
          reasons,
          values: slice,
          formulaPct: numericCount ? formulaCount / numericCount : 0,
          hasFormulas: formulaCount > 0,
          periodCount: slice.length,
        });
      }
    }
  }

  candidates.sort((a, b) => b.score - a.score);

  const deduped = [];
  const seenRanges = new Set();
  /** @type {{ sheet: string, row: number, startCol: number, endCol: number }[]} */
  const acceptedSpans = [];

  for (const c of candidates) {
    const key = `${c.sheet}!${c.rangeRef}`;
    if (seenRanges.has(key)) continue;

    const { r0, c0, c1 } = parseRangeRef(c.rangeRef);
    const overlaps = acceptedSpans.some(
      (s) =>
        s.sheet === c.sheet &&
        s.row === r0 &&
        !(c1 < s.startCol || c0 > s.endCol)
    );
    if (overlaps) continue;

    seenRanges.add(key);
    acceptedSpans.push({ sheet: c.sheet, row: r0, startCol: c0, endCol: c1 });
    deduped.push(c);
    if (deduped.length >= 8) break;
  }

  return deduped;
}

/**
 * @param {import('xlsx').WorkBook} wb
 * @param {string} sheetName
 * @param {string} rangeRef
 */
export function analyzeMappedRange(wb, sheetName, rangeRef) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  const { r0, c0, c1 } = parseRangeRef(rangeRef);
  const resolved = resolveRowLabels(ws, r0, c0);
  const { values, formulaCount, numericCount } = extractNumericRun(ws, r0, c0, c1);
  const valid = values.filter((v) => Number.isFinite(v));

  /** @type {string[]} */
  const warnings = [];
  const formulaPct = numericCount ? formulaCount / numericCount : 0;

  if (formulaPct > 0.5) {
    warnings.push(
      "This range is mostly formulas. Overwriting will replace calculations and likely break your model. Map the input/driver row instead."
    );
  } else if (formulaPct > 0) {
    warnings.push(
      "Some cells contain formulas. Those cells will be overwritten with hardcoded numbers on export."
    );
  }

  if (valid.length < 2) {
    warnings.push("Range has fewer than 2 numeric values.");
  }

  const varScore = scoreVariation(valid);
  if (valid.length >= 2 && varScore < 2) {
    warnings.push("Values barely change across periods — double-check this is the right driver row.");
  }

  return {
    label: resolved.label,
    metricName: resolved.metricName,
    scenarioLabel: resolved.scenarioLabel,
    metricLabel: resolved.metricLabel,
    formulaCount,
    numericCount,
    formulaPct,
    hasFormulas: formulaCount > 0,
    warnings,
    values: valid,
  };
}

/**
 * Pick the sheet most likely to contain a growth driver row.
 * Prefers sheets with high-confidence matches, then best top score.
 * @param {import('xlsx').WorkBook} wb
 * @returns {string | null}
 */
export function findSuggestedSheet(wb) {
  /** @type {{ sheet: string, topScore: number, hasHigh: boolean }[]} */
  const ranked = [];

  for (const sheetName of wb.SheetNames) {
    const candidates = detectDriverCandidates(wb, sheetName);
    if (!candidates.length) continue;
    ranked.push({
      sheet: sheetName,
      topScore: candidates[0].score,
      hasHigh: candidates.some((c) => c.confidence === "high"),
    });
  }

  if (!ranked.length) return null;

  ranked.sort((a, b) => {
    if (a.hasHigh !== b.hasHigh) return a.hasHigh ? -1 : 1;
    return b.topScore - a.topScore;
  });

  return ranked[0].sheet;
}
