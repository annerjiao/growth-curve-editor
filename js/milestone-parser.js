/**
 * Parse free-text growth goals into milestone objects.
 * Uses a server proxy (Anthropic Claude) when VITE_MILESTONE_PARSE_URL is set;
 * otherwise local heuristics in the browser.
 */

/** @typedef {{ month: number, value: number, label: string }} ParsedMilestone */

/** @type {string} */
export const MILESTONE_PARSE_URL = import.meta.env.VITE_MILESTONE_PARSE_URL?.trim() ?? "";

/**
 * @param {string} text
 * @param {number} totalMonths
 * @returns {ParsedMilestone[]}
 */
export function parseMilestoneTextLocal(text, totalMonths) {
  if (!text?.trim()) return [];

  const segments = text
    .split(/\n+|;\s*|,\s*(?=[^,]*(?:month|m\d|year|launch|by|at|@|\d+k|\d+m\b))/i)
    .flatMap((s) => s.split(/\s+and\s+then\s+|\s+then\s+/i))
    .map((s) => s.trim())
    .filter(Boolean);

  /** @type {ParsedMilestone[]} */
  const found = [];

  for (const segment of segments.length ? segments : [text.trim()]) {
    const parsed = parseSegment(segment, totalMonths);
    if (parsed) found.push(parsed);
  }

  return normalizeParsed(found, totalMonths);
}

/**
 * @param {string} segment
 * @param {number} totalMonths
 * @returns {ParsedMilestone | null}
 */
function parseSegment(segment, totalMonths) {
  const value = parseValue(segment);
  const month = parseMonth(segment, totalMonths);
  if (value == null && month == null) return null;

  let label = "";
  const labelMatch = segment.match(/^([^:–—-]+?)[:–—-]\s*/);
  if (labelMatch) {
    label = labelMatch[1].trim();
  } else {
    const launch = segment.match(/\b(launch|beta|series\s+[a-z]|pmf|ga)\b/i);
    if (launch) label = launch[1];
  }

  return {
    month: month ?? totalMonths,
    value: value ?? 0,
    label: cleanLabel(label),
  };
}

/** @param {string} s */
function parseValue(s) {
  const withoutMonthNoise = s.replace(/\b(?:month|mo|m)\s*\d{1,3}\b/gi, " ");

  const patterns = [
    /\b(\d+(?:\.\d+)?)\s*[kK]\b/,
    /\b(\d+(?:\.\d+)?)\s*[mM]\b/,
    /\b(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s+(?:users?|signups?|customers?|subs(?:criber)?s?)\b/i,
    /\b(?:to|reach|hit|with|about|around|~)\s*(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\b/i,
  ];

  for (const re of patterns) {
    const m = withoutMonthNoise.match(re);
    if (!m) continue;
    let n = Number(String(m[1]).replace(/,/g, ""));
    if (/[kK]/.test(m[0])) n *= 1000;
    if (/\b\d+(?:\.\d+)?\s*[mM]\b/.test(m[0])) n *= 1_000_000;
    if (Number.isFinite(n)) return Math.round(n);
  }
  return null;
}

/** @param {string} s @param {number} totalMonths */
function parseMonth(s, totalMonths) {
  const lower = s.toLowerCase();

  const explicit = [
    ...s.matchAll(/\b(?:month|mo|m)\s*(\d{1,3})\b/gi),
    ...s.matchAll(/\bM(\d{1,3})\b/g),
    ...s.matchAll(/\b@?\s*(\d{1,3})\s*(?:months?|mos?)\b/gi),
  ];
  for (const m of explicit) {
    const n = Number(m[1]);
    if (n >= 1 && n <= totalMonths) return n;
  }

  const yearMatch = lower.match(/\b(?:year|yr)\s*(\d)\b|\bend\s+of\s+year\s*(\d)\b|\by(\d)\b/);
  if (yearMatch) {
    const y = Number(yearMatch[1] || yearMatch[2] || yearMatch[3]);
    if (y >= 1) return Math.min(totalMonths, y * 12);
  }

  const quarter = lower.match(/\bq([1-4])\b|\bquarter\s*([1-4])\b/);
  if (quarter) {
    const q = Number(quarter[1] || quarter[2]);
    return Math.min(totalMonths, q * 3);
  }

  if (/\b(launch|start|begin|day\s*one)\b/i.test(s)) {
    const atLaunch = s.match(/\b(?:month|m)\s*(\d{1,3})\b/i);
    return atLaunch ? Number(atLaunch[1]) : Math.max(1, Math.round(totalMonths * 0.15));
  }

  if (/\b(end|final|horizon|by\s+month)\b/i.test(s)) return totalMonths;

  return null;
}

/** @param {string} label */
function cleanLabel(label) {
  return label
    .replace(/\b(\d+(?:\.\d+)?\s*[kKmM]?|\d{1,3}(?:,\d{3})+)\b/g, "")
    .replace(/\b(month|m\d+|by|at|with|reach|hit|users?|signups?)\b/gi, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 40);
}

/**
 * @param {ParsedMilestone[]} items
 * @param {number} totalMonths
 */
function normalizeParsed(items, totalMonths) {
  if (!items.length) return [];

  const sorted = [...items]
    .map((m) => ({
      month: Math.min(totalMonths, Math.max(1, Math.round(m.month))),
      value: Math.max(0, Math.round(m.value)),
      label: m.label || "",
    }))
    .sort((a, b) => a.month - b.month);

  /** @type {ParsedMilestone[]} */
  const deduped = [];
  for (const m of sorted) {
    const prev = deduped.at(-1);
    if (prev && prev.month === m.month) {
      prev.value = Math.max(prev.value, m.value);
      if (m.label) prev.label = m.label;
    } else {
      deduped.push({ ...m });
    }
  }

  for (let i = 1; i < deduped.length; i++) {
    if (deduped[i].value < deduped[i - 1].value) {
      deduped[i].value = deduped[i - 1].value;
    }
  }

  return deduped.slice(0, 6);
}

/**
 * @param {string} text
 * @param {number} totalMonths
 * @param {string} apiUrl
 * @returns {Promise<ParsedMilestone[]>}
 */
async function parseMilestoneTextRemote(text, totalMonths, apiUrl) {
  const res = await fetch(apiUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text, totalMonths }),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `Parse API error ${res.status}`);
  }

  const list = data.milestones ?? [];
  if (!Array.isArray(list) || !list.length) {
    throw new Error("AI returned no milestones");
  }

  return normalizeParsed(
    list.map((m) => ({
      month: Number(m.month),
      value: Number(m.value),
      label: String(m.label ?? "").trim(),
    })),
    totalMonths
  );
}

/**
 * @param {string} text
 * @param {number} totalMonths
 * @param {{ apiUrl?: string }} [options]
 * @returns {Promise<{ milestones: ParsedMilestone[], source: 'ai' | 'local' }>}
 */
export async function parseMilestoneText(text, totalMonths, options = {}) {
  const trimmed = text.trim();
  if (!trimmed) {
    return { milestones: [], source: "local" };
  }

  const apiUrl = (options.apiUrl ?? MILESTONE_PARSE_URL).trim();
  if (apiUrl) {
    try {
      const milestones = await parseMilestoneTextRemote(trimmed, totalMonths, apiUrl);
      if (milestones.length) return { milestones, source: "ai" };
    } catch (err) {
      console.warn("AI milestone parse failed, using local parser:", err);
    }
  }

  return {
    milestones: parseMilestoneTextLocal(trimmed, totalMonths),
    source: "local",
  };
}
