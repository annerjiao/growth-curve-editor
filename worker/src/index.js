/**
 * Cloudflare Worker — proxies milestone parsing to Anthropic.
 * Keeps ANTHROPIC_API_KEY server-side; never ship it to the browser.
 */

const SYSTEM_PROMPT = (totalMonths) =>
  `You extract growth milestones for a financial model spreadsheet. The planning horizon is ${totalMonths} months (month 1 through month ${totalMonths}).

Return ONLY valid JSON with this exact shape:
{"milestones":[{"label":string,"month":number,"value":number}]}

Rules:
- month is 1-indexed (1..${totalMonths})
- value is cumulative total (users, signups, or customers) by that month
- at most 6 milestones
- sorted by month ascending
- values must be non-decreasing
- label is a short human name (e.g. "Launch", "Series A") or empty string
- interpret informal phrases like "5k", "end of year 1" (month 12), "year 2" (month 24)`;

/** @param {string} origin @param {string} allowedCsv */
function isAllowedOrigin(origin, allowedCsv) {
  if (!origin) return false;
  const allowed = allowedCsv.split(",").map((s) => s.trim()).filter(Boolean);
  return allowed.some((a) => {
    if (a.endsWith("*")) return origin.startsWith(a.slice(0, -1));
    return origin === a;
  });
}

/** @param {string} origin @param {string} allowedCsv */
function corsHeaders(origin, allowedCsv) {
  if (!isAllowedOrigin(origin, allowedCsv)) return {};
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
  };
}

/** @param {unknown} body @param {number} totalMonths */
function normalizeMilestones(body, totalMonths) {
  const list = body?.milestones ?? body;
  if (!Array.isArray(list)) throw new Error("Invalid milestones array");

  const sorted = list
    .map((m) => ({
      month: Math.min(totalMonths, Math.max(1, Math.round(Number(m.month) || 1))),
      value: Math.max(0, Math.round(Number(m.value) || 0)),
      label: String(m.label ?? "").trim().slice(0, 40),
    }))
    .sort((a, b) => a.month - b.month);

  /** @type {typeof sorted} */
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

/** @param {string} text @param {number} totalMonths @param {Env} env */
async function parseWithAnthropic(text, totalMonths, env) {
  const apiKey = env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY not configured on worker");

  const model = env.ANTHROPIC_MODEL || "claude-sonnet-4-5";

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model,
      max_tokens: 1024,
      temperature: 0.2,
      system: SYSTEM_PROMPT(totalMonths),
      messages: [{ role: "user", content: text }],
    }),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Anthropic ${res.status}: ${err.slice(0, 300)}`);
  }

  const data = await res.json();
  const block = data.content?.find((b) => b.type === "text");
  const raw = block?.text?.trim() ?? "";
  const jsonText = raw.replace(/^```json\s*/i, "").replace(/```\s*$/, "").trim();
  const parsed = JSON.parse(jsonText);
  return normalizeMilestones(parsed, totalMonths);
}

/** @typedef {{ ANTHROPIC_API_KEY: string, ANTHROPIC_MODEL?: string, ALLOWED_ORIGINS: string, RATE_LIMIT_PER_HOUR?: string }} Env */

/** @type {Map<string, { start: number, count: number }>} */
const rateLimitBuckets = new Map();

/** @param {string} ip @param {number} maxPerHour */
function isRateLimited(ip, maxPerHour) {
  if (!ip || maxPerHour <= 0) return false;
  const now = Date.now();
  const windowMs = 60 * 60 * 1000;
  let bucket = rateLimitBuckets.get(ip);
  if (!bucket || now - bucket.start > windowMs) {
    bucket = { start: now, count: 0 };
  }
  bucket.count += 1;
  rateLimitBuckets.set(ip, bucket);
  return bucket.count > maxPerHour;
}

export default {
  /** @param {Request} request @param {Env} env */
  async fetch(request, env) {
    const origin = request.headers.get("Origin") ?? "";
    const allowed = env.ALLOWED_ORIGINS ?? "";
    const cors = corsHeaders(origin, allowed);

    if (request.method === "OPTIONS") {
      if (!Object.keys(cors).length) {
        return new Response(null, { status: 403 });
      }
      return new Response(null, { status: 204, headers: cors });
    }

    if (request.method !== "POST") {
      return new Response(JSON.stringify({ error: "Method not allowed" }), {
        status: 405,
        headers: { "Content-Type": "application/json", ...cors },
      });
    }

    if (!Object.keys(cors).length) {
      return new Response(JSON.stringify({ error: "Origin not allowed" }), {
        status: 403,
        headers: { "Content-Type": "application/json" },
      });
    }

    const ip = request.headers.get("CF-Connecting-IP") ?? "unknown";
    const maxPerHour = Number(env.RATE_LIMIT_PER_HOUR) || 40;
    if (isRateLimited(ip, maxPerHour)) {
      return new Response(JSON.stringify({ error: "Rate limit exceeded — try again later." }), {
        status: 429,
        headers: { "Content-Type": "application/json", ...cors },
      });
    }

    try {
      const body = await request.json();
      const text = String(body.text ?? "").trim();
      const totalMonths = Math.min(120, Math.max(1, Number(body.totalMonths) || 36));

      if (!text || text.length > 4000) {
        return new Response(JSON.stringify({ error: "text required (max 4000 chars)" }), {
          status: 400,
          headers: { "Content-Type": "application/json", ...cors },
        });
      }

      const milestones = await parseWithAnthropic(text, totalMonths, env);
      return new Response(JSON.stringify({ milestones, source: "ai" }), {
        status: 200,
        headers: { "Content-Type": "application/json", ...cors },
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : "Parse failed";
      return new Response(JSON.stringify({ error: message }), {
        status: 500,
        headers: { "Content-Type": "application/json", ...cors },
      });
    }
  },
};
