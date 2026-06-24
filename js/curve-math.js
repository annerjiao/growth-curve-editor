/**
 * Monotonic PCHIP interpolation for S-curve / hockey-stick growth.
 * Cumulative at knot periods → smooth curve → monthly flow values.
 */

export function pchipSlopes(x, y) {
  const n = x.length;
  const h = [];
  const delta = [];
  for (let i = 0; i < n - 1; i++) {
    h.push(x[i + 1] - x[i]);
    delta.push(h[i] ? (y[i + 1] - y[i]) / h[i] : 0);
  }
  const m = new Array(n).fill(0);
  for (let i = 1; i < n - 1; i++) {
    if (delta[i - 1] * delta[i] <= 0) m[i] = 0;
    else {
      const w1 = 2 * h[i] + h[i - 1];
      const w2 = h[i] + 2 * h[i - 1];
      m[i] = (w1 + w2) / (w1 / delta[i - 1] + w2 / delta[i]);
    }
  }
  m[0] = delta[0];
  m[n - 1] = delta[n - 2];
  return { m, h };
}

export function pchipEval(x, y, xq) {
  const { m, h } = pchipSlopes(x, y);
  return xq.map((x0) => {
    if (x0 <= x[0]) return y[0];
    if (x0 >= x[x.length - 1]) return y[y.length - 1];
    let i = 0;
    while (i < x.length - 1 && x0 > x[i + 1]) i++;
    const t = (x0 - x[i]) / h[i];
    const t2 = t * t;
    const t3 = t2 * t;
    return (
      y[i] * (2 * t3 - 3 * t2 + 1) +
      y[i + 1] * (-2 * t3 + 3 * t2) +
      m[i] * h[i] * (t3 - 2 * t2 + t) +
      m[i + 1] * h[i] * (t3 - t2)
    );
  });
}

/** @param {number[]} knotMonths */
/** @param {number[]} knotCumulative */
/** @param {number} totalMonths */
export function curveFromKnots(knotMonths, knotCumulative, totalMonths = 36) {
  const x = [0, ...knotMonths];
  const y = [0, ...knotCumulative];
  const months = Array.from({ length: totalMonths }, (_, i) => i + 1);
  let cumulative = pchipEval(x, y, months);

  for (let i = 1; i < cumulative.length; i++) {
    cumulative[i] = Math.max(cumulative[i], cumulative[i - 1]);
  }

  const signups = [];
  let prev = 0;
  for (let i = 0; i < cumulative.length; i++) {
    let s = Math.max(0, Math.round(cumulative[i]) - Math.round(prev));
    if (i === 0 && s === 0) s = 1;
    signups.push(s);
    prev += s;
  }

  const target = knotCumulative[knotCumulative.length - 1];
  const diff = target - signups.reduce((a, b) => a + b, 0);
  signups[signups.length - 1] += diff;

  cumulative = signups.map((_, i) =>
    signups.slice(0, i + 1).reduce((a, b) => a + b, 0)
  );

  return { months, signups, cumulative };
}

export function signupsToKnots(signups, knotMonths) {
  return knotMonths.map((m) =>
    signups.slice(0, m).reduce((a, b) => a + b, 0)
  );
}

export function computeKnotMonths(totalMonths, count = 6) {
  if (totalMonths <= 1) return [1];
  if (count <= 1) return [totalMonths];

  const raw = [];
  for (let i = 1; i <= count; i++) {
    raw.push(Math.max(1, Math.round((totalMonths * i) / count)));
  }
  const unique = [...new Set(raw)];
  if (unique[unique.length - 1] !== totalMonths) {
    unique[unique.length - 1] = totalMonths;
  }
  return unique;
}

export function flowToCumulative(flow) {
  const cumulative = [];
  let sum = 0;
  for (const v of flow) {
    sum += Number(v) || 0;
    cumulative.push(sum);
  }
  return cumulative;
}

export function enforceMonotonicKnotMonths(knotMonths, totalMonths) {
  const months = knotMonths.map((m) =>
    Math.min(totalMonths, Math.max(1, Math.round(m)))
  );
  for (let i = 1; i < months.length; i++) {
    months[i] = Math.max(months[i], months[i - 1] + 1);
  }
  months[months.length - 1] = totalMonths;
  return months;
}

export function enforceMonotonicKnots(knotCumulative) {
  for (let i = 1; i < knotCumulative.length; i++) {
    knotCumulative[i] = Math.max(knotCumulative[i], knotCumulative[i - 1]);
  }
  return knotCumulative;
}

export function formatAxis(n) {
  if (n >= 1e6) return (n / 1e6).toFixed(1) + "M";
  if (n >= 1e3) return (n / 1e3).toFixed(0) + "K";
  return String(Math.round(n));
}

export function fmt(n) {
  return (n ?? 0).toLocaleString();
}
