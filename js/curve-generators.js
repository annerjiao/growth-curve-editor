import { computeKnotMonths, enforceMonotonicKnots } from "./curve-math.js";

/** @typedef {'milestone'|'linear'|'s_curve'|'hockey_stick'|'delayed_launch'} CurveType */

/**
 * @param {number} totalMonths
 * @param {number} start
 * @param {number} end
 */
export function generateLinear(totalMonths, start, end) {
  const knotMonths = computeKnotMonths(totalMonths);
  const span = Math.max(end - start, 0);
  const knotCumulative = knotMonths.map((m) =>
    Math.round(start + (span * m) / totalMonths)
  );
  return { knotMonths, knotCumulative: enforceMonotonicKnots(knotCumulative) };
}

/**
 * @param {{ month: number, value: number, label?: string }[]} milestones
 * @param {number} totalMonths
 */
export function normalizeMilestones(milestones, totalMonths) {
  /** @type {{ month: number, value: number, label: string }[]} */
  const sorted = [...(milestones || [])]
    .map((m) => ({
      month: Math.min(totalMonths, Math.max(1, Math.round(Number(m.month) || 1))),
      value: Math.max(0, Number(m.value) || 0),
      label: String(m.label ?? "").trim(),
    }))
    .sort((a, b) => a.month - b.month);

  /** @type {{ month: number, value: number, label: string }[]} */
  const deduped = [];
  for (const m of sorted) {
    const prev = deduped[deduped.length - 1];
    if (prev && prev.month === m.month) {
      prev.value = m.value;
      if (m.label) prev.label = m.label;
    } else {
      deduped.push({ ...m });
    }
  }

  if (!deduped.length) {
    deduped.push({ month: totalMonths, value: 0, label: "" });
  } else if (deduped[deduped.length - 1].month !== totalMonths) {
    const last = deduped[deduped.length - 1];
    deduped.push({ month: totalMonths, value: last.value, label: last.label });
  }

  return deduped;
}

/**
 * @param {number} totalMonths
 * @param {{ month: number, value: number, label?: string }[]} milestones
 */
export function generateMilestone(totalMonths, milestones) {
  const sorted = normalizeMilestones(milestones, totalMonths);
  const knotMonths = sorted.map((m) => m.month);
  const knotCumulative = enforceMonotonicKnots(sorted.map((m) => m.value));
  const knotLabels = sorted.map((m) => m.label);
  return { knotMonths, knotCumulative, knotLabels };
}

/**
 * @param {number} totalMonths
 * @param {number} early
 * @param {number} mid
 * @param {number} finalValue
 */
export function generateSCurve(totalMonths, early, mid, finalValue) {
  const earlyMonth = Math.max(1, Math.round(totalMonths * 0.17));
  const midMonth = Math.max(earlyMonth + 1, Math.round(totalMonths * 0.5));
  return generateMilestone(totalMonths, [
    { month: earlyMonth, value: early },
    { month: midMonth, value: mid },
    { month: totalMonths, value: finalValue },
  ]);
}

/**
 * @param {number} totalMonths
 * @param {number} flatUntil
 * @param {number} flatLevel
 * @param {number} finalValue
 */
export function generateHockeyStick(totalMonths, flatUntil, flatLevel, finalValue) {
  const flatMonth = Math.min(Math.max(1, flatUntil), totalMonths - 1);
  const knotMonths = computeKnotMonths(totalMonths);
  const knotCumulative = knotMonths.map((m) => {
    if (m <= flatMonth) {
      const t = flatMonth ? m / flatMonth : 0;
      return Math.round(flatLevel * t * 0.35);
    }
    const t = (m - flatMonth) / (totalMonths - flatMonth);
    const eased = Math.pow(t, 2.2);
    return Math.round(flatLevel + (finalValue - flatLevel) * eased);
  });
  knotCumulative[knotCumulative.length - 1] = finalValue;
  return { knotMonths, knotCumulative: enforceMonotonicKnots(knotCumulative) };
}

/**
 * @param {number} totalMonths
 * @param {number} launchMonth
 * @param {number} finalValue
 * @param {'linear'|'s_curve'} ramp
 */
export function generateDelayedLaunch(
  totalMonths,
  launchMonth,
  finalValue,
  ramp = "s_curve"
) {
  const launch = Math.min(Math.max(1, launchMonth), totalMonths - 1);
  const post = generateSCurve(
    totalMonths - launch + 1,
    Math.round(finalValue * 0.05),
    Math.round(finalValue * 0.45),
    finalValue
  );
  if (ramp === "linear") {
    Object.assign(
      post,
      generateLinear(totalMonths - launch + 1, 0, finalValue)
    );
  }

  const knotMonths = computeKnotMonths(totalMonths);
  const knotCumulative = knotMonths.map((m) => {
    if (m < launch) return 0;
    const idx = Math.min(
      post.knotMonths.length - 1,
      post.knotMonths.findIndex((km) => km >= m - launch + 1)
    );
    const localMonth = m - launch + 1;
    const prevIdx = Math.max(0, idx - 1);
    const m0 = post.knotMonths[prevIdx];
    const m1 = post.knotMonths[idx];
    const v0 = post.knotCumulative[prevIdx];
    const v1 = post.knotCumulative[idx];
    const span = m1 - m0 || 1;
    const t = (localMonth - m0) / span;
    return Math.round(v0 + (v1 - v0) * t);
  });
  knotCumulative[knotCumulative.length - 1] = finalValue;
  return { knotMonths, knotCumulative: enforceMonotonicKnots(knotCumulative) };
}

export const CURVE_TYPES = [
  {
    id: "milestone",
    label: "Milestone-led",
    description: "Set targets at specific months — most flexible.",
  },
  {
    id: "linear",
    label: "Linear ramp",
    description: "Steady growth from start to end.",
  },
  {
    id: "s_curve",
    label: "S-curve",
    description: "Slow start, fast middle, taper at end.",
  },
  {
    id: "hockey_stick",
    label: "Hockey stick",
    description: "Flat early phase, then sharp lift.",
  },
  {
    id: "delayed_launch",
    label: "Delayed launch",
    description: "Near-zero until launch, then ramp.",
  },
];

/**
 * @param {number[]} cumulative
 * @returns {{ milestones: { month: number, value: number, label: string }[] }}
 */
export function inferMilestoneParams(cumulative) {
  const T = cumulative.length || 36;
  const final = cumulative[T - 1] || 0;
  const max = cumulative.length ? Math.max(...cumulative, 0) : 0;

  if (T < 2 || max === 0) {
    return {
      milestones: [
        { month: Math.max(1, Math.round(T * 0.33)), value: 500, label: "" },
        { month: Math.max(2, Math.round(T * 0.66)), value: 5000, label: "" },
        { month: T, value: 15000, label: "" },
      ],
    };
  }

  const m1 = Math.max(1, Math.round(T * 0.33));
  const m2 = Math.max(m1 + 1, Math.round(T * 0.66));
  return {
    milestones: [
      { month: m1, value: cumulative[m1 - 1] || 0, label: "" },
      { month: m2, value: cumulative[m2 - 1] || 0, label: "" },
      { month: T, value: final, label: "" },
    ],
  };
}

/**
 * @param {number[]} cumulative
 * @returns {{ type: CurveType, params: Record<string, number|object> }}
 */
export function inferCurveType(cumulative) {
  const T = cumulative.length;
  if (T < 2) {
    return {
      type: "milestone",
      params: {
        milestones: [{ month: T, value: cumulative[T - 1] || 0, label: "" }],
      },
    };
  }

  const final = cumulative[T - 1] || 1;
  const q1End = Math.floor(T / 4);
  const q1Sum = cumulative[q1End - 1] || 0;

  if (final > 0 && q1Sum / final < 0.05) {
    const launchMonth =
      cumulative.findIndex((v) => v > final * 0.02) + 1 || Math.floor(T / 3);
    return {
      type: "delayed_launch",
      params: { launchMonth, finalValue: final },
    };
  }

  const third = Math.floor(T / 3);
  const slope = (from, to) => {
    if (to <= from) return 0;
    return (cumulative[to - 1] - cumulative[from - 1]) / (to - from);
  };
  const s1 = slope(1, third);
  const s2 = slope(third + 1, third * 2);
  const s3 = slope(third * 2 + 1, T);

  if (s1 > 0 && s3 > 0 && s2 > s1 * 1.2 && s2 > s3 * 1.2) {
    return {
      type: "s_curve",
      params: {
        early: cumulative[Math.floor(T * 0.17) - 1] || cumulative[0],
        mid: cumulative[Math.floor(T * 0.5) - 1] || Math.round(final * 0.45),
        finalValue: final,
      },
    };
  }

  if (s1 > 0 && s3 > s1 * 4) {
    const flatUntil = Math.floor(T / 3);
    return {
      type: "hockey_stick",
      params: {
        flatUntil,
        flatLevel: cumulative[flatUntil - 1] || Math.round(final * 0.05),
        finalValue: final,
      },
    };
  }

  let ssRes = 0;
  let ssTot = 0;
  const mean = final / 2;
  for (let i = 0; i < T; i++) {
    const expected = (final * (i + 1)) / T;
    ssRes += (cumulative[i] - expected) ** 2;
    ssTot += (cumulative[i] - mean) ** 2;
  }
  const r2 = ssTot ? 1 - ssRes / ssTot : 0;
  if (r2 > 0.92) {
    return { type: "linear", params: { start: 0, end: final } };
  }

  const m1 = Math.max(1, Math.round(T * 0.33));
  const m2 = Math.max(m1 + 1, Math.round(T * 0.66));
  return {
    type: "milestone",
    params: inferMilestoneParams(cumulative),
  };
}

/**
 * @param {CurveType} type
 * @param {number} totalMonths
 * @param {Record<string, unknown>} params
 */
export function generateFromType(type, totalMonths, params) {
  switch (type) {
    case "linear":
      return generateLinear(totalMonths, params.start ?? 0, params.end ?? 0);
    case "s_curve":
      return generateSCurve(
        totalMonths,
        params.early ?? 0,
        params.mid ?? 0,
        params.finalValue ?? 0
      );
    case "hockey_stick":
      return generateHockeyStick(
        totalMonths,
        params.flatUntil ?? 12,
        params.flatLevel ?? 0,
        params.finalValue ?? 0
      );
    case "delayed_launch":
      return generateDelayedLaunch(
        totalMonths,
        params.launchMonth ?? 6,
        params.finalValue ?? 0,
        params.ramp ?? "s_curve"
      );
    case "milestone":
    default:
      return generateMilestone(totalMonths, params.milestones ?? []);
  }
}
