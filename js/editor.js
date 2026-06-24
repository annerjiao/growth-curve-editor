import {
  curveFromKnots,
  formatAxis,
  fmt,
  enforceMonotonicKnots,
  enforceMonotonicKnotMonths,
} from "./curve-math.js";

const CASE_COLORS = [
  "#2f4f46",
  "#3d6b4f",
  "#b54a4a",
  "#c9a227",
  "#5a6b8a",
  "#8a5a7a",
];

function themeColor(name, fallback) {
  const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  return v || fallback;
}

const PAD = { top: 28, right: 24, bottom: 44, left: 72 };
const LABEL_ANGLE = -Math.PI / 4;

/** @param {CanvasRenderingContext2D} ctx @param {number} x @param {number} y @param {string} label @param {number} i @param {number | null} prevX @param {number[]} knotMonths */
function drawMilestoneLabel(ctx, x, y, label, i, prevX, knotMonths) {
  const monthGap = i > 0 ? knotMonths[i] - knotMonths[i - 1] : Infinity;
  const pxGap = prevX != null ? Math.abs(x - prevX) : Infinity;

  let lift = 10;
  let angle = LABEL_ANGLE;
  if (monthGap <= 4 || pxGap < 72) {
    lift += i % 2 === 0 ? 0 : 16;
    angle = i % 2 === 0 ? LABEL_ANGLE : -LABEL_ANGLE;
  }

  ctx.save();
  ctx.translate(x, y - lift);
  ctx.rotate(angle);
  ctx.fillStyle = themeColor("--chart-label", "#8a847a");
  ctx.font = '10px "IBM Plex Mono", monospace';
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillText(label, 6, 0);
  ctx.restore();
}

/**
 * @param {HTMLElement} canvas
 * @param {object} opts
 */
export function createEditor(canvas, opts = {}) {
  const ctx = canvas.getContext("2d");
  let scenarios = {};
  let activeKey = "";
  let totalMonths = 36;
  let dragIndex = -1;
  let allowKnotMonthEdit = false;
  let onChange = opts.onChange ?? (() => {});

  function getActive() {
    return scenarios[activeKey];
  }

  function plotBounds(c) {
    const derived = c.derived;
    return {
      maxCum: Math.max(...derived.cumulative, 1),
      maxSign: Math.max(...derived.signups, 1),
    };
  }

  function toCanvas(month, cumValue, maxCum, W, H) {
    const w = W - PAD.left - PAD.right;
    const h = H - PAD.top - PAD.bottom;
    const x = PAD.left + ((month - 1) / (totalMonths - 1)) * w;
    const y = PAD.top + h - (cumValue / maxCum) * h;
    return { x, y };
  }

  function fromCanvasX(x, W) {
    const w = W - PAD.left - PAD.right;
    const ratio = (x - PAD.left) / w;
    return Math.max(
      1,
      Math.min(totalMonths, Math.round(1 + ratio * (totalMonths - 1)))
    );
  }

  function fromCanvasY(y, maxCum, H) {
    const h = H - PAD.top - PAD.bottom;
    const ratio = 1 - (y - PAD.top) / h;
    return Math.max(0, ratio * maxCum);
  }

  function recompute(key, persist = true) {
    const c = scenarios[key];
    if (!c) return;
    c.derived = curveFromKnots(c.knotMonths, c.knotCumulative, totalMonths);
    c.flow = c.derived.signups;
    c.stats = {
      p25: c.derived.cumulative[Math.floor(totalMonths * 0.25) - 1] ?? 0,
      p50: c.derived.cumulative[Math.floor(totalMonths * 0.5) - 1] ?? 0,
      final: c.derived.cumulative[totalMonths - 1] ?? 0,
    };
    if (persist) onChange(key, c);
    if (key === activeKey) render();
  }

  function renderChart() {
    const c = getActive();
    if (!c?.derived) return;

    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    if (rect.width <= 0) return;

    canvas.width = Math.round(rect.width * dpr);
    canvas.height = Math.round(420 * dpr);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const W = rect.width;
    const H = 420;

    ctx.clearRect(0, 0, W, H);
    const { cumulative, signups } = c.derived;
    const { maxCum, maxSign } = plotBounds(c);
    const colorIdx = Object.keys(scenarios).indexOf(activeKey);
    const color = CASE_COLORS[colorIdx % CASE_COLORS.length];

    for (const [key, oc] of Object.entries(scenarios)) {
      if (key === activeKey) continue;
      const cum =
        oc.derived?.cumulative ??
        oc.flow?.map((_, i) =>
          oc.flow.slice(0, i + 1).reduce((a, b) => a + b, 0)
        );
      if (!cum) continue;
      const omax = Math.max(...cum, 1);
      const idx = Object.keys(scenarios).indexOf(key);
      ctx.strokeStyle = CASE_COLORS[idx % CASE_COLORS.length];
      ctx.globalAlpha = 0.2;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      cum.forEach((v, i) => {
        const x = PAD.left + (i / (totalMonths - 1)) * (W - PAD.left - PAD.right);
        const y =
          PAD.top +
          (H - PAD.top - PAD.bottom) -
          (v / omax) * (H - PAD.top - PAD.bottom);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
      ctx.globalAlpha = 1;
    }

    ctx.strokeStyle = themeColor("--chart-grid", "#e5e2db");
    ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
      const y = PAD.top + ((H - PAD.top - PAD.bottom) * i) / 4;
      ctx.beginPath();
      ctx.moveTo(PAD.left, y);
      ctx.lineTo(W - PAD.right, y);
      ctx.stroke();
    }

    ctx.fillStyle = themeColor("--chart-label", "#8a847a");
    ctx.font = '11px "IBM Plex Mono", monospace';
    ctx.textAlign = "right";
    for (let i = 0; i <= 4; i++) {
      const val = maxCum * (1 - i / 4);
      const y = PAD.top + ((H - PAD.top - PAD.bottom) * i) / 4;
      ctx.fillText(formatAxis(val), PAD.left - 8, y + 4);
    }

    ctx.textAlign = "center";
    c.knotMonths.forEach((m) => {
      const { x } = toCanvas(m, 0, maxCum, W, H);
      ctx.fillText(`M${m}`, x, H - 12);
    });

    const bandH = 48;
    signups.forEach((s, i) => {
      const { x } = toCanvas(i + 1, 0, maxCum, W, H);
      const barW = Math.max(2, (W - PAD.left - PAD.right) / totalMonths - 1);
      const barH = (s / maxSign) * bandH;
      ctx.fillStyle = themeColor("--chart-flow", "rgba(61, 107, 79, 0.42)");
      ctx.fillRect(x - barW / 2, H - PAD.bottom - barH, barW, barH);
    });

    ctx.strokeStyle = color;
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    cumulative.forEach((v, i) => {
      const { x, y } = toCanvas(i + 1, v, maxCum, W, H);
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();

    let prevKnotX = null;
    c.knotMonths.forEach((m, i) => {
      const { x, y } = toCanvas(m, c.knotCumulative[i], maxCum, W, H);
      ctx.fillStyle = i === dragIndex ? "#ffffff" : themeColor("--chart-knot", "#c9a227");
      ctx.strokeStyle = themeColor("--chart-knot-stroke", "#1a1814");
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(x, y, 8, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();

      const label = c.knotLabels?.[i];
      if (label) {
        drawMilestoneLabel(ctx, x, y, label, i, prevKnotX, c.knotMonths);
      }
      prevKnotX = x;
    });
  }

  function hitTestKnot(px, py, W, H) {
    const c = getActive();
    const { maxCum } = plotBounds(c);
    for (let i = 0; i < c.knotMonths.length; i++) {
      const { x, y } = toCanvas(c.knotMonths[i], c.knotCumulative[i], maxCum, W, H);
      if (Math.hypot(px - x, py - y) < 14) return i;
    }
    return -1;
  }

  function onPointerDown(e) {
    const rect = canvas.getBoundingClientRect();
    dragIndex = hitTestKnot(
      e.clientX - rect.left,
      e.clientY - rect.top,
      rect.width,
      420
    );
    if (dragIndex >= 0) canvas.setPointerCapture(e.pointerId);
  }

  function onPointerMove(e) {
    if (dragIndex < 0) return;
    const c = getActive();
    const rect = canvas.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;
    const W = rect.width;
    const { maxCum } = plotBounds(c);

    if (allowKnotMonthEdit && c.knotMonths.length > 1) {
      const minMonth = dragIndex > 0 ? c.knotMonths[dragIndex - 1] + 1 : 1;
      const maxMonth =
        dragIndex < c.knotMonths.length - 1
          ? c.knotMonths[dragIndex + 1] - 1
          : totalMonths;
      c.knotMonths[dragIndex] = Math.max(
        minMonth,
        Math.min(maxMonth, fromCanvasX(px, W))
      );
      c.knotMonths = enforceMonotonicKnotMonths(c.knotMonths, totalMonths);
    }

    let val = fromCanvasY(py, maxCum, 420);
    const floor = dragIndex > 0 ? c.knotCumulative[dragIndex - 1] : 0;
    const ceil =
      dragIndex < c.knotCumulative.length - 1
        ? c.knotCumulative[dragIndex + 1]
        : Infinity;
    c.knotCumulative[dragIndex] = Math.round(
      Math.max(floor, Math.min(ceil, val))
    );
    recompute(activeKey, true);
  }

  function onPointerUp(e) {
    dragIndex = -1;
    try {
      canvas.releasePointerCapture(e.pointerId);
    } catch (_) {}
  }

  canvas.addEventListener("pointerdown", onPointerDown);
  canvas.addEventListener("pointermove", onPointerMove);
  canvas.addEventListener("pointerup", onPointerUp);
  canvas.addEventListener("pointerleave", onPointerUp);
  window.addEventListener("resize", () => render());

  if (typeof ResizeObserver !== "undefined") {
    const ro = new ResizeObserver((entries) => {
      if (entries[0]?.contentRect.width > 0) render();
    });
    ro.observe(canvas);
  }

  return {
    setScenarios(data, months, active, options = {}) {
      scenarios = data;
      totalMonths = months;
      activeKey = active || Object.keys(scenarios)[0];
      allowKnotMonthEdit = Boolean(options.allowKnotMonthEdit);
      for (const key of Object.keys(scenarios)) {
        const c = scenarios[key];
        if (!c.knotLabels || c.knotLabels.length !== c.knotMonths.length) {
          c.knotLabels = c.knotMonths.map((_, i) => c.knotLabels?.[i] ?? "");
        }
        recompute(key, false);
      }
      render();
    },
    selectScenario(key) {
      activeKey = key;
      render();
    },
    getScenarios() {
      return scenarios;
    },
    getActiveKey() {
      return activeKey;
    },
    updateKnot(i, value) {
      const c = getActive();
      c.knotCumulative[i] = Math.max(0, Number(value) || 0);
      enforceMonotonicKnots(c.knotCumulative);
      recompute(activeKey, true);
    },
    updateKnotMonth(i, month) {
      const c = getActive();
      c.knotMonths[i] = Number(month) || 1;
      c.knotMonths = enforceMonotonicKnotMonths(c.knotMonths, totalMonths);
      recompute(activeKey, true);
    },
    updateKnotLabel(i, label) {
      const c = getActive();
      if (!c.knotLabels) c.knotLabels = [];
      c.knotLabels[i] = label;
      onChange(activeKey, c);
      render();
    },
    deleteKnot(i) {
      const c = getActive();
      if (!c || c.knotMonths.length <= 2) return;
      c.knotMonths.splice(i, 1);
      c.knotCumulative.splice(i, 1);
      if (c.knotLabels) c.knotLabels.splice(i, 1);
      c.knotMonths = enforceMonotonicKnotMonths(c.knotMonths, totalMonths);
      enforceMonotonicKnots(c.knotCumulative);
      recompute(activeKey, true);
    },
    addKnot() {
      const c = getActive();
      if (!c || c.knotMonths.length >= 6) return false;

      let splitAt = 0;
      let bestGap = 0;
      for (let i = 0; i < c.knotMonths.length - 1; i++) {
        const gap = c.knotMonths[i + 1] - c.knotMonths[i];
        if (gap > bestGap) {
          bestGap = gap;
          splitAt = i;
        }
      }
      if (bestGap < 2) return false;

      const m0 = c.knotMonths[splitAt];
      const m1 = c.knotMonths[splitAt + 1];
      const newMonth = Math.max(m0 + 1, Math.min(m1 - 1, Math.round((m0 + m1) / 2)));
      const v0 = c.knotCumulative[splitAt];
      const v1 = c.knotCumulative[splitAt + 1];
      const newValue = Math.round((v0 + v1) / 2);

      c.knotMonths.splice(splitAt + 1, 0, newMonth);
      c.knotCumulative.splice(splitAt + 1, 0, newValue);
      if (!c.knotLabels) c.knotLabels = c.knotMonths.map(() => "");
      c.knotLabels.splice(splitAt + 1, 0, "");

      c.knotMonths = enforceMonotonicKnotMonths(c.knotMonths, totalMonths);
      enforceMonotonicKnots(c.knotCumulative);
      recompute(activeKey, true);
      return true;
    },
    render,
    fmt,
  };

  function render() {
    renderChart();
    opts.onRender?.(scenarios, activeKey, totalMonths);
  }
}

/** @param {HTMLCanvasElement} canvas */
export function drawMiniChart(canvas, cumulative, color, bold = false) {
  const mctx = canvas.getContext("2d");
  const W = canvas.width;
  const H = canvas.height;
  mctx.clearRect(0, 0, W, H);
  if (!cumulative?.length) return;
  const max = Math.max(...cumulative, 1);
  mctx.strokeStyle = color;
  mctx.lineWidth = bold ? 2.5 : 1.5;
  mctx.globalAlpha = bold ? 1 : 0.65;
  mctx.beginPath();
  cumulative.forEach((v, i) => {
    const x = (i / (cumulative.length - 1)) * (W - 8) + 4;
    const y = H - 4 - (v / max) * (H - 8);
    if (i === 0) mctx.moveTo(x, y);
    else mctx.lineTo(x, y);
  });
  mctx.stroke();
  mctx.globalAlpha = 1;
}
