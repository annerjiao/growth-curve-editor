import {
  computeKnotMonths,
  curveFromKnots,
  flowToCumulative,
  signupsToKnots,
} from "./curve-math.js";
import {
  CURVE_TYPES,
  generateFromType,
  inferCurveType,
} from "./curve-generators.js";
import {
  downloadUpdatedWorkbook,
  getSheetPreview,
  listSheetNames,
  loadWorkbook,
  readFlowRange,
} from "./xlsx-io.js";
import { createEditor, drawMiniChart } from "./editor.js";
import { parseMilestoneText, MILESTONE_PARSE_URL } from "./milestone-parser.js";
import {
  analyzeMappedRange,
  detectDriverCandidates,
  findSuggestedSheet,
} from "./range-detector.js";

const CASE_COLORS = [
  "#2f4f46",
  "#3d6b4f",
  "#b54a4a",
  "#c9a227",
  "#5a6b8a",
  "#8a5a7a",
];

/** @type {import('xlsx').WorkBook | null} */
let workbook = null;
/** @type {ArrayBuffer | null} */
let originalWorkbookBuffer = null;
let workbookName = "model-updated.xlsx";
let demoMode = false;

/** @type {{ id: string, label: string, sheet: string, rangeRef: string, flow: number[], knotMonths: number[], knotCumulative: number[], knotLabels: string[] }[]} */
let scenarios = [];

let totalMonths = 36;
let metricLabel = "Users";
let activeScenarioId = "";
let selectedCurveType = "milestone";
/** @type {Record<string, unknown>} */
let curveParams = {};
let mapListenersBound = false;
/** @type {import('./range-detector.js').DriverCandidate[]} */
let mapSuggestions = [];
let selectedSuggestionKey = "";
/** @type {'sheet' | 'suggest' | 'manual'} */
let mapPhase = "sheet";

const editor = createEditor(document.getElementById("chart"), {
  onChange: (key, c) => {
    const s = scenarios.find((x) => x.id === key);
    if (s) {
      s.flow = c.flow;
      s.knotMonths = [...c.knotMonths];
      s.knotCumulative = [...c.knotCumulative];
      s.knotLabels = c.knotLabels ? [...c.knotLabels] : [];
    }
    renderCaseGrid();
    renderKnotTable();
    renderStats();
  },
  onRender: () => {
    renderCaseGrid();
    renderKnotTable();
    renderStats();
  },
});

function uid() {
  return "s_" + Math.random().toString(36).slice(2, 9);
}

function getActiveScenario() {
  return scenarios.find((s) => s.id === activeScenarioId);
}

function setStep(step) {
  document.querySelectorAll(".step-panel").forEach((el) => {
    el.classList.toggle("active", el.id === `step-${step}`);
  });
  document.querySelectorAll(".step-tab").forEach((btn) => {
    const n = btn.dataset.step;
    btn.classList.toggle("active", n === step);
    btn.disabled =
      n !== "upload" &&
      n !== step &&
      !isStepReachable(n, step);
  });
  if (step === "edit") {
    requestAnimationFrame(() => editor.render());
  }
}

function isStepReachable(target, current) {
  const order = ["upload", "map", "curve", "edit", "export"];
  const minIdx = scenarios.length ? order.indexOf("map") : 0;
  const currentIdx = order.indexOf(current);
  const targetIdx = order.indexOf(target);
  if (target === "upload") return true;
  if (targetIdx <= currentIdx) return true;
  return targetIdx <= minIdx + (scenarios.length ? 4 : 0);
}

function enableStepsFrom(step) {
  const order = ["upload", "map", "curve", "edit", "export"];
  const idx = order.indexOf(step);
  document.querySelectorAll(".step-tab").forEach((btn) => {
    btn.disabled = order.indexOf(btn.dataset.step) > idx;
  });
}

// --- Upload ---

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const uploadStatus = document.getElementById("uploadStatus");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer?.files?.[0];
  if (file) handleFile(file);
});
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
});

async function handleFile(file) {
  if (!file.name.match(/\.xlsx$/i)) {
    uploadStatus.textContent = "Please upload a .xlsx file.";
    return;
  }
  const buffer = await file.arrayBuffer();
  originalWorkbookBuffer = buffer.slice(0);
  workbook = loadWorkbook(originalWorkbookBuffer);
  workbookName = file.name.replace(/\.xlsx$/i, "-updated.xlsx");
  demoMode = false;
  scenarios = [];
  uploadStatus.textContent = `Loaded ${file.name}`;
  initMapStep();
  enableStepsFrom("map");
  setStep("map");
}

document.getElementById("scratchBtn").addEventListener("click", () => {
  workbook = null;
  originalWorkbookBuffer = null;
  demoMode = true;
  workbookName = "growth-model.xlsx";
  scenarios = [
    {
      id: uid(),
      label: "Base",
      sheet: "",
      rangeRef: "",
      flow: Array(36).fill(0),
      knotMonths: computeKnotMonths(36),
      knotCumulative: Array(6).fill(0),
      knotLabels: [],
    },
  ];
  totalMonths = 36;
  activeScenarioId = scenarios[0].id;
  metricLabel = "Users";
  initCurveStep();
  enableStepsFrom("curve");
  setStep("curve");
});

// --- Map ---

function initMapStep() {
  const sheets = listSheetNames(workbook);
  const sheetSelect = document.getElementById("sheetSelect");
  sheetSelect.innerHTML =
    `<option value="">Choose a sheet…</option>` +
    sheets.map((s) => `<option value="${escapeAttr(s)}">${escapeHtml(s)}</option>`).join("");

  mapPhase = "sheet";
  selectedSuggestionKey = "";
  mapSuggestions = [];

  scenarios = [
    {
      id: uid(),
      label: "Base",
      sheet: "",
      rangeRef: "",
      flow: [],
      knotMonths: [],
      knotCumulative: [],
      knotLabels: [],
    },
  ];

  if (!mapListenersBound) {
    sheetSelect.addEventListener("change", onSheetSelected);
    document.getElementById("mapManualBtn").addEventListener("click", showManualMapPhase);
    document.getElementById("mapBackToSuggestBtn").addEventListener("click", showSuggestMapPhase);
    mapListenersBound = true;
  }

  const defaultSheet = workbook ? findSuggestedSheet(workbook) : null;
  if (defaultSheet) {
    sheetSelect.value = defaultSheet;
    onSheetSelected();
  } else {
    updateMapPhaseUI();
  }
}

function onSheetSelected() {
  const sheet = document.getElementById("sheetSelect").value;
  document.getElementById("mapError").textContent = "";
  selectedSuggestionKey = "";

  if (!sheet) {
    mapPhase = "sheet";
    scenarios.forEach((s) => {
      s.sheet = "";
      s.rangeRef = "";
    });
    updateMapPhaseUI();
    return;
  }

  scenarios.forEach((s) => {
    s.sheet = sheet;
    s.rangeRef = "";
  });

  mapPhase = "suggest";
  refreshMapSuggestions();
  updateMapPhaseUI();
}

function showManualMapPhase() {
  mapPhase = "manual";
  selectedSuggestionKey = "";
  scenarios.forEach((s) => {
    s.rangeRef = "";
  });
  document.getElementById("mapError").textContent = "";
  renderScenarioRows();
  renderSheetPreview();
  updateMapPhaseUI();
}

function showSuggestMapPhase() {
  if (!document.getElementById("sheetSelect").value) {
    mapPhase = "sheet";
  } else {
    mapPhase = "suggest";
  }
  selectedSuggestionKey = "";
  scenarios.forEach((s) => {
    s.rangeRef = "";
  });
  document.getElementById("mapError").textContent = "";
  refreshMapSuggestions();
  updateMapPhaseUI();
}

function updateMapPhaseUI() {
  const sheet = document.getElementById("sheetSelect").value;
  const hint = document.getElementById("mapHint");
  const continueBtn = document.getElementById("mapContinueBtn");

  document.getElementById("mapSuggestPhase").classList.toggle("hidden", mapPhase === "sheet");
  document.getElementById("mapManualPhase").classList.toggle("hidden", mapPhase !== "manual");
  document.getElementById("mapSelectedBar").classList.add("hidden");

  if (mapPhase === "sheet") {
    hint.textContent =
      "Choose the sheet that contains your monthly growth numbers (users, signups, etc.).";
    continueBtn.disabled = true;
    return;
  }

  if (mapPhase === "suggest") {
    hint.textContent = sheet
      ? `Scanning "${sheet}" for rows that change across month columns.`
      : "Choose a sheet to see suggestions.";
    continueBtn.disabled = !selectedSuggestionKey;
    return;
  }

  hint.textContent = "Enter the cell range for your driver row.";
  continueBtn.disabled = false;
  refreshRangeWarning();
}

function suggestionKey(c) {
  return `${c.sheet}!${c.rangeRef}`;
}

function refreshMapSuggestions() {
  const list = document.getElementById("suggestionList");
  if (!workbook) {
    list.innerHTML = "";
    mapSuggestions = [];
    return;
  }

  const sheet = document.getElementById("sheetSelect").value;
  if (!sheet) {
    list.innerHTML = "";
    mapSuggestions = [];
    return;
  }

  mapSuggestions = detectDriverCandidates(workbook, sheet);

  if (mapSuggestions.length === 0) {
    list.innerHTML =
      "<p class='suggestion-empty'>No matches on this sheet. Try another sheet, or enter a range manually.</p>";
    return;
  }

  list.innerHTML = "";

  mapSuggestions.forEach((c) => {
    const key = suggestionKey(c);
    const card = document.createElement("button");
    card.type = "button";
    card.className = `suggestion-card${key === selectedSuggestionKey ? " selected" : ""}`;
    card.dataset.key = key;

    const badges = [
      `<span class="suggestion-badge ${c.confidence}">${c.confidence} match</span>`,
      c.hasFormulas
        ? `<span class="suggestion-badge formula-warn">Has formulas</span>`
        : `<span class="suggestion-badge safe">Hardcoded</span>`,
    ].join("");

    card.innerHTML = `
      <div class="suggestion-card-head">
        <strong>${escapeHtml(c.metricName || c.label)}</strong>
        ${c.scenarioLabel && c.metricName ? `<span class="suggestion-scenario">${escapeHtml(c.scenarioLabel)}</span>` : ""}
        ${badges}
        <span class="suggestion-range">${escapeHtml(c.rangeRef)} · ${c.periodCount} periods</span>
      </div>
      <canvas class="suggestion-spark" width="120" height="36" aria-hidden="true"></canvas>
      <ul class="suggestion-reasons">${c.reasons.map((r) => `<li>${escapeHtml(r)}</li>`).join("")}</ul>`;

    const spark = card.querySelector(".suggestion-spark");
    if (c.values.length >= 2) {
      drawMiniChart(spark, c.values, "#2f4f46", false);
    }

    card.addEventListener("click", () => applySuggestion(c));
    list.appendChild(card);
  });
}

function applySuggestion(candidate) {
  selectedSuggestionKey = suggestionKey(candidate);
  scenarios[0].sheet = candidate.sheet;
  scenarios[0].rangeRef = candidate.rangeRef;
  if (candidate.scenarioLabel || candidate.label) {
    scenarios[0].label =
      (candidate.scenarioLabel || candidate.label).slice(0, 40) || "Base";
  }
  const metric =
    candidate.metricName || candidate.metricLabel || "Users";
  document.getElementById("metricLabel").value = metric;
  metricLabel = metric;

  document.getElementById("mapError").textContent = "";
  document.querySelectorAll(".suggestion-card").forEach((el) => {
    el.classList.toggle("selected", el.dataset.key === selectedSuggestionKey);
  });

  const bar = document.getElementById("mapSelectedBar");
  bar.classList.remove("hidden");
  bar.textContent = `Using ${candidate.metricName ? `${candidate.metricName}${candidate.scenarioLabel ? ` · ${candidate.scenarioLabel}` : ""}` : candidate.label} · ${candidate.rangeRef} (${candidate.periodCount} periods)`;

  document.getElementById("mapContinueBtn").disabled = false;
}

function refreshRangeWarning() {
  const el = document.getElementById("rangeWarning");
  if (!workbook || !scenarios[0]?.rangeRef?.trim()) {
    el.classList.add("hidden");
    el.textContent = "";
    return;
  }

  try {
    const s = scenarios[0];
    const analysis = analyzeMappedRange(workbook, s.sheet, s.rangeRef.trim());
    if (analysis.warnings.length === 0) {
      el.classList.add("hidden");
      el.textContent = "";
      return;
    }

    el.classList.remove("hidden");
    el.classList.toggle("info", analysis.formulaPct > 0 && analysis.formulaPct <= 0.5);
    el.innerHTML = analysis.warnings.map((w) => escapeHtml(w)).join(" ");
  } catch {
    el.classList.add("hidden");
    el.textContent = "";
  }
}

function renderScenarioRows() {
  const list = document.getElementById("scenarioList");
  list.innerHTML = "";
  scenarios.forEach((s, i) => {
    const row = document.createElement("div");
    row.className = "scenario-row";
    row.innerHTML = `
      <label>Scenario name
        <input type="text" data-field="label" data-id="${s.id}" value="${escapeAttr(s.label)}" />
      </label>
      <label>Cell range (e.g. C8:AN8)
        <input type="text" data-field="rangeRef" data-id="${s.id}" value="${escapeAttr(s.rangeRef)}" placeholder="C8:AN8" />
      </label>
      ${scenarios.length > 1 ? `<button type="button" class="remove-btn" data-remove="${s.id}">Remove</button>` : "<span></span>"}
    `;
    list.appendChild(row);
  });

  list.querySelectorAll("input").forEach((inp) => {
    inp.addEventListener("change", () => {
      const s = scenarios.find((x) => x.id === inp.dataset.id);
      if (s) s[inp.dataset.field] = inp.value;
      if (inp.dataset.field === "rangeRef") {
        selectedSuggestionKey = "";
        document.getElementById("mapSelectedBar").classList.add("hidden");
        refreshRangeWarning();
        document.querySelectorAll(".suggestion-card").forEach((el) => {
          el.classList.remove("selected");
        });
      }
    });
    inp.addEventListener("input", () => {
      if (inp.dataset.field !== "rangeRef") return;
      const s = scenarios.find((x) => x.id === inp.dataset.id);
      if (s) s[inp.dataset.field] = inp.value;
      selectedSuggestionKey = "";
      refreshRangeWarning();
    });
  });
  list.querySelectorAll("[data-remove]").forEach((btn) => {
    btn.addEventListener("click", () => {
      scenarios = scenarios.filter((x) => x.id !== btn.dataset.remove);
      renderScenarioRows();
    });
  });
}

function renderSheetPreview() {
  const el = document.getElementById("sheetPreview");
  if (!workbook) {
    el.innerHTML = "";
    return;
  }
  const sheet = document.getElementById("sheetSelect").value;
  const preview = getSheetPreview(workbook, sheet);
  if (!preview.rows.length) {
    el.innerHTML = "<p class='hint'>Sheet is empty.</p>";
    return;
  }
  el.innerHTML = `
    <p class="hint">Preview of <strong>${escapeHtml(sheet)}</strong> (top-left)</p>
    <table><tbody>
      ${preview.rows.map((row) => `<tr>${row.map((c) => `<td>${escapeHtml(String(c))}</td>`).join("")}</tr>`).join("")}
    </tbody></table>`;
}

document.getElementById("addScenarioBtn").addEventListener("click", () => {
  const sheet = document.getElementById("sheetSelect")?.value || "";
  scenarios.push({
    id: uid(),
    label: `Scenario ${scenarios.length + 1}`,
    sheet,
    rangeRef: "",
    flow: [],
    knotMonths: [],
    knotCumulative: [],
    knotLabels: [],
  });
  renderScenarioRows();
});

document.getElementById("mapContinueBtn").addEventListener("click", () => {
  const err = document.getElementById("mapError");
  err.textContent = "";
  if (mapPhase === "manual") {
    metricLabel = document.getElementById("metricLabel").value.trim() || "Users";
  }
  const sheet = document.getElementById("sheetSelect").value;
  if (!sheet) {
    err.textContent = "Choose a sheet first.";
    return;
  }

  try {
    for (const s of scenarios) {
      s.sheet = sheet;
      if (!s.rangeRef.trim()) throw new Error(`Enter a range for "${s.label}".`);
      if (workbook) {
        const analysis = analyzeMappedRange(workbook, s.sheet, s.rangeRef.trim());
        if (analysis.formulaPct > 0.5) {
          throw new Error(
            "This range is mostly formulas — overwriting will break your model. Pick a suggested hardcoded driver row, or map the input row your formulas reference."
          );
        }
        const data = readFlowRange(workbook, s.sheet, s.rangeRef.trim());
        s.flow = data.values;
        s.rangeRef = s.rangeRef.trim();
      }
      totalMonths = s.flow.length || totalMonths;
      s.knotMonths = computeKnotMonths(totalMonths);
      s.knotCumulative = signupsToKnots(s.flow, s.knotMonths);
      s.knotLabels = [];
    }
    if (scenarios.length === 0) throw new Error("Add at least one scenario.");
    if (scenarios.some((s) => s.flow.length < 2)) {
      throw new Error("Each range needs at least 2 periods.");
    }
    activeScenarioId = scenarios[0].id;
    initCurveStep();
    enableStepsFrom("curve");
    setStep("curve");
  } catch (e) {
    err.textContent = e.message;
  }
});

// --- Curve setup ---

function initCurveStep() {
  const active = getActiveScenario();
  if (!active) return;

  if (!active.flow?.length) {
    totalMonths = 36;
    active.flow = Array(totalMonths).fill(0);
    active.knotMonths = computeKnotMonths(totalMonths);
    active.knotCumulative = Array(active.knotMonths.length).fill(0);
  } else {
    totalMonths = active.flow.length;
  }

  const cumulative = flowToCumulative(active.flow);
  const inferred = inferCurveType(cumulative);
  selectedCurveType = inferred.type;
  curveParams = { ...inferred.params };

  renderCurveTypeGrid();
  renderCurveForm();
  updatePreviewChart();
}

function renderCurveTypeGrid() {
  const grid = document.getElementById("curveTypeGrid");
  grid.innerHTML = "";
  CURVE_TYPES.forEach((t) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `curve-type-card${t.id === selectedCurveType ? " active" : ""}`;
    btn.innerHTML = `<strong>${escapeHtml(t.label)}</strong><span>${escapeHtml(t.description)}</span>`;
    btn.addEventListener("click", () => {
      selectedCurveType = t.id;
      renderCurveTypeGrid();
      renderCurveForm();
      updatePreviewChart();
    });
    grid.appendChild(btn);
  });
}

function renderCurveForm() {
  const form = document.getElementById("curveForm");
  form.innerHTML = "";
  const T = totalMonths;

  const addField = (label, key, value, step = 1) => {
    const wrap = document.createElement("label");
    wrap.innerHTML = `${label}<input type="number" min="0" step="${step}" data-key="${key}" value="${value}" />`;
    form.appendChild(wrap);
  };

  switch (selectedCurveType) {
    case "linear":
      addField("Start (cumulative)", "start", curveParams.start ?? 0);
      addField(`End @ M${T}`, "end", curveParams.end ?? 10000);
      break;
    case "s_curve":
      addField("Early cumulative", "early", curveParams.early ?? 200);
      addField("Mid cumulative", "mid", curveParams.mid ?? Math.round((curveParams.finalValue ?? 10000) * 0.45));
      addField(`Final @ M${T}`, "finalValue", curveParams.finalValue ?? 10000);
      break;
    case "hockey_stick":
      addField("Flat through month", "flatUntil", curveParams.flatUntil ?? Math.floor(T / 3), 1);
      addField("Level after flat", "flatLevel", curveParams.flatLevel ?? 500);
      addField(`Final @ M${T}`, "finalValue", curveParams.finalValue ?? 20000);
      break;
    case "delayed_launch":
      addField("Launch month", "launchMonth", curveParams.launchMonth ?? 6, 1);
      addField(`Final @ M${T}`, "finalValue", curveParams.finalValue ?? 12000);
      break;
    case "milestone":
    default: {
      const ms = curveParams.milestones ?? [
        { month: Math.round(T * 0.33), value: 500, label: "" },
        { month: Math.round(T * 0.66), value: 5000, label: "" },
        { month: T, value: 15000, label: "" },
      ];
      curveParams.milestones = ms;

      const textPanel = document.createElement("div");
      textPanel.className = "milestone-text-panel";
      textPanel.innerHTML = `
        <label class="milestone-text-label">
          Describe your growth goals
          <textarea id="milestoneFreeText" rows="3" placeholder="e.g. Launch at month 6 with 100 users, hit 500 by end of year 1, reach 5k by month 36"></textarea>
        </label>
        <div class="milestone-parse-row">
          <button type="button" class="btn-secondary" id="milestoneParseBtn">Parse into milestones</button>
          <span id="milestoneParseStatus" class="status-text"></span>
        </div>
        <p class="hint milestone-ai-hint">${MILESTONE_PARSE_URL ? "Describe your goals in plain English — we'll parse them with AI." : "Describe your goals in plain English — parsed locally in your browser."}</p>`;
      form.appendChild(textPanel);

      const hint = document.createElement("p");
      hint.className = "hint milestone-form-hint";
      hint.textContent =
        "Or set milestones manually below. Add labels like “Launch” or “Series A” — they appear on the chart.";
      form.appendChild(hint);

      ms.forEach((m, i) => {
        const block = document.createElement("div");
        block.className = "milestone-block";
        block.dataset.index = String(i);

        const head = document.createElement("div");
        head.className = "milestone-block-header";
        head.innerHTML = `<span class="milestone-block-title">Milestone ${i + 1}</span>`;
        if (ms.length > 1) {
          const delBtn = document.createElement("button");
          delBtn.type = "button";
          delBtn.className = "btn-link milestone-delete-btn";
          delBtn.textContent = "Remove";
          delBtn.addEventListener("click", () => {
            syncParamsFromForm();
            curveParams.milestones.splice(i, 1);
            renderCurveForm();
            updatePreviewChart();
          });
          head.appendChild(delBtn);
        }
        block.appendChild(head);
        form.appendChild(block);

        const addTextField = (label, key, value, placeholder = "") => {
          const wrap = document.createElement("label");
          wrap.innerHTML = `${label}<input type="text" data-key="${key}" value="${escapeAttr(value)}" placeholder="${escapeAttr(placeholder)}" />`;
          form.appendChild(wrap);
        };

        addTextField("Label", `ms_l_${i}`, m.label ?? "", "e.g. Product launch");
        addField(`Month (M1–M${T})`, `ms_m_${i}`, m.month, 1);
        addField("Cumulative target", `ms_v_${i}`, m.value);
      });

      const addBtn = document.createElement("button");
      addBtn.type = "button";
      addBtn.className = "btn-secondary add-milestone-btn";
      addBtn.textContent = "+ Add milestone";
      addBtn.disabled = ms.length >= 6;
      addBtn.addEventListener("click", () => {
        syncParamsFromForm();
        const last = curveParams.milestones.at(-1);
        curveParams.milestones.push({
          month: T,
          value: last?.value ?? 0,
          label: "",
        });
        renderCurveForm();
        updatePreviewChart();
      });
      form.appendChild(addBtn);
      break;
    }
  }

  form.querySelectorAll("input").forEach((inp) => {
    inp.addEventListener("input", () => {
      syncParamsFromForm();
      updatePreviewChart();
    });
  });
}

function syncParamsFromForm() {
  const form = document.getElementById("curveForm");
  const T = totalMonths;
  switch (selectedCurveType) {
    case "linear":
      curveParams.start = num(form, "start");
      curveParams.end = num(form, "end");
      break;
    case "s_curve":
      curveParams.early = num(form, "early");
      curveParams.mid = num(form, "mid");
      curveParams.finalValue = num(form, "finalValue");
      break;
    case "hockey_stick":
      curveParams.flatUntil = num(form, "flatUntil");
      curveParams.flatLevel = num(form, "flatLevel");
      curveParams.finalValue = num(form, "finalValue");
      break;
    case "delayed_launch":
      curveParams.launchMonth = num(form, "launchMonth");
      curveParams.finalValue = num(form, "finalValue");
      break;
    case "milestone":
    default: {
      const ms = curveParams.milestones ?? [];
      const count = ms.length || 3;
      curveParams.milestones = [];
      for (let i = 0; i < count; i++) {
        curveParams.milestones.push({
          month: Math.min(T, num(form, `ms_m_${i}`) || T),
          value: num(form, `ms_v_${i}`),
          label: text(form, `ms_l_${i}`),
        });
      }
      break;
    }
  }
}

function num(form, key) {
  const el = form.querySelector(`[data-key="${key}"]`);
  return el ? Number(el.value) || 0 : 0;
}

function text(form, key) {
  const el = form.querySelector(`[data-key="${key}"]`);
  return el ? el.value.trim() : "";
}

function updatePreviewChart() {
  syncParamsFromForm();
  const { knotMonths, knotCumulative } = generateFromType(
    selectedCurveType,
    totalMonths,
    curveParams
  );
  const derived = curveFromKnots(knotMonths, knotCumulative, totalMonths);
  const canvas = document.getElementById("previewChart");
  drawMiniChart(canvas, derived.cumulative, "#2f4f46", true);
}

document.getElementById("curveApplyBtn").addEventListener("click", () => {
  const err = document.getElementById("curveError");
  err.textContent = "";
  try {
    syncParamsFromForm();
    applyCurveToActive();
    enableStepsFrom("edit");
    setStep("edit");
    initEditStep();
  } catch (e) {
    err.textContent = e.message;
  }
});

document.getElementById("curveForm").addEventListener("click", async (e) => {
  if (e.target.id !== "milestoneParseBtn" || selectedCurveType !== "milestone") return;

  const form = document.getElementById("curveForm");
  const textEl = form.querySelector("#milestoneFreeText");
  const statusEl = form.querySelector("#milestoneParseStatus");
  const text = textEl?.value?.trim() ?? "";

  if (!text) {
    if (statusEl) statusEl.textContent = "Enter a description first.";
    return;
  }

  e.target.disabled = true;
  if (statusEl) statusEl.textContent = MILESTONE_PARSE_URL ? "Asking Claude…" : "Parsing…";

  try {
    const { milestones, source } = await parseMilestoneText(text, totalMonths);
    if (!milestones.length) {
      if (statusEl) {
        statusEl.textContent =
          "Couldn't find milestones — try month + number (e.g. “500 users by month 12”).";
      }
      return;
    }

    syncParamsFromForm();
    curveParams.milestones = milestones;
    renderCurveForm();
    updatePreviewChart();

    const refreshed = document.getElementById("curveForm");
    const newText = refreshed.querySelector("#milestoneFreeText");
    const newStatus = refreshed.querySelector("#milestoneParseStatus");
    if (newText) newText.value = text;
    if (newStatus) {
      newStatus.textContent =
        source === "ai"
          ? `Parsed ${milestones.length} milestone${milestones.length === 1 ? "" : "s"} with Claude Sonnet.`
          : `Parsed ${milestones.length} milestone${milestones.length === 1 ? "" : "s"} locally.`;
    }
  } catch (err) {
    if (statusEl) statusEl.textContent = err?.message || "Parse failed.";
  } finally {
    const btn = document.getElementById("milestoneParseBtn");
    if (btn) btn.disabled = false;
  }
});

function applyCurveToActive() {
  const result = generateFromType(
    selectedCurveType,
    totalMonths,
    curveParams
  );
  const active = getActiveScenario();
  active.knotMonths = result.knotMonths;
  active.knotCumulative = result.knotCumulative;
  active.knotLabels = result.knotLabels ?? [];
  const derived = curveFromKnots(result.knotMonths, result.knotCumulative, totalMonths);
  active.flow = derived.signups;
}

// --- Edit ---

function initEditStep() {
  const editorScenarios = {};
  scenarios.forEach((s) => {
    editorScenarios[s.id] = {
      label: s.label,
      knotMonths: [...s.knotMonths],
      knotCumulative: [...s.knotCumulative],
      knotLabels: [...(s.knotLabels || [])],
      flow: [...s.flow],
    };
  });
  editor.setScenarios(editorScenarios, totalMonths, activeScenarioId, {
    allowKnotMonthEdit: selectedCurveType === "milestone",
  });
  renderCaseGrid();
  renderKnotTable();
  renderStats();
  document.getElementById("activeCaseLabel").textContent =
    getActiveScenario()?.label ?? "—";
  document.getElementById("editHint").textContent =
    selectedCurveType === "milestone"
      ? "Milestone mode: drag knots left/right to change timing, up/down for value. Labels show above each knot."
      : "Drag yellow knots on the cumulative curve. Monthly flow (pink bars) is derived automatically.";
}

function renderCaseGrid() {
  const grid = document.getElementById("caseGrid");
  grid.innerHTML = "";
  const data = editor.getScenarios();

  scenarios.forEach((s, i) => {
    const c = data[s.id];
    const card = document.createElement("button");
    card.type = "button";
    card.className = `case-card${s.id === activeScenarioId ? " active" : ""}`;
    card.innerHTML = `
      <div class="case-card-head">
        <span class="case-dot" style="background:${CASE_COLORS[i % CASE_COLORS.length]}"></span>
        <span class="case-name">${escapeHtml(s.label)}</span>
        <span class="case-range">${escapeHtml(s.rangeRef || "demo")}</span>
      </div>
      <canvas class="mini-chart" data-id="${s.id}" width="220" height="64"></canvas>
      <div class="case-stats">
        <span>Mid <b>${fmt(c?.stats?.p50)}</b></span>
        <span>Final <b>${fmt(c?.stats?.final)}</b></span>
      </div>`;
    card.addEventListener("click", () => {
      activeScenarioId = s.id;
      editor.selectScenario(s.id);
      document.getElementById("activeCaseLabel").textContent = s.label;
      renderCaseGrid();
      renderKnotTable();
      renderStats();
    });
    grid.appendChild(card);
  });

  document.querySelectorAll(".mini-chart").forEach((mini) => {
    const c = data[mini.dataset.id];
    const idx = scenarios.findIndex((s) => s.id === mini.dataset.id);
    if (c?.derived) {
      drawMiniChart(
        mini,
        c.derived.cumulative,
        CASE_COLORS[idx % CASE_COLORS.length],
        mini.dataset.id === activeScenarioId
      );
    }
  });
}

function renderKnotTable() {
  const table = document.getElementById("knotTable");
  const thead = table.querySelector("thead tr");
  const tbody = table.querySelector("tbody");
  const s = getActiveScenario();
  const data = editor.getScenarios()[activeScenarioId];
  const isMilestone = selectedCurveType === "milestone";
  if (!s || !data) return;

  thead.innerHTML = isMilestone
    ? "<th>Month</th><th>Label</th><th>Cumulative</th><th></th>"
    : "<th>Month</th><th>Cumulative</th>";

  tbody.innerHTML = "";
  const canDeleteKnot = isMilestone && s.knotMonths.length > 2;
  s.knotMonths.forEach((m, i) => {
    const tr = document.createElement("tr");
    const labelVal = escapeAttr(data.knotLabels?.[i] ?? "");
    if (isMilestone) {
      tr.innerHTML = `
        <td><input type="number" min="1" max="${totalMonths}" step="1" data-field="month" data-i="${i}" value="${m}" /></td>
        <td><input type="text" data-field="label" data-i="${i}" value="${labelVal}" placeholder="Milestone label" /></td>
        <td><input type="number" min="0" step="100" data-field="value" data-i="${i}" value="${data.knotCumulative[i]}" /></td>
        <td class="knot-actions">${canDeleteKnot ? `<button type="button" class="btn-link knot-delete-btn" data-i="${i}">Remove</button>` : ""}</td>`;
    } else {
      tr.innerHTML = `<td>M${m}</td><td><input type="number" min="0" step="100" data-field="value" data-i="${i}" value="${data.knotCumulative[i]}" /></td>`;
    }
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll("input").forEach((inp) => {
    inp.addEventListener("change", () => {
      const i = Number(inp.dataset.i);
      if (inp.dataset.field === "month") {
        editor.updateKnotMonth(i, inp.value);
      } else if (inp.dataset.field === "label") {
        editor.updateKnotLabel(i, inp.value);
      } else {
        editor.updateKnot(i, inp.value);
      }
      const updated = editor.getScenarios()[activeScenarioId];
      s.knotMonths = [...updated.knotMonths];
      s.knotCumulative = [...updated.knotCumulative];
      s.knotLabels = updated.knotLabels ? [...updated.knotLabels] : [];
      s.flow = [...updated.flow];
      if (inp.dataset.field === "month") renderKnotTable();
    });
  });

  tbody.querySelectorAll(".knot-delete-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const i = Number(btn.dataset.i);
      editor.deleteKnot(i);
      const updated = editor.getScenarios()[activeScenarioId];
      s.knotMonths = [...updated.knotMonths];
      s.knotCumulative = [...updated.knotCumulative];
      s.knotLabels = updated.knotLabels ? [...updated.knotLabels] : [];
      s.flow = [...updated.flow];
      curveParams.milestones = s.knotMonths.map((month, idx) => ({
        month,
        value: s.knotCumulative[idx] ?? 0,
        label: s.knotLabels?.[idx] ?? "",
      }));
      renderKnotTable();
      renderCaseGrid();
      renderStats();
    });
  });
}

function renderStats() {
  const grid = document.getElementById("statsGrid");
  const data = editor.getScenarios()[activeScenarioId];
  if (!data?.derived) return;
  const d = data.derived;
  const q1 = Math.floor(totalMonths * 0.25);
  const q2 = Math.floor(totalMonths * 0.5);
  grid.innerHTML = `
    <div class="stat"><span class="label">M${q1}</span><span class="value">${fmt(d.cumulative[q1 - 1])}</span></div>
    <div class="stat"><span class="label">M${q2}</span><span class="value">${fmt(d.cumulative[q2 - 1])}</span></div>
    <div class="stat"><span class="label">M${totalMonths}</span><span class="value">${fmt(d.cumulative[totalMonths - 1])}</span></div>
    <div class="stat"><span class="label">Total ${escapeHtml(metricLabel)}</span><span class="value">${fmt(d.cumulative[totalMonths - 1])}</span></div>`;
}

document.getElementById("reopenCurveBtn").addEventListener("click", () => {
  const s = getActiveScenario();
  if (s) {
    const inferred = inferCurveType(flowToCumulative(s.flow));
    selectedCurveType = inferred.type;
    curveParams = { ...inferred.params };
    if (selectedCurveType === "milestone") {
      curveParams.milestones = s.knotMonths.map((month, i) => ({
        month,
        value: s.knotCumulative[i] ?? 0,
        label: s.knotLabels?.[i] ?? "",
      }));
    }
  }
  initCurveStep();
  setStep("curve");
});

document.getElementById("editContinueBtn").addEventListener("click", () => {
  scenarios.forEach((s) => {
    const c = editor.getScenarios()[s.id];
    if (c) s.flow = [...c.flow];
  });
  initExportStep();
  enableStepsFrom("export");
  setStep("export");
});

// --- Export ---

function initExportStep() {
  const ul = document.getElementById("exportSummary");
  ul.innerHTML = scenarios
    .map(
      (s) =>
        `<li><strong>${escapeHtml(s.label)}</strong> — ${escapeHtml(s.sheet || "demo")} ${escapeHtml(s.rangeRef || "")} (${s.flow.length} periods, final cumulative ${fmt(flowToCumulative(s.flow).at(-1))})</li>`
    )
    .join("");
}

document.getElementById("downloadBtn").addEventListener("click", async () => {
  if (demoMode || !workbook || !originalWorkbookBuffer) {
    alert("Demo mode: upload a real .xlsx to download an updated workbook.");
    return;
  }
  const updates = scenarios.map((s) => ({
    sheet: s.sheet,
    rangeRef: s.rangeRef,
    values: s.flow,
  }));
  try {
    await downloadUpdatedWorkbook(originalWorkbookBuffer, updates, workbookName);
  } catch (err) {
    alert(err?.message || "Export failed. Try again.");
  }
});

document.getElementById("startOverBtn").addEventListener("click", () => {
  workbook = null;
  originalWorkbookBuffer = null;
  scenarios = [];
  demoMode = false;
  fileInput.value = "";
  uploadStatus.textContent = "";
  setStep("upload");
  document.querySelectorAll(".step-tab").forEach((btn, i) => {
    btn.disabled = i > 0;
  });
});

// --- Nav ---

document.querySelectorAll("[data-goto]").forEach((btn) => {
  btn.addEventListener("click", () => setStep(btn.dataset.goto));
});

document.querySelectorAll(".step-tab").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (!btn.disabled) setStep(btn.dataset.step);
  });
});

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeAttr(s) {
  return escapeHtml(s).replace(/'/g, "&#39;");
}

function fmt(n) {
  return (n ?? 0).toLocaleString();
}
