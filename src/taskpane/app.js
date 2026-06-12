// Taskpane wiring: connects the UI to the chart engines and Office.js layers.

/* global Office */

import { toWaterfall, toMatrix, toGantt } from "../lib/parse.js";
import {
  layoutWaterfall, layoutStacked, layoutClustered, layoutMekko, layoutGantt
} from "../lib/chartmath.js";
import { primsToSvg } from "../lib/svgpreview.js";
import { SCHEMES } from "../lib/palette.js";
import { insertPrimitives, updateChart, newChartId } from "../office/render.js";
import { createGrid } from "./grid.js";
import { readSlideTitles, insertAgendaSlide } from "../office/agenda.js";
import {
  insertHarveyBall, insertCheckbox, insertProcessFlow,
  insertTrafficLight, insertStamp
} from "../office/elements.js";
import { applyLayout } from "../office/layout.js";

// Default chart frame on a 16:9 slide (960x540pt), leaving room for a title.
const CHART_FRAME = { x: 120, y: 110, w: 720, h: 360 };
const PREVIEW_BOX = { x: 100, y: 95, w: 760, h: 390 };

const EXAMPLES = {
  waterfall: [
    ["Label", "Value"],
    ["2024 Revenue", "820"], ["Volume", "95"], ["Price", "40"],
    ["Churn", "-60"], ["FX", "-25"], ["2025 Revenue", "e"]
  ],
  stacked: [
    ["", "2022", "2023", "2024", "2025"],
    ["Americas", "120", "135", "150", "170"],
    ["EMEA", "90", "95", "100", "110"],
    ["APAC", "45", "55", "70", "85"]
  ],
  mekko: [
    ["", "Segment A", "Segment B", "Segment C"],
    ["Us", "40", "25", "10"],
    ["Comp 1", "30", "45", "15"],
    ["Comp 2", "30", "30", "75"]
  ],
  gantt: [
    ["Task", "Start", "End"],
    ["Discovery", "1", "3"], ["Design", "2", "5"],
    ["Build", "4", "10"], ["Test", "9", "12"], ["Launch", "12", "12"]
  ]
};
EXAMPLES.stacked100 = EXAMPLES.stacked;
EXAMPLES.clustered = EXAMPLES.stacked;

const FORMAT_HINTS = {
  waterfall:
    'Two columns: label, value. Use "e" (or "total"/"=") in the value column for a computed subtotal/total bar.',
  stacked:
    "Matrix: first row = categories (leave the corner cell empty), each following row = series name + values. Totals are computed for you.",
  stacked100:
    "Matrix: first row = categories, rows = series. Columns are normalized to 100%.",
  clustered:
    "Matrix: first row = categories, rows = series. Bars are drawn side by side.",
  mekko:
    "Matrix: first row = categories, rows = series. Column width is proportional to the column total.",
  gantt:
    "Three columns: task, start, end (numeric units such as weeks). Equal start/end renders a milestone."
};

const $ = (id) => document.getElementById(id);
let grid;
let currentChartId = null; // last chart inserted this session -> Update mode

Office.onReady((info) => {
  if (info.host && info.host !== Office.HostType.PowerPoint) {
    setStatus("This add-in is designed for PowerPoint.", true);
  }
  wireTabs();
  wireCharts();
  wireAgenda();
  wireElements();
  wireLayout();
});

// ------------------------------------------------------------------ tabs

function wireTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      tab.classList.add("active");
      $(`panel-${tab.dataset.tab}`).classList.add("active");
      setStatus("");
    });
  });
}

// ----------------------------------------------------------------- charts

function wireCharts() {
  grid = createGrid($("datasheet"), { onChange: refreshPreview });
  grid.setData(EXAMPLES.stacked);

  $("chart-type").addEventListener("change", () => {
    updateFormatHint();
    resetSync();
    refreshPreview();
  });
  ["opt-labels", "opt-totals", "opt-axis", "opt-cagr", "opt-diff", "opt-mean",
   "opt-decimals", "opt-scheme"]
    .forEach((id) => $(id).addEventListener("change", refreshPreview));

  $("btn-transpose").addEventListener("click", () => {
    const rows = grid.getData();
    if (!rows.length) return;
    const cols = Math.max(...rows.map((r) => r.length));
    grid.setData(Array.from({ length: cols }, (_, c) =>
      Array.from({ length: rows.length }, (_, r) => rows[r][c] ?? "")));
    setStatus("Rows and columns swapped.");
  });

  $("btn-example").addEventListener("click", () => {
    grid.setData(EXAMPLES[$("chart-type").value]);
    setStatus("Example loaded — edit the datasheet or press Insert.");
  });
  $("btn-clear").addEventListener("click", () => {
    grid.clear();
    resetSync();
  });

  $("btn-insert-chart").addEventListener("click", () => run(insertOrUpdate));
  $("btn-insert-new").addEventListener("click", () => run(() => doInsert(true)));

  updateFormatHint();
  refreshPreview();
}

function buildPrims(frame) {
  const type = $("chart-type").value;
  const rows = grid.getData();
  if (!rows.length) throw new Error("The datasheet is empty.");
  const decRaw = $("opt-decimals").value;
  const o = {
    showLabels: $("opt-labels").checked,
    showTotals: $("opt-totals").checked,
    axis: $("opt-axis").checked,
    cagr: $("opt-cagr").checked,
    diff: $("opt-diff").checked,
    meanLine: $("opt-mean").checked,
    decimals: decRaw === "" ? undefined : Number(decRaw),
    palette: SCHEMES[$("opt-scheme").value] || null
  };
  if (type === "waterfall") return layoutWaterfall(toWaterfall(rows), frame, o);
  if (type === "stacked") return layoutStacked(toMatrix(rows), frame, o);
  if (type === "stacked100") return layoutStacked(toMatrix(rows), frame, { ...o, percent: true });
  if (type === "clustered") return layoutClustered(toMatrix(rows), frame, o);
  if (type === "mekko") return layoutMekko(toMatrix(rows), frame, o);
  if (type === "gantt") return layoutGantt(toGantt(rows), frame, o);
  throw new Error(`Unknown chart type: ${type}`);
}

function refreshPreview() {
  const el = $("chart-preview");
  try {
    const prims = buildPrims(CHART_FRAME);
    el.innerHTML = primsToSvg(prims, PREVIEW_BOX);
  } catch (e) {
    el.innerHTML = `<div class="empty">${escapeHtml(e.message)}</div>`;
  }
}

async function insertOrUpdate() {
  if (currentChartId) {
    const prims = buildPrims(CHART_FRAME);
    const replaced = await updateChart(currentChartId, prims);
    setStatus(replaced
      ? "Chart updated on the slide."
      : "Previous chart not found on this slide — inserted fresh.");
  } else {
    await doInsert(false);
  }
}

async function doInsert(asNew) {
  const prims = buildPrims(CHART_FRAME);
  currentChartId = newChartId();
  await insertPrimitives(prims, { chartId: currentChartId });
  enterSyncMode();
  setStatus(asNew ? "New chart inserted." : "Chart inserted — edits here now update it in place.");
}

function enterSyncMode() {
  $("btn-insert-chart").textContent = "Update chart on slide";
  $("btn-insert-new").hidden = false;
}

function resetSync() {
  currentChartId = null;
  $("btn-insert-chart").textContent = "Insert chart on current slide";
  $("btn-insert-new").hidden = true;
}

function updateFormatHint() {
  $("format-hint").textContent = FORMAT_HINTS[$("chart-type").value];
}

// ----------------------------------------------------------------- agenda

function wireAgenda() {
  $("btn-scan-titles").addEventListener("click", () =>
    run(async () => {
      const titles = await readSlideTitles();
      $("agenda-items").value = titles.map((t) => t.title).join("\n");
    }, "Titles scanned — edit the list, then insert.")
  );
  $("btn-insert-agenda").addEventListener("click", () =>
    run(async () => {
      const items = $("agenda-items").value
        .split("\n").map((s) => s.trim()).filter(Boolean);
      if (!items.length) throw new Error("Add at least one agenda item.");
      await insertAgendaSlide(items, { heading: $("agenda-heading").value || "Agenda" });
    }, "Agenda slide added at the end of the deck.")
  );
}

// --------------------------------------------------------------- elements

function wireElements() {
  document.querySelectorAll("[data-harvey]").forEach((b) =>
    b.addEventListener("click", () =>
      run(() => insertHarveyBall(Number(b.dataset.harvey)), "Harvey ball inserted."))
  );
  document.querySelectorAll("[data-check]").forEach((b) =>
    b.addEventListener("click", () =>
      run(() => insertCheckbox(b.dataset.check), "Checkbox inserted."))
  );
  document.querySelectorAll("[data-light]").forEach((b) =>
    b.addEventListener("click", () =>
      run(() => insertTrafficLight(b.dataset.light), "Traffic light inserted."))
  );
  document.querySelectorAll("[data-stamp]").forEach((b) =>
    b.addEventListener("click", () =>
      run(() => insertStamp(b.dataset.stamp), "Stamp inserted."))
  );
  $("btn-flow").addEventListener("click", () =>
    run(() => {
      const steps = $("flow-steps").value.split(",").map((s) => s.trim()).filter(Boolean);
      return insertProcessFlow(steps);
    }, "Process flow inserted.")
  );
}

// ----------------------------------------------------------------- layout

function wireLayout() {
  document.querySelectorAll("[data-layout]").forEach((b) =>
    b.addEventListener("click", () =>
      run(() => applyLayout(b.dataset.layout), "Done."))
  );
}

// ------------------------------------------------------------------ misc

async function run(fn, okMessage) {
  try {
    setStatus("Working…");
    await fn();
    if (okMessage) setStatus(okMessage);
  } catch (e) {
    const msg = e?.debugInfo?.message || e?.message || String(e);
    setStatus(msg, true);
    console.error(e);
  }
}

function setStatus(msg, isError = false) {
  const el = $("status");
  el.textContent = msg;
  el.className = `status ${isError ? "err" : msg ? "ok" : ""}`;
}

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
