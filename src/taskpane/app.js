// Taskpane wiring: connects the UI to the chart engines and Office.js layers.

/* global Office */

import { parseTable, toWaterfall, toMatrix, toGantt } from "../lib/parse.js";
import { layoutWaterfall, layoutStacked, layoutMekko, layoutGantt } from "../lib/chartmath.js";
import { insertPrimitives } from "../office/render.js";
import { readSlideTitles, insertAgendaSlide } from "../office/agenda.js";
import {
  insertHarveyBall, insertCheckbox, insertProcessFlow,
  insertTrafficLight, insertStamp
} from "../office/elements.js";
import { applyLayout } from "../office/layout.js";

// Default chart frame on a 16:9 slide (960x540pt), leaving room for a title.
const CHART_FRAME = { x: 120, y: 110, w: 720, h: 360 };

const EXAMPLES = {
  waterfall:
    "Label\tValue\n2024 Revenue\t820\nVolume\t95\nPrice\t40\nChurn\t-60\nFX\t-25\n2025 Revenue\te",
  stacked:
    "\tQ1\tQ2\tQ3\tQ4\nAmericas\t120\t135\t150\t170\nEMEA\t90\t95\t100\t110\nAPAC\t45\t55\t70\t85",
  stacked100:
    "\tQ1\tQ2\tQ3\tQ4\nAmericas\t120\t135\t150\t170\nEMEA\t90\t95\t100\t110\nAPAC\t45\t55\t70\t85",
  mekko:
    "\tSegment A\tSegment B\tSegment C\nUs\t40\t25\t10\nComp 1\t30\t45\t15\nComp 2\t30\t30\t75",
  gantt:
    "Task\tStart\tEnd\nDiscovery\t1\t3\nDesign\t2\t5\nBuild\t4\t10\nTest\t9\t12\nLaunch\t12\t12"
};

const FORMAT_HINTS = {
  waterfall:
    'Two columns: label, value. Use "e" (or "total"/"=") in the value column for a computed subtotal/total bar.',
  stacked:
    "Matrix: first row = categories (leave the corner cell empty), each following row = series name + values.",
  stacked100:
    "Matrix: first row = categories, each following row = series name + values. Columns are normalized to 100%.",
  mekko:
    "Matrix: first row = categories, rows = series. Column width is proportional to the column total.",
  gantt:
    "Three columns: task, start, end (numeric units such as weeks). Equal start/end renders a milestone."
};

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  if (info.host && info.host !== Office.HostType.PowerPoint) {
    setStatus("This add-in is designed for PowerPoint.", true);
  }
  wireTabs();
  wireCharts();
  wireAgenda();
  wireElements();
  wireLayout();
  updateFormatHint();
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
  $("chart-type").addEventListener("change", updateFormatHint);
  $("btn-example").addEventListener("click", () => {
    $("chart-data").value = EXAMPLES[$("chart-type").value];
    setStatus("Example loaded — press Insert.");
  });
  $("btn-insert-chart").addEventListener("click", () => run(insertChart, "Chart inserted."));
}

function updateFormatHint() {
  $("format-hint").textContent = FORMAT_HINTS[$("chart-type").value];
}

async function insertChart() {
  const type = $("chart-type").value;
  const raw = $("chart-data").value;
  if (!raw.trim()) throw new Error("Paste some data first (or click Load example).");
  const rows = parseTable(raw);
  const showLabels = $("opt-labels").checked;
  const showTotals = $("opt-totals").checked;

  let prims;
  if (type === "waterfall") {
    prims = layoutWaterfall(toWaterfall(rows), CHART_FRAME, { showLabels });
  } else if (type === "stacked" || type === "stacked100") {
    prims = layoutStacked(toMatrix(rows), CHART_FRAME, {
      percent: type === "stacked100", showLabels, showTotals
    });
  } else if (type === "mekko") {
    prims = layoutMekko(toMatrix(rows), CHART_FRAME, { showLabels, showTotals });
  } else if (type === "gantt") {
    prims = layoutGantt(toGantt(rows), CHART_FRAME);
  } else {
    throw new Error(`Unknown chart type: ${type}`);
  }
  await insertPrimitives(prims, { namePrefix: `SlideCharts ${type}` });
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
    setStatus(okMessage || "Done.");
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
