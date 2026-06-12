// Pure chart layout engines. Each layout* function converts data + a frame
// (in points: {x, y, w, h}) into a flat list of drawing primitives:
//
//   {kind:'rect', x,y,w,h, fill, line?, text?, fontSize?, fontColor?, bold?}
//   {kind:'line', x1,y1,x2,y2, color, weight, dash?}   (axis-parallel only)
//   {kind:'text', x,y,w,h, text, fontSize?, fontColor?, bold?, align?}
//
// Keeping these free of Office.js makes them unit-testable with plain node.

import { fmt, pct } from "./format.js";
import { COLORS, seriesColor, contrastText } from "./palette.js";

const LABEL_H = 14; // height of a value-label text box
const FONT = 10;

// ---------------------------------------------------------------- waterfall

export function layoutWaterfall(rows, frame, opts = {}) {
  const o = {
    gap: 0.4,
    posColor: COLORS.positive,
    negColor: COLORS.negative,
    totalColor: COLORS.total,
    showLabels: true,
    ...opts
  };
  if (!rows.length) throw new Error("Waterfall needs at least one row");

  // Resolve cumulative levels; "e" rows become totals spanning 0..running.
  let run = 0;
  const levels = [0];
  const bars = rows.map((r) => {
    if (r.value === "e") {
      levels.push(run);
      return { label: r.label, from: 0, to: run, total: true, value: run };
    }
    const from = run;
    run += r.value;
    levels.push(run);
    return { label: r.label, from, to: run, total: false, value: r.value };
  });

  const vMax = Math.max(...levels, 0);
  const vMin = Math.min(...levels, 0);
  const topPad = LABEL_H + 4;
  const botPad = LABEL_H + 6 + (vMin < 0 ? LABEL_H : 0);
  const plotH = frame.h - topPad - botPad;
  const span = vMax - vMin || 1;
  const yOf = (v) => frame.y + topPad + ((vMax - v) * plotH) / span;

  const slot = frame.w / bars.length;
  const barW = slot * (1 - o.gap);
  const prims = [];

  bars.forEach((b, i) => {
    const x = frame.x + slot * i + (slot - barW) / 2;
    const yTop = yOf(Math.max(b.from, b.to));
    const h = Math.max(Math.abs(yOf(b.from) - yOf(b.to)), 1);
    const fill = b.total ? o.totalColor : b.value >= 0 ? o.posColor : o.negColor;
    prims.push({ kind: "rect", x, y: yTop, w: barW, h, fill, meta: { role: "bar", from: b.from, to: b.to } });

    if (o.showLabels) {
      const above = b.total || b.value >= 0;
      const ly = above ? yTop - LABEL_H : yTop + h + 2;
      prims.push({
        kind: "text", x: x - slot * 0.2, y: ly, w: barW + slot * 0.4, h: LABEL_H,
        text: fmt(b.value), fontSize: FONT, bold: b.total, align: "center",
        meta: { role: "value" }
      });
    }
    prims.push({
      kind: "text", x: frame.x + slot * i, y: frame.y + frame.h - LABEL_H,
      w: slot, h: LABEL_H, text: b.label, fontSize: FONT, align: "center",
      meta: { role: "category" }
    });

    if (i < bars.length - 1) {
      const cy = yOf(b.to);
      prims.push({
        kind: "line", x1: x + barW, y1: cy,
        x2: frame.x + slot * (i + 1) + (slot - barW) / 2, y2: cy,
        color: COLORS.connector, weight: 0.75, dash: "dash",
        meta: { role: "connector", level: b.to }
      });
    }
  });

  // Baseline at value 0 sits on top so it is always visible.
  prims.push({
    kind: "line", x1: frame.x, y1: yOf(0), x2: frame.x + frame.w, y2: yOf(0),
    color: COLORS.axis, weight: 1, meta: { role: "baseline" }
  });
  return prims;
}

// ------------------------------------------------------- stacked column

export function layoutStacked(data, frame, opts = {}) {
  const o = { percent: false, gap: 0.35, showTotals: true, showLabels: true, showLegend: true, ...opts };
  const { categories, series } = data;
  const n = categories.length;
  if (!n || !series.length) throw new Error("Stacked chart needs categories and series");

  const totals = categories.map((_, c) => series.reduce((s, sr) => s + (sr.values[c] || 0), 0));
  const vMax = o.percent ? 100 : Math.max(...totals, 1);

  const topPad = o.showTotals && !o.percent ? LABEL_H + 4 : 6;
  const legendH = o.showLegend ? LABEL_H + 4 : 0;
  const botPad = LABEL_H + 4 + legendH;
  const plotH = frame.h - topPad - botPad;
  const plotBottom = frame.y + topPad + plotH;

  const slot = frame.w / n;
  const barW = slot * (1 - o.gap);
  const prims = [];

  prims.push({
    kind: "line", x1: frame.x, y1: plotBottom, x2: frame.x + frame.w, y2: plotBottom,
    color: COLORS.axis, weight: 1, meta: { role: "baseline" }
  });

  categories.forEach((cat, c) => {
    const x = frame.x + slot * c + (slot - barW) / 2;
    let cursor = plotBottom;
    series.forEach((sr, s) => {
      const raw = sr.values[c] || 0;
      const val = o.percent ? (totals[c] ? (raw / totals[c]) * 100 : 0) : raw;
      const h = (val * plotH) / vMax;
      if (h <= 0) return;
      const fill = sr.color || seriesColor(s);
      const seg = {
        kind: "rect", x, y: cursor - h, w: barW, h, fill,
        line: { color: "#FFFFFF", weight: 0.75 },
        meta: { role: "segment", series: sr.name, category: cat, value: raw }
      };
      if (o.showLabels && h >= 11) {
        seg.text = o.percent ? pct(val) : fmt(raw);
        seg.fontSize = FONT;
        seg.fontColor = contrastText(fill);
      }
      prims.push(seg);
      cursor -= h;
    });

    if (o.showTotals && !o.percent) {
      prims.push({
        kind: "text", x: frame.x + slot * c, y: cursor - LABEL_H, w: slot, h: LABEL_H,
        text: fmt(totals[c]), fontSize: FONT, bold: true, align: "center",
        meta: { role: "total" }
      });
    }
    prims.push({
      kind: "text", x: frame.x + slot * c, y: plotBottom + 3, w: slot, h: LABEL_H,
      text: cat, fontSize: FONT, align: "center", meta: { role: "category" }
    });
  });

  if (o.showLegend) prims.push(...legendRow(series, frame, frame.y + frame.h - LABEL_H));
  return prims;
}

// ------------------------------------------------------------------ mekko

export function layoutMekko(data, frame, opts = {}) {
  const o = { colGap: 2, showLabels: true, showTotals: true, showLegend: true, ...opts };
  const { categories, series } = data;
  const n = categories.length;
  if (!n || !series.length) throw new Error("Mekko chart needs categories and series");

  const totals = categories.map((_, c) => series.reduce((s, sr) => s + (sr.values[c] || 0), 0));
  const grand = totals.reduce((a, b) => a + b, 0);
  if (grand <= 0) throw new Error("Mekko chart needs positive totals");

  const topPad = o.showTotals ? LABEL_H + 4 : 6;
  const legendH = o.showLegend ? LABEL_H + 4 : 0;
  const botPad = LABEL_H + 4 + legendH;
  const plotH = frame.h - topPad - botPad;
  const plotTop = frame.y + topPad;
  const usableW = frame.w - o.colGap * (n - 1);

  const prims = [];
  let x = frame.x;
  categories.forEach((cat, c) => {
    const colW = (totals[c] / grand) * usableW;
    let cursor = plotTop + plotH;
    series.forEach((sr, s) => {
      const raw = sr.values[c] || 0;
      const share = totals[c] ? raw / totals[c] : 0;
      const h = share * plotH;
      if (h <= 0) return;
      const fill = sr.color || seriesColor(s);
      const seg = {
        kind: "rect", x, y: cursor - h, w: colW, h, fill,
        line: { color: "#FFFFFF", weight: 0.75 },
        meta: { role: "segment", series: sr.name, category: cat, value: raw, share }
      };
      if (o.showLabels && h >= 11 && colW >= 24) {
        seg.text = pct(share * 100);
        seg.fontSize = FONT;
        seg.fontColor = contrastText(fill);
      }
      prims.push(seg);
      cursor -= h;
    });

    if (o.showTotals) {
      prims.push({
        kind: "text", x, y: plotTop - LABEL_H - 2, w: colW, h: LABEL_H,
        text: fmt(totals[c]), fontSize: FONT, bold: true, align: "center",
        meta: { role: "total" }
      });
    }
    prims.push({
      kind: "text", x, y: plotTop + plotH + 3, w: colW, h: LABEL_H,
      text: cat, fontSize: FONT, align: "center", meta: { role: "category" }
    });
    x += colW + o.colGap;
  });

  if (o.showLegend) prims.push(...legendRow(series, frame, frame.y + frame.h - LABEL_H));
  return prims;
}

// ------------------------------------------------------------------ gantt

export function layoutGantt(tasks, frame, opts = {}) {
  const o = { unitLabel: "", barColor: COLORS.ganttBar, ...opts };
  if (!tasks.length) throw new Error("Gantt needs at least one task");

  const tMin = Math.floor(Math.min(...tasks.map((t) => t.start)));
  const tMax = Math.ceil(Math.max(...tasks.map((t) => t.end)));
  const units = Math.max(tMax - tMin, 1);

  const labelW = Math.min(140, frame.w * 0.25);
  const headerH = LABEL_H + 4;
  const gridX = frame.x + labelW;
  const gridW = frame.w - labelW;
  const rowH = Math.min((frame.h - headerH) / tasks.length, 30);
  const xOf = (t) => gridX + ((t - tMin) / units) * gridW;

  const prims = [];

  // Header tick labels + vertical grid lines per unit.
  for (let u = tMin; u <= tMax; u++) {
    const gx = xOf(u);
    prims.push({
      kind: "line", x1: gx, y1: frame.y + headerH, x2: gx,
      y2: frame.y + headerH + rowH * tasks.length,
      color: COLORS.grid, weight: 0.75, meta: { role: "grid", unit: u }
    });
    if (u < tMax) {
      prims.push({
        kind: "text", x: gx, y: frame.y, w: gridW / units, h: LABEL_H,
        text: `${o.unitLabel}${u}`, fontSize: FONT - 1, align: "center",
        fontColor: "#595959", meta: { role: "tick" }
      });
    }
  }

  tasks.forEach((t, i) => {
    const rowY = frame.y + headerH + rowH * i;
    if (i % 2 === 1) {
      prims.push({
        kind: "rect", x: frame.x, y: rowY, w: frame.w, h: rowH,
        fill: "#F2F2F2", meta: { role: "rowShade" }
      });
    }
    prims.push({
      kind: "text", x: frame.x, y: rowY + (rowH - LABEL_H) / 2, w: labelW - 6, h: LABEL_H,
      text: t.label, fontSize: FONT, align: "left", meta: { role: "task" }
    });
    const barH = Math.min(rowH * 0.55, 14);
    if (t.end > t.start) {
      prims.push({
        kind: "rect", x: xOf(t.start), y: rowY + (rowH - barH) / 2,
        w: xOf(t.end) - xOf(t.start), h: barH, fill: o.barColor,
        meta: { role: "bar", start: t.start, end: t.end }
      });
    } else {
      // Zero-duration task -> milestone diamond.
      const d = Math.min(rowH * 0.5, 12);
      prims.push({
        kind: "rect", shape: "diamond", x: xOf(t.start) - d / 2, y: rowY + (rowH - d) / 2,
        w: d, h: d, fill: COLORS.ganttMilestone, meta: { role: "milestone", at: t.start }
      });
    }
  });

  // Re-stack: shading first so bars/text drawn after stay on top.
  prims.sort((a, b) => z(a) - z(b));
  return prims;
}

function z(p) {
  const order = { rowShade: 0, grid: 1 };
  return order[p.meta?.role] ?? 2;
}

// ------------------------------------------------------------------ legend

function legendRow(series, frame, y) {
  const prims = [];
  const itemW = Math.min(frame.w / series.length, 120);
  const totalW = itemW * series.length;
  let x = frame.x + (frame.w - totalW) / 2;
  series.forEach((sr, s) => {
    prims.push({
      kind: "rect", x, y: y + 3, w: 8, h: 8,
      fill: sr.color || seriesColor(s), meta: { role: "legendSwatch" }
    });
    prims.push({
      kind: "text", x: x + 11, y, w: itemW - 14, h: LABEL_H,
      text: sr.name, fontSize: FONT - 1, align: "left", meta: { role: "legendLabel" }
    });
    x += itemW;
  });
  return prims;
}
