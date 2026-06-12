// Pure chart layout engines. Each layout* function converts data + a frame
// (in points: {x, y, w, h}) into a flat list of drawing primitives:
//
//   {kind:'rect', x,y,w,h, fill, line?, text?, fontSize?, fontColor?, bold?}
//   {kind:'line', x1,y1,x2,y2, color, weight, dash?}   (axis-parallel only)
//   {kind:'text', x,y,w,h, text, fontSize?, fontColor?, bold?, align?}
//   {kind:'arrow', dir:'right'|'upDown', x,y,w,h, fill}
//
// Keeping these free of Office.js makes them unit-testable with plain node.

import { fmt, pct } from "./format.js";
import { COLORS, seriesColor, contrastText } from "./palette.js";

const LABEL_H = 14; // height of a value-label text box
const FONT = 10;
const AXIS_W = 40;   // left gutter when a value axis is shown
const DIFF_W = 54;   // right gutter when a difference arrow is shown
const CAGR_H = 30;   // extra top space when a CAGR arrow is shown

// "Nice" tick steps for a value axis covering [min, max].
export function niceTicks(min, max, target = 5) {
  const lo = Math.min(min, 0);
  const hi = Math.max(max, 0);
  const span = hi - lo || 1;
  const raw = span / target;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const ticks = [];
  for (let v = Math.ceil(lo / step) * step; v <= hi + step / 1e6; v += step) {
    ticks.push(Math.abs(v) < step / 1e6 ? 0 : Number(v.toFixed(10)));
  }
  return { step, ticks };
}

// Compound annual growth rate between the first and last of n values.
export function cagr(first, last, periods) {
  if (first <= 0 || last <= 0 || periods < 1) return null;
  return Math.pow(last / first, 1 / periods) - 1;
}

function pick(o, sr, s) {
  if (sr.color) return sr.color;
  if (o.palette) return o.palette[s % o.palette.length];
  return seriesColor(s);
}

// White ellipse bubble with black outline, as think-cell draws on its CAGR
// and difference arrows.
function bubble(cx, cy, text, role, extra = {}) {
  const w = Math.max(34, text.length * 5.6 + 14);
  return {
    kind: "rect", shape: "ellipse", x: cx - w / 2, y: cy - 9, w, h: 18,
    fill: "#FFFFFF", line: { color: "#000000", weight: 1 },
    text, fontSize: FONT - 1, bold: true, fontColor: "#000000",
    meta: { role, ...extra }
  };
}

function axisPrims(ticks, yOf, x0, x1, decimals) {
  const prims = [];
  for (const t of ticks) {
    const y = yOf(t);
    prims.push({
      kind: "line", x1: x0, y1: y, x2: x1, y2: y,
      color: COLORS.grid, weight: 0.75, meta: { role: "gridline", value: t }
    });
    prims.push({
      kind: "text", x: x0 - AXIS_W, y: y - LABEL_H / 2, w: AXIS_W - 6, h: LABEL_H,
      text: fmt(t, decimals), fontSize: FONT - 1, fontColor: "#595959",
      align: "right", meta: { role: "tickLabel", value: t }
    });
  }
  return prims;
}

// ---------------------------------------------------------------- waterfall

export function layoutWaterfall(rows, frame, opts = {}) {
  const o = {
    gap: 0.4,
    posColor: COLORS.positive,
    negColor: COLORS.negative,
    totalColor: COLORS.total,
    showLabels: true,
    axis: false,
    decimals: undefined,
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
  const axisW = o.axis ? AXIS_W : 0;
  const inner = { x: frame.x + axisW, w: frame.w - axisW };
  const topPad = LABEL_H + 4;
  const botPad = LABEL_H + 6 + (vMin < 0 ? LABEL_H : 0);
  const plotH = frame.h - topPad - botPad;
  const span = vMax - vMin || 1;
  const yOf = (v) => frame.y + topPad + ((vMax - v) * plotH) / span;

  const slot = inner.w / bars.length;
  const barW = slot * (1 - o.gap);
  const prims = [];

  if (o.axis) {
    const { ticks } = niceTicks(vMin, vMax);
    prims.push(...axisPrims(ticks, yOf, inner.x, inner.x + inner.w, o.decimals));
  }

  bars.forEach((b, i) => {
    const x = inner.x + slot * i + (slot - barW) / 2;
    const yTop = yOf(Math.max(b.from, b.to));
    const h = Math.max(Math.abs(yOf(b.from) - yOf(b.to)), 1);
    const fill = b.total ? o.totalColor : b.value >= 0 ? o.posColor : o.negColor;
    const bar = { kind: "rect", x, y: yTop, w: barW, h, fill, meta: { role: "bar", from: b.from, to: b.to } };

    if (o.showLabels && h >= 16 && !b.total) {
      // think-cell puts segment labels inside the bar when they fit.
      bar.text = fmt(b.value, o.decimals);
      bar.fontSize = FONT;
      bar.fontColor = contrastText(fill);
      prims.push(bar);
    } else {
      prims.push(bar);
      if (o.showLabels) {
        const above = b.total || b.value >= 0;
        const ly = above ? yTop - LABEL_H : yTop + h + 2;
        prims.push({
          kind: "text", x: x - slot * 0.2, y: ly, w: barW + slot * 0.4, h: LABEL_H,
          text: fmt(b.value, o.decimals), fontSize: FONT, bold: b.total, align: "center",
          meta: { role: "value" }
        });
      }
    }
    prims.push({
      kind: "text", x: inner.x + slot * i, y: frame.y + frame.h - LABEL_H,
      w: slot, h: LABEL_H, text: b.label, fontSize: FONT, align: "center",
      meta: { role: "category" }
    });

    if (i < bars.length - 1) {
      const cy = yOf(b.to);
      prims.push({
        kind: "line", x1: x + barW, y1: cy,
        x2: inner.x + slot * (i + 1) + (slot - barW) / 2, y2: cy,
        color: COLORS.connector, weight: 0.75, dash: "dash",
        meta: { role: "connector", level: b.to }
      });
    }
  });

  // Baseline at value 0 sits on top so it is always visible.
  prims.push({
    kind: "line", x1: inner.x, y1: yOf(0), x2: inner.x + inner.w, y2: yOf(0),
    color: COLORS.axis, weight: 1, meta: { role: "baseline" }
  });
  return prims;
}

// ------------------------------------------------------- stacked column

export function layoutStacked(data, frame, opts = {}) {
  const o = {
    percent: false, gap: 0.35, showTotals: true, showLabels: true,
    showLegend: true, axis: false, cagr: false, diff: false, meanLine: false,
    decimals: undefined, palette: null, ...opts
  };
  const { categories, series } = data;
  const n = categories.length;
  if (!n || !series.length) throw new Error("Stacked chart needs categories and series");

  // Net totals label the columns; positive/negative sums set the scale, since
  // negative segments stack below the baseline (as in think-cell).
  const totals = categories.map((_, c) => series.reduce((s, sr) => s + (sr.values[c] || 0), 0));
  const posSums = categories.map((_, c) =>
    series.reduce((s, sr) => s + Math.max(sr.values[c] || 0, 0), 0));
  const negSums = categories.map((_, c) =>
    series.reduce((s, sr) => s + Math.min(sr.values[c] || 0, 0), 0));
  const vMax = o.percent ? 100 : Math.max(...posSums, 1);
  const vMin = o.percent ? 0 : Math.min(...negSums, 0);
  const withCagr = o.cagr && !o.percent && n >= 2;
  const withDiff = o.diff && !o.percent && n >= 2;

  const axisW = o.axis ? AXIS_W : 0;
  const diffW = withDiff ? DIFF_W : 0;
  const inner = { x: frame.x + axisW, w: frame.w - axisW - diffW };

  const topPad = (o.showTotals && !o.percent ? LABEL_H + 4 : 6) + (withCagr ? CAGR_H : 0);
  const legendH = o.showLegend ? LABEL_H + 4 : 0;
  const botPad = LABEL_H + 4 + legendH;
  const plotH = frame.h - topPad - botPad;
  const span = vMax - vMin || 1;
  const yOf = (v) => frame.y + topPad + ((vMax - v) * plotH) / span;
  const plotBottom = frame.y + topPad + plotH;

  const slot = inner.w / n;
  const barW = slot * (1 - o.gap);
  const prims = [];

  if (o.axis) {
    const { ticks } = niceTicks(vMin, vMax);
    prims.push(...axisPrims(ticks, yOf, inner.x, inner.x + inner.w, o.decimals));
  }

  prims.push({
    kind: "line", x1: inner.x, y1: yOf(0), x2: inner.x + inner.w, y2: yOf(0),
    color: COLORS.axis, weight: 1, meta: { role: "baseline" }
  });

  categories.forEach((cat, c) => {
    const x = inner.x + slot * c + (slot - barW) / 2;
    let up = 0;   // cumulative positive value
    let down = 0; // cumulative negative value
    series.forEach((sr, s) => {
      const raw = sr.values[c] || 0;
      if (raw === 0) return;
      const val = o.percent ? (posSums[c] ? (Math.max(raw, 0) / posSums[c]) * 100 : 0) : raw;
      if (o.percent && val <= 0) return;
      let y, h;
      if (val >= 0) {
        h = (val * plotH) / span;
        up += val;
        y = yOf(up);
      } else {
        h = (-val * plotH) / span;
        y = yOf(down);
        down += val;
      }
      if (h <= 0) return;
      const fill = pick(o, sr, s);
      const seg = {
        kind: "rect", x, y, w: barW, h, fill,
        line: { color: "#FFFFFF", weight: 0.75 },
        meta: { role: "segment", series: sr.name, category: cat, value: raw }
      };
      if (o.showLabels && h >= 11) {
        seg.text = o.percent ? pct(val, o.decimals) : fmt(raw, o.decimals);
        seg.fontSize = FONT;
        seg.fontColor = contrastText(fill);
      }
      prims.push(seg);
    });

    if (o.showTotals && !o.percent) {
      prims.push({
        kind: "text", x: inner.x + slot * c, y: yOf(posSums[c]) - LABEL_H, w: slot, h: LABEL_H,
        text: fmt(totals[c], o.decimals), fontSize: FONT, bold: true, align: "center",
        meta: { role: "total" }
      });
    }
    prims.push({
      kind: "text", x: inner.x + slot * c, y: plotBottom + 3, w: slot, h: LABEL_H,
      text: cat, fontSize: FONT, align: "center", meta: { role: "category" }
    });
  });

  if (o.meanLine && !o.percent) {
    const mean = totals.reduce((a, b) => a + b, 0) / n;
    const y = yOf(mean);
    prims.push({
      kind: "line", x1: inner.x, y1: y, x2: inner.x + inner.w, y2: y,
      color: "#404040", weight: 1, dash: "dash", meta: { role: "valueLine", value: mean }
    });
    prims.push({
      kind: "text", x: inner.x + inner.w - 90, y: y - LABEL_H - 1, w: 90, h: LABEL_H,
      text: `Ø ${fmt(mean, o.decimals)}`, fontSize: FONT - 1, align: "right",
      meta: { role: "valueLineLabel", value: mean }
    });
  }
  if (withCagr) {
    prims.push(...cagrArrow(totals, n, slot, inner, frame, o.decimals));
  }
  if (withDiff) {
    prims.push(...diffArrow(totals, yOf, inner));
  }
  if (o.showLegend) prims.push(...legendRow(series, frame, frame.y + frame.h - LABEL_H, o));
  return prims;
}

// CAGR arrow across the top, first column center -> last column center,
// with a think-cell-style "+x.x% p.a." ellipse bubble on the shaft.
function cagrArrow(totals, n, slot, inner, frame, decimals) {
  const rate = cagr(totals[0], totals[n - 1], n - 1);
  if (rate == null) return [];
  const x0 = inner.x + slot * 0.5;
  const x1 = inner.x + slot * (n - 0.5);
  const y = frame.y + LABEL_H + 4;
  const sign = rate >= 0 ? "+" : "";
  const text = `${sign}${pct(rate * 100, decimals ?? 1)} p.a.`;
  return [
    {
      kind: "arrow", dir: "right", x: x0, y, w: x1 - x0, h: 9,
      fill: "#404040", meta: { role: "cagrArrow", rate }
    },
    bubble((x0 + x1) / 2, y + 4.5, text, "cagrLabel", { rate })
  ];
}

// Vertical double-headed arrow comparing first and last column totals.
function diffArrow(totals, yOf, inner) {
  const from = totals[0];
  const to = totals[totals.length - 1];
  if (from === 0) return [];
  const change = ((to - from) / Math.abs(from)) * 100;
  const yA = yOf(from);
  const yB = yOf(to);
  const top = Math.min(yA, yB);
  const h = Math.max(Math.abs(yA - yB), 8);
  const x = inner.x + inner.w + 10;
  const sign = change >= 0 ? "+" : "";
  return [
    {
      kind: "line", x1: inner.x + inner.w - 4, y1: yA, x2: x + 9, y2: yA,
      color: COLORS.connector, weight: 0.75, dash: "dash", meta: { role: "diffGuide" }
    },
    {
      kind: "arrow", dir: "upDown", x, y: top, w: 9, h,
      fill: "#404040", meta: { role: "diffArrow", change }
    },
    bubble(x + 4.5, top + h / 2, `${sign}${pct(change)}`, "diffLabel", { change })
  ];
}

// ------------------------------------------------------ clustered column

export function layoutClustered(data, frame, opts = {}) {
  const o = {
    gap: 0.3, showLabels: true, showLegend: true, axis: false,
    decimals: undefined, palette: null, ...opts
  };
  const { categories, series } = data;
  const n = categories.length;
  const k = series.length;
  if (!n || !k) throw new Error("Clustered chart needs categories and series");

  const vMax = Math.max(...series.flatMap((sr) => sr.values), 1);
  const axisW = o.axis ? AXIS_W : 0;
  const inner = { x: frame.x + axisW, w: frame.w - axisW };

  const topPad = LABEL_H + 4;
  const legendH = o.showLegend ? LABEL_H + 4 : 0;
  const botPad = LABEL_H + 4 + legendH;
  const plotH = frame.h - topPad - botPad;
  const plotBottom = frame.y + topPad + plotH;
  const yOf = (v) => plotBottom - (v * plotH) / vMax;

  const slot = inner.w / n;
  const groupW = slot * (1 - o.gap);
  const barW = groupW / k;
  const prims = [];

  if (o.axis) {
    const { ticks } = niceTicks(0, vMax);
    prims.push(...axisPrims(ticks, yOf, inner.x, inner.x + inner.w, o.decimals));
  }
  prims.push({
    kind: "line", x1: inner.x, y1: plotBottom, x2: inner.x + inner.w, y2: plotBottom,
    color: COLORS.axis, weight: 1, meta: { role: "baseline" }
  });

  categories.forEach((cat, c) => {
    const gx = inner.x + slot * c + (slot - groupW) / 2;
    series.forEach((sr, s) => {
      const raw = sr.values[c] || 0;
      const h = (raw * plotH) / vMax;
      const x = gx + s * barW;
      if (h > 0) {
        prims.push({
          kind: "rect", x, y: plotBottom - h, w: barW - 1, h,
          fill: pick(o, sr, s),
          meta: { role: "bar", series: sr.name, category: cat, value: raw }
        });
      }
      if (o.showLabels) {
        prims.push({
          kind: "text", x: x - 4, y: plotBottom - h - LABEL_H, w: barW + 7, h: LABEL_H,
          text: fmt(raw, o.decimals), fontSize: FONT - 1, align: "center",
          meta: { role: "value" }
        });
      }
    });
    prims.push({
      kind: "text", x: inner.x + slot * c, y: plotBottom + 3, w: slot, h: LABEL_H,
      text: cat, fontSize: FONT, align: "center", meta: { role: "category" }
    });
  });

  if (o.showLegend) prims.push(...legendRow(series, frame, frame.y + frame.h - LABEL_H, o));
  return prims;
}

// ------------------------------------------------------------------ mekko

export function layoutMekko(data, frame, opts = {}) {
  const o = {
    colGap: 2, showLabels: true, showTotals: true, showLegend: true,
    decimals: undefined, palette: null, ...opts
  };
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
      const fill = pick(o, sr, s);
      const seg = {
        kind: "rect", x, y: cursor - h, w: colW, h, fill,
        line: { color: "#FFFFFF", weight: 0.75 },
        meta: { role: "segment", series: sr.name, category: cat, value: raw, share }
      };
      if (o.showLabels && h >= 11 && colW >= 24) {
        seg.text = pct(share * 100, o.decimals);
        seg.fontSize = FONT;
        seg.fontColor = contrastText(fill);
      }
      prims.push(seg);
      cursor -= h;
    });

    if (o.showTotals) {
      prims.push({
        kind: "text", x, y: plotTop - LABEL_H - 2, w: colW, h: LABEL_H,
        text: fmt(totals[c], o.decimals), fontSize: FONT, bold: true, align: "center",
        meta: { role: "total" }
      });
    }
    prims.push({
      kind: "text", x, y: plotTop + plotH + 3, w: colW, h: LABEL_H,
      text: cat, fontSize: FONT, align: "center", meta: { role: "category" }
    });
    x += colW + o.colGap;
  });

  if (o.showLegend) prims.push(...legendRow(series, frame, frame.y + frame.h - LABEL_H, o));
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

function legendRow(series, frame, y, o = {}) {
  const prims = [];
  const itemW = Math.min(frame.w / series.length, 120);
  const totalW = itemW * series.length;
  let x = frame.x + (frame.w - totalW) / 2;
  series.forEach((sr, s) => {
    prims.push({
      kind: "rect", x, y: y + 3, w: 8, h: 8,
      fill: pick(o, sr, s), meta: { role: "legendSwatch" }
    });
    prims.push({
      kind: "text", x: x + 11, y, w: itemW - 14, h: LABEL_H,
      text: sr.name, fontSize: FONT - 1, align: "left", meta: { role: "legendLabel" }
    });
    x += itemW;
  });
  return prims;
}
