import test from "node:test";
import assert from "node:assert/strict";
import {
  layoutWaterfall,
  layoutStacked,
  layoutClustered,
  layoutMekko,
  layoutGantt,
  niceTicks,
  cagr
} from "../src/lib/chartmath.js";
import { primsToSvg } from "../src/lib/svgpreview.js";

const FRAME = { x: 100, y: 100, w: 600, h: 300 };

const rects = (prims, role) =>
  prims.filter((p) => p.kind === "rect" && p.meta?.role === role);
const close = (a, b, eps = 0.01) => assert.ok(Math.abs(a - b) < eps, `${a} !~ ${b}`);

test("waterfall: bars, totals, and connectors follow cumulative math", () => {
  const rows = [
    { label: "Start", value: 100 },
    { label: "Growth", value: 50 },
    { label: "Costs", value: -30 },
    { label: "End", value: "e" }
  ];
  const prims = layoutWaterfall(rows, FRAME);
  const bars = rects(prims, "bar");
  assert.equal(bars.length, 4);

  // Total bar spans 0..120
  assert.deepEqual(
    { from: bars[3].meta.from, to: bars[3].meta.to },
    { from: 0, to: 120 }
  );

  // Connectors sit at cumulative levels 100, 150, 120
  const connectors = prims.filter((p) => p.meta?.role === "connector");
  assert.deepEqual(connectors.map((c) => c.meta.level), [100, 150, 120]);
  connectors.forEach((c) => assert.equal(c.y1, c.y2, "connector is horizontal"));

  // Vertical scale: cost bar height / start bar height === 30 / 100
  close(bars[2].h / bars[0].h, 30 / 100);

  // The "Costs" bar floats: its bottom is the top of the cumulative-120 level,
  // i.e. bottom of bar2 aligns with top of total bar.
  close(bars[2].y + bars[2].h, bars[3].y);

  // Everything stays inside the frame
  for (const b of bars) {
    assert.ok(b.x >= FRAME.x && b.x + b.w <= FRAME.x + FRAME.w + 0.01);
    assert.ok(b.y >= FRAME.y && b.y + b.h <= FRAME.y + FRAME.h + 0.01);
  }
});

test("waterfall: negative cumulative keeps baseline at zero level", () => {
  const rows = [
    { label: "A", value: -50 },
    { label: "B", value: 80 },
    { label: "T", value: "e" }
  ];
  const prims = layoutWaterfall(rows, FRAME);
  const baseline = prims.find((p) => p.meta?.role === "baseline");
  const bars = rects(prims, "bar");
  // Negative bar hangs below the baseline
  close(bars[0].y, baseline.y1);
  // Total (30) sits above the baseline
  close(bars[2].y + bars[2].h, baseline.y1);
});

test("stacked: segment heights are proportional and stack bottom-up", () => {
  const data = {
    categories: ["Q1", "Q2"],
    series: [
      { name: "A", values: [10, 20] },
      { name: "B", values: [30, 20] }
    ]
  };
  const prims = layoutStacked(data, FRAME, { showLegend: false });
  const segs = rects(prims, "segment");
  assert.equal(segs.length, 4);

  const q1 = segs.filter((s) => s.meta.category === "Q1");
  // B (30) is 3x the height of A (10)
  close(q1[1].h / q1[0].h, 3);
  // A is drawn first from the baseline, B stacks directly on top of it
  close(q1[0].y, q1[1].y + q1[1].h);
});

test("stacked: 100% mode gives equal column heights", () => {
  const data = {
    categories: ["Q1", "Q2"],
    series: [
      { name: "A", values: [10, 200] },
      { name: "B", values: [30, 200] }
    ]
  };
  const prims = layoutStacked(data, FRAME, { percent: true, showLegend: false });
  const segs = rects(prims, "segment");
  const colH = (cat) =>
    segs.filter((s) => s.meta.category === cat).reduce((a, s) => a + s.h, 0);
  close(colH("Q1"), colH("Q2"));
  // Q1: A is 25%, B is 75%
  const q1 = segs.filter((s) => s.meta.category === "Q1");
  close(q1[0].h / colH("Q1"), 0.25);
});

test("mekko: column widths proportional to totals, all columns full height", () => {
  const data = {
    categories: ["X", "Y", "Z"],
    series: [
      { name: "A", values: [10, 20, 10] },
      { name: "B", values: [10, 40, 10] }
    ]
  };
  const prims = layoutMekko(data, FRAME, { showLegend: false, colGap: 2 });
  const segs = rects(prims, "segment");
  const colW = (cat) => segs.find((s) => s.meta.category === cat).w;
  // Totals: X=20, Y=60, Z=20 -> Y is 3x wider than X
  close(colW("Y") / colW("X"), 3);
  // Widths + gaps fill the frame exactly
  close(colW("X") + colW("Y") + colW("Z") + 2 * 2, FRAME.w);
  // Every column's segments sum to the same plot height
  const colH = (cat) =>
    segs.filter((s) => s.meta.category === cat).reduce((a, s) => a + s.h, 0);
  close(colH("X"), colH("Y"));
  close(colH("X"), colH("Z"));
});

test("gantt: bar positions follow the time scale; milestones for zero duration", () => {
  const tasks = [
    { label: "Design", start: 1, end: 3 },
    { label: "Build", start: 2, end: 6 },
    { label: "Launch", start: 6, end: 6 }
  ];
  const prims = layoutGantt(tasks, FRAME);
  const bars = rects(prims, "bar");
  assert.equal(bars.length, 2);
  // Build (4 units) is twice as wide as Design (2 units)
  close(bars[1].w / bars[0].w, 2);
  // Design starts at t=1 == grid origin (tMin=1)
  const grid = prims.filter((p) => p.meta?.role === "grid");
  close(bars[0].x, grid[0].x1);
  // Milestone exists for Launch
  assert.equal(rects(prims, "milestone").length, 1);
  // Row shading is drawn before bars (z-order)
  const shadeIdx = prims.findIndex((p) => p.meta?.role === "rowShade");
  const barIdx = prims.findIndex((p) => p.meta?.role === "bar");
  assert.ok(shadeIdx < barIdx);
});

test("niceTicks: round steps covering the range, always including 0", () => {
  const t1 = niceTicks(0, 437);
  assert.equal(t1.step, 100);
  assert.deepEqual(t1.ticks, [0, 100, 200, 300, 400]);
  const t2 = niceTicks(-80, 120);
  assert.ok(t2.ticks.includes(0));
  assert.ok(t2.ticks[0] >= -80 && t2.ticks.at(-1) <= 120);
});

test("cagr: 100 -> 200 over 3 periods is ~26%", () => {
  const r = cagr(100, 200, 3);
  assert.ok(Math.abs(r - (Math.pow(2, 1 / 3) - 1)) < 1e-12);
  assert.equal(cagr(0, 100, 3), null);
  assert.equal(cagr(100, -5, 3), null);
});

test("stacked: CAGR arrow spans first to last column centers with rate label", () => {
  const data = {
    categories: ["2022", "2023", "2024", "2025"],
    series: [{ name: "A", values: [100, 120, 150, 200] }]
  };
  const prims = layoutStacked(data, FRAME, { cagr: true, showLegend: false });
  const arrow = prims.find((p) => p.kind === "arrow" && p.meta.role === "cagrArrow");
  const label = prims.find((p) => p.meta?.role === "cagrLabel");
  assert.ok(arrow && label);
  close(arrow.meta.rate, Math.pow(2, 1 / 3) - 1, 1e-9);
  assert.match(label.text, /CAGR \+26/);
  // arrow spans from center of col 0 to center of col 3
  const segs = rects(prims, "segment");
  const c0 = segs[0].x + segs[0].w / 2;
  const c3 = segs.at(-1).x + segs.at(-1).w / 2;
  close(arrow.x, c0, 1);
  close(arrow.x + arrow.w, c3, 1);
});

test("stacked: difference arrow compares first and last totals", () => {
  const data = {
    categories: ["A", "B"],
    series: [{ name: "S", values: [100, 150] }]
  };
  const prims = layoutStacked(data, FRAME, { diff: true, showLegend: false });
  const arrow = prims.find((p) => p.kind === "arrow" && p.meta.role === "diffArrow");
  const label = prims.find((p) => p.meta?.role === "diffLabel");
  assert.ok(arrow && label);
  close(arrow.meta.change, 50);
  assert.equal(label.text, "+50%");
  // arrow sits to the right of the plot, inside the frame
  assert.ok(arrow.x > FRAME.x + FRAME.w - 60 && arrow.x + arrow.w < FRAME.x + FRAME.w);
});

test("stacked: value axis adds gridlines at tick levels and shifts bars right", () => {
  const data = {
    categories: ["A", "B"],
    series: [{ name: "S", values: [300, 437] }]
  };
  const plain = layoutStacked(data, FRAME, { showLegend: false });
  const withAxis = layoutStacked(data, FRAME, { axis: true, showLegend: false });
  const grid = withAxis.filter((p) => p.meta?.role === "gridline");
  const tickLabels = withAxis.filter((p) => p.meta?.role === "tickLabel");
  assert.equal(grid.length, tickLabels.length);
  assert.deepEqual(grid.map((g) => g.meta.value), [0, 100, 200, 300, 400]);
  // bars shift right to make room for the axis gutter
  assert.ok(rects(withAxis, "segment")[0].x > rects(plain, "segment")[0].x);
  // gridlines are horizontal and span the plot width
  grid.forEach((g) => assert.equal(g.y1, g.y2));
});

test("clustered: k bars per category, heights proportional to values", () => {
  const data = {
    categories: ["Q1", "Q2"],
    series: [
      { name: "A", values: [10, 20] },
      { name: "B", values: [30, 40] }
    ]
  };
  const prims = layoutClustered(data, FRAME, { showLegend: false });
  const bars = rects(prims, "bar");
  assert.equal(bars.length, 4);
  const q1 = bars.filter((b) => b.meta.category === "Q1");
  close(q1[1].h / q1[0].h, 3);
  // bars within a group don't overlap
  assert.ok(q1[0].x + q1[0].w <= q1[1].x + 0.01);
});

test("svg preview renders all primitive kinds", () => {
  const data = {
    categories: ["2022", "2023", "2024"],
    series: [{ name: "A", values: [100, 130, 160] }]
  };
  const prims = layoutStacked(data, FRAME, { cagr: true, diff: true, axis: true });
  const svg = primsToSvg(prims);
  assert.ok(svg.startsWith("<svg"));
  assert.ok(svg.includes("<rect"));
  assert.ok(svg.includes("<line"));
  assert.ok(svg.includes("<text"));
  assert.ok(svg.includes("<polygon")); // arrows
  assert.ok(svg.includes("CAGR"));
  assert.ok(!svg.includes("undefined"));
});
