import test from "node:test";
import assert from "node:assert/strict";
import {
  layoutWaterfall,
  layoutStacked,
  layoutMekko,
  layoutGantt
} from "../src/lib/chartmath.js";

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
