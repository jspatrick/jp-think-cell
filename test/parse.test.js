import test from "node:test";
import assert from "node:assert/strict";
import {
  parseTable,
  toWaterfall,
  toMatrix,
  toGantt,
  toNumber
} from "../src/lib/parse.js";
import { fmt, pct } from "../src/lib/format.js";

test("parseTable: detects tabs, semicolons, commas; drops blank lines", () => {
  assert.deepEqual(parseTable("a\t1\nb\t2\n\n"), [["a", "1"], ["b", "2"]]);
  assert.deepEqual(parseTable("a;1\r\nb;2"), [["a", "1"], ["b", "2"]]);
  assert.deepEqual(parseTable("a,1\nb,2"), [["a", "1"], ["b", "2"]]);
});

test("toNumber: strips currency, thousands separators, percent", () => {
  assert.equal(toNumber("1,234.5"), 1234.5);
  assert.equal(toNumber("$ 1,000"), 1000);
  assert.equal(toNumber("45%"), 45);
  assert.ok(Number.isNaN(toNumber("abc")));
});

test("toWaterfall: skips header, honors e/total markers", () => {
  const rows = parseTable("Label\tValue\nStart\t100\nCosts\t-30\nEnd\te");
  const wf = toWaterfall(rows);
  assert.deepEqual(wf, [
    { label: "Start", value: 100 },
    { label: "Costs", value: -30 },
    { label: "End", value: "e" }
  ]);
});

test("toWaterfall: no header row also works", () => {
  const wf = toWaterfall(parseTable("Start\t100\nTotal\t="));
  assert.equal(wf[1].value, "e");
});

test("toMatrix: header categories + named series", () => {
  const rows = parseTable("\tQ1\tQ2\nAlpha\t1\t2\nBeta\t3\t4");
  const m = toMatrix(rows);
  assert.deepEqual(m.categories, ["Q1", "Q2"]);
  assert.deepEqual(m.series[1], { name: "Beta", values: [3, 4] });
});

test("toGantt: header skipped, start/end normalized", () => {
  const rows = parseTable("Task\tStart\tEnd\nDesign\t3\t1\nShip\t6\t6");
  const g = toGantt(rows);
  assert.deepEqual(g[0], { label: "Design", start: 1, end: 3 });
  assert.deepEqual(g[1], { label: "Ship", start: 6, end: 6 });
});

test("fmt/pct formatting", () => {
  assert.equal(fmt(1234.4), "1,234");
  assert.equal(fmt(1.25), "1.3");
  assert.equal(fmt(-30), "-30");
  assert.equal(pct(45.2), "45%");
});
