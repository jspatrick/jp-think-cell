// Parsing of pasted datasheet text (think-cell style: paste straight from Excel).
// Supports tab, semicolon, and comma separated input.

export function detectDelimiter(text) {
  if (text.includes("\t")) return "\t";
  if (text.includes(";")) return ";";
  return ",";
}

export function parseTable(text) {
  const delim = detectDelimiter(text);
  return text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => line.split(delim).map((c) => c.trim()))
    .filter((row) => row.some((c) => c !== ""));
}

export function toNumber(cell) {
  if (cell == null) return NaN;
  const cleaned = String(cell).replace(/[,\s%$€£]/g, "");
  if (cleaned === "" || cleaned === "-") return NaN;
  return Number(cleaned);
}

function isNumeric(cell) {
  return !Number.isNaN(toNumber(cell));
}

// Waterfall input: rows of [label, value]. Value "e" (or "total"/"=") marks a
// computed subtotal/total bar, as in think-cell's datasheet. A header row is
// skipped automatically when its second cell is neither numeric nor a marker.
export function toWaterfall(rows) {
  if (rows.length === 0) throw new Error("No data");
  let start = 0;
  const first = rows[0];
  if (first.length >= 2 && !isNumeric(first[1]) && !isTotalMarker(first[1])) {
    start = 1;
  }
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const [label, raw] = rows[i];
    if (raw == null) continue;
    if (isTotalMarker(raw)) {
      out.push({ label: label || "Total", value: "e" });
    } else {
      const v = toNumber(raw);
      if (Number.isNaN(v)) throw new Error(`Row ${i + 1}: "${raw}" is not a number`);
      out.push({ label: label || "", value: v });
    }
  }
  if (out.length === 0) throw new Error("No data rows found");
  return out;
}

function isTotalMarker(cell) {
  const c = String(cell).trim().toLowerCase();
  return c === "e" || c === "total" || c === "=";
}

// Matrix input for stacked / mekko charts:
//   (corner)  Cat1  Cat2 ...
//   SeriesA   10    12
//   SeriesB    5     6
export function toMatrix(rows) {
  if (rows.length < 2) throw new Error("Need a header row plus at least one series row");
  const categories = rows[0].slice(1).filter((c) => c !== "");
  const series = [];
  for (let i = 1; i < rows.length; i++) {
    const name = rows[i][0] || `Series ${i}`;
    const values = categories.map((_, c) => {
      const v = toNumber(rows[i][c + 1]);
      return Number.isNaN(v) ? 0 : v;
    });
    series.push({ name, values });
  }
  if (categories.length === 0) throw new Error("No categories found in header row");
  return { categories, series };
}

// Gantt input: rows of [task, start, end] in arbitrary integer time units
// (weeks, months...). Header row auto-skipped. end === start renders a milestone.
export function toGantt(rows) {
  if (rows.length === 0) throw new Error("No data");
  let start = 0;
  if (rows[0].length >= 3 && !isNumeric(rows[0][1])) start = 1;
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const [label, s, e] = rows[i];
    const sv = toNumber(s);
    const ev = toNumber(e ?? s);
    if (Number.isNaN(sv) || Number.isNaN(ev)) {
      throw new Error(`Row ${i + 1}: start/end must be numbers`);
    }
    out.push({ label: label || `Task ${i}`, start: Math.min(sv, ev), end: Math.max(sv, ev) });
  }
  if (out.length === 0) throw new Error("No task rows found");
  return out;
}
