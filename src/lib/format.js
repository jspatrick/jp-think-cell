// Number formatting for chart labels.

export function fmt(v, decimals) {
  if (v == null || Number.isNaN(v)) return "";
  let d = decimals;
  if (d == null) {
    d = Math.abs(v) < 10 && !Number.isInteger(v) ? 1 : 0;
  }
  const rounded = Number(v.toFixed(d));
  return rounded.toLocaleString("en-US", {
    minimumFractionDigits: 0,
    maximumFractionDigits: d
  });
}

export function pct(v, decimals = 0) {
  if (v == null || Number.isNaN(v)) return "";
  return `${fmt(v, decimals)}%`;
}
