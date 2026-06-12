// Default series palette (Office-like) and semantic colors used by the chart engines.

export const PALETTE = [
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
  "#5B9BD5", "#70AD47", "#264478", "#9E480E",
  "#636363", "#997300", "#255E91", "#43682B"
];

export const COLORS = {
  positive: "#70AD47",
  negative: "#C00000",
  total: "#A5A5A5",
  axis: "#000000",
  connector: "#A6A6A6",
  grid: "#D9D9D9",
  label: "#000000",
  labelOnFill: "#FFFFFF",
  ganttBar: "#4472C4",
  ganttMilestone: "#C00000"
};

export function seriesColor(i) {
  return PALETTE[i % PALETTE.length];
}

// Relative luminance check so segment labels stay readable on any fill.
export function contrastText(hex) {
  const v = hex.replace("#", "");
  const r = parseInt(v.slice(0, 2), 16);
  const g = parseInt(v.slice(2, 4), 16);
  const b = parseInt(v.slice(4, 6), 16);
  const lum = 0.299 * r + 0.587 * g + 0.114 * b;
  return lum > 150 ? "#000000" : "#FFFFFF";
}
