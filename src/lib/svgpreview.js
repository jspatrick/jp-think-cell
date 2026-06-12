// Renders chartmath primitives to an SVG string for the live taskpane
// preview (think-cell shows the chart updating as you type; we mirror the
// exact same primitives that will be inserted on the slide).

const DASH = { dash: "4 3", dot: "1 2" };

export function primsToSvg(prims, viewBox = { x: 0, y: 0, w: 960, h: 540 }) {
  const parts = [];
  for (const p of prims) {
    if (p.kind === "rect") parts.push(rect(p));
    else if (p.kind === "line") parts.push(line(p));
    else if (p.kind === "text") parts.push(text(p));
    else if (p.kind === "arrow") parts.push(arrow(p));
  }
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.w} ${viewBox.h}" ` +
    `font-family="Calibri, Arial, sans-serif">` +
    `<rect x="${viewBox.x}" y="${viewBox.y}" width="${viewBox.w}" height="${viewBox.h}" fill="#ffffff"/>` +
    parts.join("") +
    `</svg>`
  );
}

function rect(p) {
  const stroke = p.line ? ` stroke="${p.line.color}" stroke-width="${p.line.weight ?? 1}"` : "";
  let shape;
  if (p.shape === "ellipse") {
    shape = `<ellipse cx="${n(p.x + p.w / 2)}" cy="${n(p.y + p.h / 2)}" rx="${n(p.w / 2)}" ry="${n(p.h / 2)}" fill="${p.fill || "none"}"${stroke}/>`;
  } else if (p.shape === "diamond") {
    const cx = p.x + p.w / 2, cy = p.y + p.h / 2;
    shape = `<polygon points="${cx},${p.y} ${p.x + p.w},${cy} ${cx},${p.y + p.h} ${p.x},${cy}" fill="${p.fill}"${stroke}/>`;
  } else {
    shape = `<rect x="${n(p.x)}" y="${n(p.y)}" width="${n(p.w)}" height="${n(p.h)}" fill="${p.fill || "none"}"${stroke}/>`;
  }
  if (p.text) {
    shape += centeredText(p.x + p.w / 2, p.y + p.h / 2, p.text, p);
  }
  return shape;
}

function line(p) {
  const dash = p.dash && DASH[p.dash] ? ` stroke-dasharray="${DASH[p.dash]}"` : "";
  return `<line x1="${n(p.x1)}" y1="${n(p.y1)}" x2="${n(p.x2)}" y2="${n(p.y2)}" stroke="${p.color}" stroke-width="${p.weight ?? 1}"${dash}/>`;
}

function text(p) {
  const align = p.align || "center";
  let x, anchor;
  if (align === "left") { x = p.x; anchor = "start"; }
  else if (align === "right") { x = p.x + p.w; anchor = "end"; }
  else { x = p.x + p.w / 2; anchor = "middle"; }
  return centeredText(x, p.y + p.h / 2, p.text, p, anchor);
}

function centeredText(x, y, str, p, anchor = "middle") {
  const weight = p.bold ? ` font-weight="bold"` : "";
  return (
    `<text x="${n(x)}" y="${n(y)}" text-anchor="${anchor}" dominant-baseline="central" ` +
    `font-size="${p.fontSize ?? 10}" fill="${p.fontColor || "#000000"}"${weight}>${esc(str)}</text>`
  );
}

function arrow(p) {
  const f = p.fill || "#404040";
  if (p.dir === "upDown") {
    const cx = p.x + p.w / 2;
    const head = Math.min(p.w, p.h / 3);
    const bw = p.w * 0.34;
    return (
      `<polygon fill="${f}" points="` +
      `${n(cx)},${n(p.y)} ${n(p.x + p.w)},${n(p.y + head)} ${n(cx + bw / 2)},${n(p.y + head)} ` +
      `${n(cx + bw / 2)},${n(p.y + p.h - head)} ${n(p.x + p.w)},${n(p.y + p.h - head)} ${n(cx)},${n(p.y + p.h)} ` +
      `${n(p.x)},${n(p.y + p.h - head)} ${n(cx - bw / 2)},${n(p.y + p.h - head)} ` +
      `${n(cx - bw / 2)},${n(p.y + head)} ${n(p.x)},${n(p.y + head)}"/>`
    );
  }
  // right arrow
  const cy = p.y + p.h / 2;
  const head = Math.min(p.h * 1.2, p.w / 4, 14);
  const bh = p.h * 0.4;
  return (
    `<polygon fill="${f}" points="` +
    `${n(p.x)},${n(cy - bh / 2)} ${n(p.x + p.w - head)},${n(cy - bh / 2)} ${n(p.x + p.w - head)},${n(p.y)} ` +
    `${n(p.x + p.w)},${n(cy)} ${n(p.x + p.w - head)},${n(p.y + p.h)} ${n(p.x + p.w - head)},${n(cy + bh / 2)} ` +
    `${n(p.x)},${n(cy + bh / 2)}"/>`
  );
}

function n(v) {
  return Number(v.toFixed(2));
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
