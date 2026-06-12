// Office.js rendering layer: turns chartmath primitives into native
// PowerPoint shapes on the active slide. Charts are inserted as plain
// shapes (like think-cell's output), so users can recolor or tweak them
// with normal PowerPoint tools afterwards.
//
// Every shape of a chart is tagged via its name with the chart id, which
// lets updateChart() replace a previously inserted chart in place —
// mirroring think-cell's datasheet -> slide sync.

/* global PowerPoint, Office */

const DASH_MAP = { dash: "Dash", dot: "Dot", solid: "Solid" };

export function getTargetSlide(context) {
  if (Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    return context.presentation.getSelectedSlides().getItemAt(0);
  }
  return context.presentation.slides.getItemAt(0);
}

export function newChartId() {
  return `SC[${Math.random().toString(36).slice(2, 7)}]`;
}

export async function insertPrimitives(prims, { chartId = newChartId() } = {}) {
  await PowerPoint.run(async (context) => {
    const slide = getTargetSlide(context);
    drawAll(slide.shapes, prims, chartId);
    await context.sync();
  });
  return chartId;
}

// Replace a previously inserted chart: delete all shapes tagged with the
// chart id on the current slide, then draw the new primitives.
export async function updateChart(chartId, prims) {
  let replaced = false;
  await PowerPoint.run(async (context) => {
    const slide = getTargetSlide(context);
    slide.shapes.load("items/name,items/id");
    await context.sync();

    const old = slide.shapes.items.filter((s) => s.name.startsWith(chartId));
    replaced = old.length > 0;
    old.forEach((s) => s.delete());
    drawAll(slide.shapes, prims, chartId);
    await context.sync();
  });
  return replaced;
}

function drawAll(shapes, prims, chartId) {
  let i = 0;
  for (const p of prims) {
    i += 1;
    const name = `${chartId} ${p.meta?.role || p.kind} ${i}`;
    if (p.kind === "rect") drawRect(shapes, p, name);
    else if (p.kind === "line") drawLine(shapes, p, name);
    else if (p.kind === "text") drawText(shapes, p, name);
    else if (p.kind === "arrow") drawArrow(shapes, p, name);
  }
}

function geometricType(p) {
  if (p.shape === "diamond") return PowerPoint.GeometricShapeType.diamond;
  if (p.shape === "ellipse") return PowerPoint.GeometricShapeType.ellipse;
  if (p.shape === "chevron") return PowerPoint.GeometricShapeType.chevron;
  if (p.shape === "pentagon") return PowerPoint.GeometricShapeType.pentagon;
  if (p.shape === "roundRect") return PowerPoint.GeometricShapeType.roundRectangle;
  return PowerPoint.GeometricShapeType.rectangle;
}

function drawRect(shapes, p, name) {
  const s = shapes.addGeometricShape(geometricType(p), {
    left: p.x, top: p.y, width: Math.max(p.w, 0.5), height: Math.max(p.h, 0.5)
  });
  s.name = name;
  if (p.fill) s.fill.setSolidColor(normColor(p.fill));
  else s.fill.clear();
  if (p.line) {
    s.lineFormat.color = normColor(p.line.color);
    s.lineFormat.weight = p.line.weight ?? 1;
  } else {
    s.lineFormat.visible = false;
  }
  if (p.text != null && p.text !== "") applyText(s, p, true);
}

function drawArrow(shapes, p, name) {
  const type = p.dir === "upDown"
    ? PowerPoint.GeometricShapeType.upDownArrow
    : PowerPoint.GeometricShapeType.rightArrow;
  const s = shapes.addGeometricShape(type, {
    left: p.x, top: p.y, width: Math.max(p.w, 1), height: Math.max(p.h, 1)
  });
  s.name = name;
  s.fill.setSolidColor(normColor(p.fill || "#404040"));
  s.lineFormat.visible = false;
}

function drawLine(shapes, p, name) {
  const s = shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: Math.min(p.x1, p.x2),
    top: Math.min(p.y1, p.y2),
    width: Math.abs(p.x2 - p.x1),
    height: Math.abs(p.y2 - p.y1)
  });
  s.name = name;
  s.lineFormat.color = normColor(p.color || "#000000");
  s.lineFormat.weight = p.weight ?? 1;
  if (p.dash && DASH_MAP[p.dash]) s.lineFormat.dashStyle = DASH_MAP[p.dash];
}

function drawText(shapes, p, name) {
  const s = shapes.addTextBox(p.text ?? "", {
    left: p.x, top: p.y, width: Math.max(p.w, 1), height: Math.max(p.h, 1)
  });
  s.name = name;
  s.fill.clear();
  s.lineFormat.visible = false;
  applyText(s, p, false);
}

function applyText(shape, p, isShapeText) {
  const tf = shape.textFrame;
  try {
    if (isShapeText) tf.textRange.text = p.text;
    tf.leftMargin = 0;
    tf.rightMargin = 0;
    tf.topMargin = 0;
    tf.bottomMargin = 0;
    tf.wordWrap = true;
    tf.verticalAlignment = "Middle";
    if (!isShapeText) tf.autoSizeSetting = "AutoSizeNone";
  } catch (e) {
    // Some text-frame properties are unavailable on older hosts; best effort.
  }
  const font = tf.textRange.font;
  font.size = p.fontSize ?? 10;
  font.color = normColor(p.fontColor ?? "#000000");
  font.bold = !!p.bold;
  font.name = p.fontName ?? "Calibri";
  try {
    tf.textRange.paragraphFormat.horizontalAlignment = alignOf(p.align);
  } catch (e) {
    // paragraphFormat needs PowerPointApi 1.4; ignore if missing.
  }
}

function alignOf(a) {
  if (a === "left") return "Left";
  if (a === "right") return "Right";
  return "Center";
}

function normColor(c) {
  return c.startsWith("#") ? c : `#${c}`;
}
