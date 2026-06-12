// Smart elements (think-cell's element gallery, simplified):
// Harvey balls, checkboxes, process flows, traffic lights, and stamps.

/* global PowerPoint */

import { getTargetSlide } from "./render.js";

// Unicode Harvey balls — render reliably across hosts without needing
// shape-adjustment APIs (which Office.js does not expose).
const HARVEY = { 0: "○", 25: "◔", 50: "◑", 75: "◕", 100: "●" };
const CHECK = { empty: "☐", checked: "☑", crossed: "☒" };

export async function insertHarveyBall(percent) {
  const glyph = HARVEY[percent] ?? HARVEY[0];
  await insertGlyph(glyph, { size: 32, color: "#4472C4", name: `Harvey ball ${percent}%` });
}

export async function insertCheckbox(state) {
  const glyph = CHECK[state] ?? CHECK.empty;
  const color = state === "crossed" ? "#C00000" : state === "checked" ? "#70AD47" : "#000000";
  await insertGlyph(glyph, { size: 28, color, name: `Checkbox ${state}` });
}

async function insertGlyph(glyph, { size, color, name }) {
  await PowerPoint.run(async (context) => {
    const shapes = getTargetSlide(context).shapes;
    const s = shapes.addTextBox(glyph, { left: 120, top: 120, width: size + 12, height: size + 12 });
    s.name = `SlideCharts ${name}`;
    s.fill.clear();
    s.lineFormat.visible = false;
    const tf = s.textFrame;
    try {
      tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
      tf.verticalAlignment = "Middle";
    } catch (e) { /* best effort */ }
    tf.textRange.font.size = size;
    tf.textRange.font.color = color;
    try {
      tf.textRange.paragraphFormat.horizontalAlignment = "Center";
    } catch (e) { /* ignore */ }
    await context.sync();
  });
}

export async function insertProcessFlow(steps) {
  const labels = steps.length ? steps : ["Step 1", "Step 2", "Step 3"];
  await PowerPoint.run(async (context) => {
    const shapes = getTargetSlide(context).shapes;
    const left = 80, top = 200, h = 60, gap = 6;
    const totalW = 800;
    const w = (totalW - gap * (labels.length - 1)) / labels.length;
    labels.forEach((label, i) => {
      const type = i === 0
        ? PowerPoint.GeometricShapeType.pentagon
        : PowerPoint.GeometricShapeType.chevron;
      const s = shapes.addGeometricShape(type, {
        left: left + i * (w + gap), top, width: w, height: h
      });
      s.name = `SlideCharts process step ${i + 1}`;
      s.fill.setSolidColor(i === labels.length - 1 ? "#264478" : "#4472C4");
      s.lineFormat.visible = false;
      const tf = s.textFrame;
      tf.textRange.text = label;
      tf.textRange.font.size = 14;
      tf.textRange.font.color = "#FFFFFF";
      tf.textRange.font.bold = true;
      try {
        tf.verticalAlignment = "Middle";
        tf.textRange.paragraphFormat.horizontalAlignment = "Center";
      } catch (e) { /* ignore */ }
    });
    await context.sync();
  });
}

export async function insertTrafficLight(state) {
  // state: 'red' | 'amber' | 'green'
  const colors = { red: "#C00000", amber: "#FFC000", green: "#70AD47" };
  await PowerPoint.run(async (context) => {
    const shapes = getTargetSlide(context).shapes;
    const left = 120, top = 120, d = 18, gap = 4;
    const frame = shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
      left: left - 5, top: top - 5, width: d + 10, height: d * 3 + gap * 2 + 10
    });
    frame.name = "SlideCharts traffic light frame";
    frame.fill.setSolidColor("#404040");
    frame.lineFormat.visible = false;
    ["red", "amber", "green"].forEach((lamp, i) => {
      const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
        left, top: top + i * (d + gap), width: d, height: d
      });
      s.name = `SlideCharts traffic light ${lamp}`;
      s.fill.setSolidColor(lamp === state ? colors[lamp] : "#737373");
      s.lineFormat.visible = false;
    });
    await context.sync();
  });
}

export async function insertStamp(text) {
  await PowerPoint.run(async (context) => {
    const shapes = getTargetSlide(context).shapes;
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: 760, top: 30, width: 160, height: 40
    });
    s.name = `SlideCharts stamp ${text}`;
    s.fill.clear();
    s.lineFormat.color = "#C00000";
    s.lineFormat.weight = 2.5;
    const tf = s.textFrame;
    tf.textRange.text = text;
    tf.textRange.font.size = 20;
    tf.textRange.font.bold = true;
    tf.textRange.font.color = "#C00000";
    try {
      tf.verticalAlignment = "Middle";
      tf.textRange.paragraphFormat.horizontalAlignment = "Center";
    } catch (e) { /* ignore */ }
    await context.sync();
  });
}
