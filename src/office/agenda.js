// Agenda builder (think-cell's "Agenda" feature, simplified):
// reads slide titles from the deck, lets the user edit the list in the
// taskpane, then inserts a styled agenda slide.

/* global PowerPoint */

const SLIDE_W = 960; // 16:9 deck in points
// (If the deck uses 4:3 the agenda still renders, just left-aligned.)

export async function readSlideTitles() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    for (const slide of slides.items) {
      slide.shapes.load("items/name,items/id");
    }
    await context.sync();

    // Heuristic: the title placeholder is usually named "Title ...".
    const titleShapes = slides.items.map((slide) =>
      slide.shapes.items.find((s) => /title/i.test(s.name)) || null
    );
    titleShapes.forEach((s) => {
      if (s) s.textFrame.textRange.load("text");
    });
    await context.sync();

    return titleShapes.map((s, i) => {
      const text = s ? (s.textFrame.textRange.text || "").trim() : "";
      return { index: i + 1, title: text || `(Slide ${i + 1} — no title)` };
    });
  });
}

export async function insertAgendaSlide(items, { heading = "Agenda" } = {}) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.add();
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[slides.items.length - 1];
    const shapes = slide.shapes;

    const title = shapes.addTextBox(heading, { left: 60, top: 40, width: 500, height: 44 });
    title.name = "SlideCharts agenda heading";
    styleText(title, { size: 30, bold: true, color: "#000000" });

    const left = 60;
    const width = SLIDE_W - 2 * left - 120;
    const rowH = 34;
    let y = 110;

    items.forEach((item, i) => {
      const num = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left, top: y, width: 26, height: 26
      });
      num.name = `SlideCharts agenda number ${i + 1}`;
      num.fill.setSolidColor("#4472C4");
      num.lineFormat.visible = false;
      num.textFrame.textRange.text = String(i + 1);
      styleText(num, { size: 14, bold: true, color: "#FFFFFF" });

      const label = shapes.addTextBox(item, {
        left: left + 38, top: y + 1, width, height: 26
      });
      label.name = `SlideCharts agenda item ${i + 1}`;
      styleText(label, { size: 16, bold: false, color: "#000000", align: "Left" });

      const rule = shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: left + 38, top: y + 29, width, height: 0
      });
      rule.name = `SlideCharts agenda rule ${i + 1}`;
      rule.lineFormat.color = "#D9D9D9";
      rule.lineFormat.weight = 0.75;

      y += rowH;
    });

    await context.sync();
  });
}

function styleText(shape, { size, bold, color, align = "Center" }) {
  const tf = shape.textFrame;
  try {
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
    tf.verticalAlignment = "Middle";
  } catch (e) { /* best effort on older hosts */ }
  const font = tf.textRange.font;
  font.size = size;
  font.bold = bold;
  font.color = color;
  font.name = "Calibri";
  try {
    tf.textRange.paragraphFormat.horizontalAlignment = align;
  } catch (e) { /* ignore */ }
}
