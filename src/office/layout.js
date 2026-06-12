// Layout tools on the current shape selection: align, distribute,
// and same-size (think-cell's "smart" alignment, simplified to one click).

/* global PowerPoint, Office */

export async function applyLayout(op) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    throw new Error("Layout tools need a newer PowerPoint (PowerPointApi 1.5).");
  }
  await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id,items/left,items/top,items/width,items/height");
    await context.sync();

    const items = sel.items;
    const minNeeded = op.startsWith("dist") ? 3 : 2;
    if (items.length < minNeeded) {
      throw new Error(`Select at least ${minNeeded} shapes first.`);
    }

    const box = boundingBox(items);
    switch (op) {
      case "alignLeft":
        items.forEach((s) => { s.left = box.left; });
        break;
      case "alignCenterH":
        items.forEach((s) => { s.left = box.left + (box.width - s.width) / 2; });
        break;
      case "alignRight":
        items.forEach((s) => { s.left = box.left + box.width - s.width; });
        break;
      case "alignTop":
        items.forEach((s) => { s.top = box.top; });
        break;
      case "alignMiddleV":
        items.forEach((s) => { s.top = box.top + (box.height - s.height) / 2; });
        break;
      case "alignBottom":
        items.forEach((s) => { s.top = box.top + box.height - s.height; });
        break;
      case "distH": {
        const sorted = [...items].sort((a, b) => a.left - b.left);
        const totalW = sorted.reduce((a, s) => a + s.width, 0);
        const space = (box.width - totalW) / (sorted.length - 1);
        let x = box.left;
        sorted.forEach((s) => { s.left = x; x += s.width + space; });
        break;
      }
      case "distV": {
        const sorted = [...items].sort((a, b) => a.top - b.top);
        const totalH = sorted.reduce((a, s) => a + s.height, 0);
        const space = (box.height - totalH) / (sorted.length - 1);
        let y = box.top;
        sorted.forEach((s) => { s.top = y; y += s.height + space; });
        break;
      }
      case "sameWidth": {
        const ref = items[0].width;
        items.forEach((s) => { s.width = ref; });
        break;
      }
      case "sameHeight": {
        const ref = items[0].height;
        items.forEach((s) => { s.height = ref; });
        break;
      }
      case "sameSize": {
        const rw = items[0].width, rh = items[0].height;
        items.forEach((s) => { s.width = rw; s.height = rh; });
        break;
      }
      default:
        throw new Error(`Unknown layout operation: ${op}`);
    }
    await context.sync();
  });
}

function boundingBox(items) {
  const left = Math.min(...items.map((s) => s.left));
  const top = Math.min(...items.map((s) => s.top));
  const right = Math.max(...items.map((s) => s.left + s.width));
  const bottom = Math.max(...items.map((s) => s.top + s.height));
  return { left, top, width: right - left, height: bottom - top };
}
