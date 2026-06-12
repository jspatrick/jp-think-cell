// Dev utility: rasterize chart layouts to PNGs (rects and lines only — text
// is approximated as a light gray strip) for a quick geometry sanity check
// without opening PowerPoint. Usage: node scripts/preview.js
import zlib from "node:zlib";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { layoutWaterfall, layoutStacked, layoutMekko, layoutGantt } from "../src/lib/chartmath.js";
import { parseTable, toWaterfall, toMatrix, toGantt } from "../src/lib/parse.js";

const OUT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "preview");
const W = 960, H = 540;
const FRAME = { x: 120, y: 110, w: 720, h: 360 };

const EXAMPLES = {
  waterfall: () => layoutWaterfall(toWaterfall(parseTable(
    "Label\tValue\n2024\t820\nVolume\t95\nPrice\t40\nChurn\t-60\nFX\t-25\n2025\te")), FRAME),
  stacked: () => layoutStacked(toMatrix(parseTable(
    "\tQ1\tQ2\tQ3\tQ4\nAmericas\t120\t135\t150\t170\nEMEA\t90\t95\t100\t110\nAPAC\t45\t55\t70\t85")), FRAME),
  stacked100: () => layoutStacked(toMatrix(parseTable(
    "\tQ1\tQ2\tQ3\tQ4\nAmericas\t120\t135\t150\t170\nEMEA\t90\t95\t100\t110\nAPAC\t45\t55\t70\t85")), FRAME, { percent: true }),
  mekko: () => layoutMekko(toMatrix(parseTable(
    "\tSeg A\tSeg B\tSeg C\nUs\t40\t25\t10\nComp 1\t30\t45\t15\nComp 2\t30\t30\t75")), FRAME),
  gantt: () => layoutGantt(toGantt(parseTable(
    "Task\tStart\tEnd\nDiscovery\t1\t3\nDesign\t2\t5\nBuild\t4\t10\nTest\t9\t12\nLaunch\t12\t12")), FRAME)
};

function render(prims) {
  const px = Buffer.alloc(W * H * 4, 255);
  const put = (x, y, c) => {
    x = Math.round(x); y = Math.round(y);
    if (x < 0 || y < 0 || x >= W || y >= H) return;
    const i = (y * W + x) * 4;
    px[i] = c[0]; px[i + 1] = c[1]; px[i + 2] = c[2]; px[i + 3] = 255;
  };
  const fillRect = (x, y, w, h, c) => {
    for (let yy = y; yy < y + h; yy++) for (let xx = x; xx < x + w; xx++) put(xx, yy, c);
  };
  const hex = (s) => {
    const v = s.replace("#", "");
    return [parseInt(v.slice(0, 2), 16), parseInt(v.slice(2, 4), 16), parseInt(v.slice(4, 6), 16)];
  };
  for (const p of prims) {
    if (p.kind === "rect") {
      fillRect(p.x, p.y, Math.max(p.w, 1), Math.max(p.h, 1), hex(p.fill || "#cccccc"));
    } else if (p.kind === "line") {
      const c = hex(p.color || "#000000");
      const steps = Math.max(Math.abs(p.x2 - p.x1), Math.abs(p.y2 - p.y1), 1);
      for (let i = 0; i <= steps; i++) {
        if (p.dash === "dash" && Math.floor(i / 4) % 2) continue;
        put(p.x1 + ((p.x2 - p.x1) * i) / steps, p.y1 + ((p.y2 - p.y1) * i) / steps, c);
      }
    } else if (p.kind === "text" && p.text) {
      fillRect(p.x + p.w * 0.25, p.y + p.h * 0.3, p.w * 0.5, p.h * 0.4, [225, 225, 225]);
    }
  }
  return encodePNG(W, H, px);
}

function encodePNG(w, h, rgba) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.alloc(h * (1 + w * 4));
  for (let y = 0; y < h; y++) {
    raw[y * (1 + w * 4)] = 0;
    rgba.copy(raw, y * (1 + w * 4) + 1, y * w * 4, (y + 1) * w * 4);
  }
  return Buffer.concat([sig, chunk("IHDR", ihdr), chunk("IDAT", zlib.deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(body) >>> 0, 0);
  return Buffer.concat([len, body, crc]);
}
let T;
function crc32(buf) {
  if (!T) {
    T = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      T[n] = c;
    }
  }
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = T[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return c ^ -1;
}

fs.mkdirSync(OUT, { recursive: true });
for (const [name, fn] of Object.entries(EXAMPLES)) {
  const file = path.join(OUT, `${name}.png`);
  fs.writeFileSync(file, render(fn()));
  console.log(`wrote ${file}`);
}
