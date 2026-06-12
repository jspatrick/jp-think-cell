// Generates the add-in icon PNGs (a tiny bar-chart glyph on green) with no
// dependencies: raw RGBA buffers encoded as PNG via zlib + hand-built chunks.
import zlib from "node:zlib";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const OUT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "assets");

const BG = [0x15, 0xa0, 0x4a, 255];   // accent green
const BAR = [255, 255, 255, 255];     // white bars

function makeIcon(size) {
  const px = Buffer.alloc(size * size * 4);
  const put = (x, y, c) => {
    if (x < 0 || y < 0 || x >= size || y >= size) return;
    const i = (y * size + x) * 4;
    px[i] = c[0]; px[i + 1] = c[1]; px[i + 2] = c[2]; px[i + 3] = c[3];
  };
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) put(x, y, BG);

  // Three bars of increasing height, sitting on a baseline.
  const pad = Math.max(2, Math.round(size * 0.15));
  const base = size - pad;
  const barW = Math.max(2, Math.round((size - 2 * pad) / 4));
  const gap = Math.max(1, Math.round(barW / 2));
  const heights = [0.35, 0.6, 0.85].map((f) => Math.round((size - 2 * pad) * f));
  let x0 = pad;
  for (const h of heights) {
    for (let x = x0; x < x0 + barW; x++) {
      for (let y = base - h; y < base; y++) put(x, y, BAR);
    }
    x0 += barW + gap;
  }
  return encodePNG(size, size, px);
}

function encodePNG(w, h, rgba) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // color type RGBA
  // raw scanlines, each prefixed with filter byte 0
  const raw = Buffer.alloc(h * (1 + w * 4));
  for (let y = 0; y < h; y++) {
    raw[y * (1 + w * 4)] = 0;
    rgba.copy(raw, y * (1 + w * 4) + 1, y * w * 4, (y + 1) * w * 4);
  }
  const idat = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    sig,
    chunk("IHDR", ihdr),
    chunk("IDAT", idat),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body) >>> 0, 0);
  return Buffer.concat([len, body, crc]);
}

let CRC_TABLE;
function crc32(buf) {
  if (!CRC_TABLE) {
    CRC_TABLE = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c;
    }
  }
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return c ^ -1;
}

fs.mkdirSync(OUT, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  const file = path.join(OUT, `icon-${size}.png`);
  fs.writeFileSync(file, makeIcon(size));
  console.log(`wrote ${file}`);
}
