// Menggambar ikon Master (PNG) tanpa dependensi: kotak biru tua, huruf "M" putih, titik emas.
// Pakai: node scripts/make-icons.js

import { writeFileSync, mkdirSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'src', 'ui', 'assets');
mkdirSync(out, { recursive: true });

const CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = t[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
})();

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(CRC(td));
  return Buffer.concat([len, td, crc]);
}

function png(size, pixel) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x, y);
      const o = y * (size * 4 + 1) + 1 + x * 4;
      raw[o] = r;
      raw[o + 1] = g;
      raw[o + 2] = b;
      raw[o + 3] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

function distSeg(px, py, ax, ay, bx, by) {
  const dx = bx - ax;
  const dy = by - ay;
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

// Koordinat dalam satuan 0..1
const SEGMENTS = [
  [0.27, 0.74, 0.27, 0.3],
  [0.27, 0.3, 0.5, 0.58],
  [0.5, 0.58, 0.73, 0.3],
  [0.73, 0.3, 0.73, 0.74],
];

function shade(u, v) {
  // Kotak bersudut membulat
  const r = 0.2;
  const cx = Math.min(Math.max(u, r), 1 - r);
  const cy = Math.min(Math.max(v, r), 1 - r);
  if (Math.hypot(u - cx, v - cy) > r) return null;
  // Titik emas (aksen AI)
  if (Math.hypot(u - 0.8, v - 0.2) < 0.075) return [245, 183, 0];
  for (const s of SEGMENTS) if (distSeg(u, v, ...s) < 0.062) return [255, 255, 255];
  // Gradasi biru
  const k = v * 0.35;
  return [Math.round(31 + 20 * k), Math.round(56 + 40 * k), Math.round(100 + 90 * k)];
}

function render(size) {
  const ss = 4;
  return png(size, (x, y) => {
    let r = 0;
    let g = 0;
    let b = 0;
    let a = 0;
    for (let i = 0; i < ss; i++)
      for (let j = 0; j < ss; j++) {
        const c = shade((x + (i + 0.5) / ss) / size, (y + (j + 0.5) / ss) / size);
        if (c) {
          r += c[0];
          g += c[1];
          b += c[2];
          a++;
        }
      }
    if (!a) return [0, 0, 0, 0];
    return [Math.round(r / a), Math.round(g / a), Math.round(b / a), Math.round((a / (ss * ss)) * 255)];
  });
}

for (const size of [16, 32, 64, 80, 128]) {
  writeFileSync(path.join(out, `icon-${size}.png`), render(size));
}
console.log('ikon dibuat di', path.relative(root, out));
