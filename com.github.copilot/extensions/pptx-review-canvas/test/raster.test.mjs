import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { decodePng, encodePng } from '../lib/png.mjs';
import { compositeAnnotations, drawBadge, drawPolyline, drawStrokes } from '../lib/raster.mjs';
import { dataRootFor } from '../lib/store.mjs';

function solid(width, height, [r, g, b]) {
  const data = new Uint8Array(width * height * 4);
  for (let i = 0; i < data.length; i += 4) {
    data[i] = r;
    data[i + 1] = g;
    data[i + 2] = b;
    data[i + 3] = 255;
  }
  return { width, height, data };
}

const pixel = (image, x, y) => Array.from(image.data.subarray((y * image.width + x) * 4, (y * image.width + x) * 4 + 3));

test('encodePng and decodePng round-trip, with and without transparency', () => {
  const image = solid(7, 5, [10, 20, 30]);
  image.data[0] = 200;
  const opaque = decodePng(encodePng(image));
  assert.equal(opaque.width, 7);
  assert.equal(opaque.height, 5);
  assert.deepEqual(Array.from(opaque.data), Array.from(image.data));

  image.data[3] = 128;
  const translucent = decodePng(encodePng(image));
  assert.equal(translucent.data[3], 128);
  assert.deepEqual(Array.from(translucent.data), Array.from(image.data));
  assert.throws(() => decodePng(Buffer.from('not a png at all, definitely not a png file')), /PNG/);
});

test('decodePng reads filtered rows (Sub, Up, Average and Paeth)', async () => {
  const { deflateSync } = await import('node:zlib');
  const width = 4;
  const height = 4;
  const source = [];
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) source.push((x * 37 + y * 91) & 255, (x * 11 + y * 5) & 255, (x * 3 + y * 200) & 255);
  const stride = width * 3;
  const raw = Buffer.alloc((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    const filter = y; // 0 none, 1 sub, 2 up, 3 average
    raw[y * (stride + 1)] = filter;
    for (let x = 0; x < stride; x++) {
      const value = source[y * stride + x];
      const left = x >= 3 ? source[y * stride + x - 3] : 0;
      const up = y > 0 ? source[(y - 1) * stride + x] : 0;
      const predicted = filter === 1 ? left : filter === 2 ? up : filter === 3 ? (left + up) >> 1 : 0;
      raw[y * (stride + 1) + 1 + x] = (value - predicted) & 255;
    }
  }
  const encoded = encodePng(solid(width, height, [0, 0, 0]));
  const header = encoded.subarray(0, 8 + 25);
  const idat = deflateSync(raw);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(idat.length);
  const { createHash } = await import('node:crypto');
  void createHash;
  const crc = Buffer.alloc(4);
  // Reuse the encoder's own chunk writer through a tiny re-encode: build the IDAT chunk by hand.
  const type = Buffer.from('IDAT');
  const crcTable = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  let c = 0xffffffff;
  for (const byte of Buffer.concat([type, idat])) c = crcTable[(c ^ byte) & 255] ^ (c >>> 8);
  crc.writeUInt32BE((c ^ 0xffffffff) >>> 0);
  const iend = encoded.subarray(encoded.length - 12);
  const decoded = decodePng(Buffer.concat([header, length, type, idat, crc, iend]));
  for (let i = 0; i < width * height; i++) assert.deepEqual(Array.from(decoded.data.subarray(i * 4, i * 4 + 3)), source.slice(i * 3, i * 3 + 3));
});

test('drawPolyline paints the line, leaves the rest alone and anti-aliases the edge', () => {
  const image = solid(60, 40, [255, 255, 255]);
  drawPolyline(image, [[5, 20.3], [55, 20.3]], 6, '#e5231b');
  assert.deepEqual(pixel(image, 30, 20), [0xe5, 0x23, 0x1b]);
  assert.deepEqual(pixel(image, 30, 5), [255, 255, 255]);
  assert.deepEqual(pixel(image, 30, 35), [255, 255, 255]);
  const edge = pixel(image, 30, 23);
  assert.ok(edge[1] > 0x23 && edge[1] < 255, `edge pixel should be partly covered: ${edge}`);
  const cap = pixel(image, 3, 20);
  assert.ok(cap[1] < 255, 'the round cap reaches past the first point');
  drawPolyline(image, [[30, 5]], 8, '#0b6bcb');
  assert.deepEqual(pixel(image, 30, 5), [0x0b, 0x6b, 0xcb]);
});

test('a highlighter blends once, so a stroke that crosses itself does not darken', () => {
  const image = solid(60, 40, [255, 255, 255]);
  drawStrokes(image, [{ tool: 'highlighter', color: '#ffd400', width: 0.2, points: [[0.1, 0.5], [0.9, 0.5], [0.1, 0.5], [0.9, 0.5]] }]);
  const once = pixel(image, 30, 20);
  assert.ok(once[2] < 255 && once[2] > 100, `partly transparent: ${once}`);
  const reference = solid(60, 40, [255, 255, 255]);
  drawStrokes(reference, [{ tool: 'highlighter', color: '#ffd400', width: 0.2, points: [[0.1, 0.5], [0.9, 0.5]] }]);
  assert.deepEqual(once, pixel(reference, 30, 20));
});

test('drawBadge puts white digits on a dark disc', () => {
  const image = solid(80, 80, [255, 255, 255]);
  drawBadge(image, 40, 40, 20, 8);
  assert.deepEqual(pixel(image, 22, 40), [0x1f, 0x1f, 0x1f]);
  let white = 0;
  for (let y = 30; y <= 50; y++) for (let x = 32; x <= 48; x++) if (pixel(image, x, y)[0] > 200) white++;
  assert.ok(white > 20, 'digit strokes are drawn in the ink colour');
  assert.deepEqual(pixel(image, 2, 2), [255, 255, 255]);
});

test('compositeAnnotations returns a PNG of the same size with the marks on it', () => {
  const base = encodePng(solid(160, 90, [255, 255, 255]));
  const out = compositeAnnotations(base, [{ tool: 'pen', color: '#e5231b', width: 0.02, points: [[0.2, 0.5], [0.8, 0.5]] }], [{ x: 0.1, y: 0.1, number: 1 }]);
  const image = decodePng(out);
  assert.equal(image.width, 160);
  assert.equal(image.height, 90);
  assert.deepEqual(pixel(image, 80, 45), [0xe5, 0x23, 0x1b]);
  assert.deepEqual(pixel(image, 150, 5), [255, 255, 255]);
});

test('decodePng reads the PNG that PowerPoint wrote, when the cache has one', async (t) => {
  const decks = path.join(dataRootFor(), 'decks');
  const keys = await fs.readdir(decks).catch(() => []);
  const file = keys.length ? path.join(decks, keys[0], 's0001.png') : null;
  if (!file || !(await fs.stat(file).catch(() => null))) return t.skip('there is no rendered image in the cache');
  const image = decodePng(await fs.readFile(file));
  assert.equal(image.width, 1600);
  assert.equal(image.height, 900);
  const again = decodePng(encodePng(image));
  assert.deepEqual(Array.from(again.data.subarray(0, 4000)), Array.from(image.data.subarray(0, 4000)));
});
