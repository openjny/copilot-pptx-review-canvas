import { deflateSync, inflateSync } from 'node:zlib';

const SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const CHANNELS = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let c = 0xffffffff;
  for (let i = 0; i < buffer.length; i++) c = CRC_TABLE[(c ^ buffer[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'ascii');
  data.copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
}

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

/** Decodes an 8-bit, non-interlaced PNG (what PowerPoint writes) into RGBA bytes. */
export function decodePng(buffer) {
  if (buffer.length < 33 || !buffer.subarray(0, 8).equals(SIGNATURE)) throw new Error('Not a PNG');
  let width = 0;
  let height = 0;
  let colorType = -1;
  let palette = null;
  let transparency = null;
  const parts = [];
  for (let offset = 8; offset + 12 <= buffer.length;) {
    const length = buffer.readUInt32BE(offset);
    const type = buffer.toString('ascii', offset + 4, offset + 8);
    const data = buffer.subarray(offset + 8, offset + 8 + length);
    if (type === 'IHDR') {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      const depth = data[8];
      colorType = data[9];
      if (depth !== 8 || !(colorType in CHANNELS) || data[12] !== 0) throw new Error(`Unsupported PNG format (depth ${depth}, color type ${colorType}, interlace ${data[12]})`);
    } else if (type === 'PLTE') palette = data;
    else if (type === 'tRNS') transparency = data;
    else if (type === 'IDAT') parts.push(data);
    else if (type === 'IEND') break;
    offset += 12 + length;
  }
  if (!width || !height || colorType < 0 || parts.length === 0) throw new Error('Malformed PNG structure');
  if (colorType === 3 && !palette) throw new Error('The PNG has no palette');
  if (width * height > 50_000_000) throw new Error('The PNG is too large');

  const channels = CHANNELS[colorType];
  const stride = width * channels;
  const raw = inflateSync(Buffer.concat(parts));
  if (raw.length < (stride + 1) * height) throw new Error('The PNG data is incomplete');

  const pixels = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    const src = y * (stride + 1) + 1;
    const dst = y * stride;
    for (let x = 0; x < stride; x++) {
      const left = x >= channels ? pixels[dst + x - channels] : 0;
      const up = y > 0 ? pixels[dst - stride + x] : 0;
      const upLeft = y > 0 && x >= channels ? pixels[dst - stride + x - channels] : 0;
      const value = raw[src + x];
      let out;
      switch (filter) {
        case 0: out = value; break;
        case 1: out = value + left; break;
        case 2: out = value + up; break;
        case 3: out = value + ((left + up) >> 1); break;
        case 4: out = value + paeth(left, up, upLeft); break;
        default: throw new Error(`Unsupported PNG filter (${filter})`);
      }
      pixels[dst + x] = out & 0xff;
    }
  }

  const rgba = new Uint8Array(width * height * 4);
  for (let i = 0, p = 0, q = 0; i < width * height; i++, p += channels, q += 4) {
    switch (colorType) {
      case 0: rgba[q] = rgba[q + 1] = rgba[q + 2] = pixels[p]; rgba[q + 3] = 255; break;
      case 2: rgba[q] = pixels[p]; rgba[q + 1] = pixels[p + 1]; rgba[q + 2] = pixels[p + 2]; rgba[q + 3] = 255; break;
      case 3: {
        const index = pixels[p];
        rgba[q] = palette[index * 3];
        rgba[q + 1] = palette[index * 3 + 1];
        rgba[q + 2] = palette[index * 3 + 2];
        rgba[q + 3] = transparency && index < transparency.length ? transparency[index] : 255;
        break;
      }
      case 4: rgba[q] = rgba[q + 1] = rgba[q + 2] = pixels[p]; rgba[q + 3] = pixels[p + 1]; break;
      default: rgba[q] = pixels[p]; rgba[q + 1] = pixels[p + 1]; rgba[q + 2] = pixels[p + 2]; rgba[q + 3] = pixels[p + 3];
    }
  }
  return { width, height, data: rgba };
}

/** Encodes RGBA bytes as a PNG (RGB when nothing is transparent). */
export function encodePng({ width, height, data }, { level = 4 } = {}) {
  let opaque = true;
  for (let i = 3; i < data.length; i += 4) {
    if (data[i] !== 255) {
      opaque = false;
      break;
    }
  }
  const channels = opaque ? 3 : 4;
  const stride = width * channels;
  const raw = Buffer.alloc((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    const row = y * (stride + 1);
    raw[row] = 0;
    let out = row + 1;
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      raw[out++] = data[i];
      raw[out++] = data[i + 1];
      raw[out++] = data[i + 2];
      if (!opaque) raw[out++] = data[i + 3];
    }
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = opaque ? 2 : 6;
  return Buffer.concat([SIGNATURE, chunk('IHDR', header), chunk('IDAT', deflateSync(raw, { level })), chunk('IEND', Buffer.alloc(0))]);
}

/** A white page of the given size, as a PNG. */
export function whitePng(width, height) {
  const data = new Uint8Array(width * height * 4);
  data.fill(255);
  return encodePng({ width, height, data });
}
