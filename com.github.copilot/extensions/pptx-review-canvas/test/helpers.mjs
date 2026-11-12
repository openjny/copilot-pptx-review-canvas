import { rmSync } from 'node:fs';
import fs from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { crc32, deflateRawSync } from 'node:zlib';
import { encodePng } from '../lib/png.mjs';

// The default language is English; a developer's own PPTX_CANVAS_LANG must not change what the tests expect.
delete process.env.PPTX_CANVAS_LANG;

/** A small valid PNG whose colour depends on `seed`, so that different slides have different bytes (and hashes). */
export function makePng(seed = 0, width = 16, height = 9) {
  const data = new Uint8Array(width * height * 4);
  for (let i = 0; i < data.length; i += 4) {
    data[i] = (seed * 37) & 255;
    data[i + 1] = (seed * 91 + 40) & 255;
    data[i + 2] = (seed * 17 + 80) & 255;
    data[i + 3] = 255;
  }
  return Buffer.from(encodePng({ width, height, data }));
}

export const PNG = makePng(0);

export function request(url, { method = 'GET', headers = {}, body } = {}) {
  const target = new URL(url);
  return new Promise((resolve, reject) => {
    const req = http.request({ host: target.hostname, port: target.port, path: target.pathname + target.search, method, headers }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
    });
    req.on('error', reject);
    req.end(body);
  });
}

export const json = (response) => JSON.parse(response.body.toString('utf8'));
export const postJson = (url, payload, headers = {}) => request(url, { method: 'POST', headers: { 'Content-Type': 'application/json', ...headers }, body: JSON.stringify(payload) });

const tempDirs = [];
process.on('exit', () => {
  for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true, maxRetries: 3 });
});

export async function makeTempDir(prefix = 'pptx-viewer-test-') {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), prefix));
  tempDirs.push(dir);
  return dir;
}

export async function makeDeckFile(dir, name = 'deck.pptx', contents = 'not a real deck') {
  const file = path.join(dir, name);
  await fs.writeFile(file, contents);
  return file;
}

/**
 * A ZIP written independently of lib/pptx-zip.mjs, so the tests do not check that code against itself.
 * An entry is `{ name, data, method: 0 | 8, descriptor: boolean }`; `descriptor` puts the sizes after the data (flag bit 3), as PowerPoint does.
 */
export function makeZip(entries, { comment = '' } = {}) {
  const parts = [];
  const central = [];
  let offset = 0;
  for (const { name, data, method = 8, descriptor = false } of entries) {
    const raw = Buffer.from(data);
    const body = method === 8 ? deflateRawSync(raw) : raw;
    const nameBytes = Buffer.from(name, 'utf8');
    const crc = crc32(raw);
    const flags = (descriptor ? 8 : 0) | 0x0800;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0x6000, 10);
    local.writeUInt16LE(0x5a21, 12);
    local.writeUInt32LE(descriptor ? 0 : crc, 14);
    local.writeUInt32LE(descriptor ? 0 : body.length, 18);
    local.writeUInt32LE(descriptor ? 0 : raw.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(4, 28);
    const extra = Buffer.from([0xaa, 0xbb, 0x00, 0x00]);
    parts.push(local, nameBytes, extra, body);
    let length = 30 + nameBytes.length + extra.length + body.length;
    if (descriptor) {
      const tail = Buffer.alloc(16);
      tail.writeUInt32LE(0x08074b50, 0);
      tail.writeUInt32LE(crc, 4);
      tail.writeUInt32LE(body.length, 8);
      tail.writeUInt32LE(raw.length, 12);
      parts.push(tail);
      length += 16;
    }
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(20, 4);
    entry.writeUInt16LE(20, 6);
    entry.writeUInt16LE(flags, 8);
    entry.writeUInt16LE(method, 10);
    entry.writeUInt16LE(0x6000, 12);
    entry.writeUInt16LE(0x5a21, 14);
    entry.writeUInt32LE(crc, 16);
    entry.writeUInt32LE(body.length, 20);
    entry.writeUInt32LE(raw.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt32LE(offset, 42);
    central.push(entry, nameBytes);
    offset += length;
  }
  const directory = Buffer.concat(central);
  const commentBytes = Buffer.from(comment);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(commentBytes.length, 20);
  return Buffer.concat([...parts, directory, end, commentBytes]);
}

export const PRESENTATION_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/**
 * The text of ppt/presentation.xml for slides with the given ids. `sections` is a list of `{ name, ids }` (omit for none).
 */
export function presentationXml(ids, { sections = null, prefix = 'p', customShows = null } = {}) {
  const list = ids.map((id, i) => `<${prefix}:sldId id="${id}" r:id="rId${i + 10}"/>`).join('');
  const shows = customShows
    ? `<${prefix}:custShowLst>${customShows
      .map((numbers, i) => `<${prefix}:custShow name="show${i}" id="${i}"><${prefix}:sldLst>${numbers.map((n) => `<${prefix}:sld r:id="rId${n + 9}"/>`).join('')}</${prefix}:sldLst></${prefix}:custShow>`)
      .join('')}</${prefix}:custShowLst>`
    : '';
  const groups = sections
    ? `<${prefix}:extLst><${prefix}:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${sections
      .map((section, i) => `<p14:section name="${section.name}" id="{0000000${i}-0000-0000-0000-000000000000}">${section.ids.length ? `<p14:sldIdLst>${section.ids.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst>` : '<p14:sldIdLst/>'}</p14:section>`)
      .join('')}</p14:sectionLst></${prefix}:ext></${prefix}:extLst>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${prefix}:presentation ${PRESENTATION_NS}><${prefix}:sldMasterIdLst><${prefix}:sldMasterId id="2147483648" r:id="rId1"/></${prefix}:sldMasterIdLst><${prefix}:sldIdLst>${list}</${prefix}:sldIdLst>  <${prefix}:sldSz cx="12192000" cy="6858000"/>${shows}${groups}</${prefix}:presentation>`;
}

/** A small .pptx-shaped ZIP: a presentation part for slides with `ids`, and some other parts that must come through untouched. */
export function makePptxBuffer(ids, options = {}) {
  return makeZip([
    { name: '[Content_Types].xml', data: '<Types/>', method: 8, descriptor: true },
    { name: '_rels/.rels', data: '<Relationships/>', method: 0 },
    { name: 'ppt/presentation.xml', data: presentationXml(ids, options), method: 8, descriptor: options.descriptor ?? true },
    ...ids.map((id) => ({ name: `ppt/slides/slide${id}.xml`, data: `<slide id="${id}">${'x'.repeat(50 + (id % 7))}</slide>`, method: 8, descriptor: true })),
    { name: 'ppt/media/image1.png', data: makePng(3), method: 0 },
  ], options.zip);
}

export async function makeDeckZip(dir, name, ids, options) {
  const file = path.join(dir, name);
  await fs.writeFile(file, makePptxBuffer(ids, options));
  return file;
}

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = { slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml', notes: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml' };
const relationships = (items) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

/**
 * A .pptx with the parts a real one has around its slides: relationships, a layout and a master, a notes page and a picture of
 * its own for every slide, and a picture they all share. `links` is a list of `[fromId, toId]`: slide `fromId` links to slide `toId`.
 */
export function makePackageBuffer(ids, { sections = null, customShows = null, links = [], zip } = {}) {
  const overrides = ids.map((id) => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="${CT.slide}"/><Override PartName="/ppt/notesSlides/notesSlide${id}.xml" ContentType="${CT.notes}"/>`).join('');
  const perSlide = ids.flatMap((id, i) => [
    { name: `ppt/slides/slide${id}.xml`, data: `<slide id="${id}">${'x'.repeat(50 + (id % 7))}</slide>`, method: i % 2 ? 0 : 8, descriptor: i % 2 === 0 },
    {
      name: `ppt/slides/_rels/slide${id}.xml.rels`,
      data: relationships([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
        ['rId2', 'notesSlide', `../notesSlides/notesSlide${id}.xml`],
        ['rId3', 'image', '../media/shared.png'],
        ['rId4', 'image', `../media/own${id}.png`],
        ...links.filter(([from]) => from === id).map(([, to], k) => [`rId${5 + k}`, 'slide', `slide${to}.xml`]),
      ]),
    },
    { name: `ppt/notesSlides/notesSlide${id}.xml`, data: `<notes id="${id}"/>` },
    { name: `ppt/notesSlides/_rels/notesSlide${id}.xml.rels`, data: relationships([['rId1', 'slide', `../slides/slide${id}.xml`]]) },
    { name: `ppt/media/own${id}.png`, data: makePng(id), method: 0 },
  ]);
  return makeZip([
    { name: '[Content_Types].xml', data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/presentation"/>${overrides}</Types>`, descriptor: true },
    { name: '_rels/.rels', data: relationships([['rId1', 'officeDocument', 'ppt/presentation.xml']]), method: 0 },
    { name: 'ppt/presentation.xml', data: presentationXml(ids, { sections, customShows }), descriptor: true },
    { name: 'ppt/_rels/presentation.xml.rels', data: relationships([['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'], ...ids.map((id, i) => [`rId${i + 10}`, 'slide', `slides/slide${id}.xml`])]) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: '<master/>' },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: relationships([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]) },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: '<layout/>' },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: relationships([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]) },
    { name: 'ppt/media/shared.png', data: makePng(7), method: 0 },
    ...perSlide,
  ], zip);
}

export async function makePackageFile(dir, name, ids, options) {
  const file = path.join(dir, name);
  await fs.writeFile(file, makePackageBuffer(ids, options));
  return file;
}

/**
 * A stand-in for the PowerPoint renderer: writes the files the real one writes.
 * `variant` changes the picture of the slides in `changed` (as if the deck had been edited there).
 */
export function fakeRenderer({ slides = 3, calls = [], failWith = null, delayMs = 0, changed = [], variant = 0, shapes = null } = {}) {
  return async ({ pptx, outDir, onProgress, signal }) => {
    calls.push({ pptx, outDir });
    if (failWith) throw failWith;
    onProgress?.(0, slides);
    const list = [];
    for (let n = 1; n <= slides; n++) {
      if (delayMs) await new Promise((resolve) => setTimeout(resolve, delayMs));
      if (signal?.aborted) {
        const { ViewerError } = await import('../lib/errors.mjs');
        throw new ViewerError('cancelled', 'cancelled');
      }
      const pad = String(n).padStart(4, '0');
      const picture = makePng(n + (changed.includes(n) ? 100 * (variant + 1) : 0));
      await fs.writeFile(path.join(outDir, `s${pad}.png`), picture);
      await fs.writeFile(path.join(outDir, `t${pad}.png`), picture);
      list.push({ n, title: `Title ${n}`, notes: `Notes ${n}\r\nsecond line` });
      onProgress?.(n, slides);
    }
    if (shapes) await fs.writeFile(path.join(outDir, 'shapes.json'), JSON.stringify(shapes));
    return { slideWidth: 960, slideHeight: 540, slides: list };
  };
}
