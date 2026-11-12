import { crc32, deflateRawSync, inflateRawSync } from 'node:zlib';
import { ViewerError } from './errors.mjs';
import { msg } from './i18n.mjs';

// Reads a .pptx (a ZIP) and replaces one entry while every other entry stays byte for byte the same.
// Only what a PowerPoint file needs is supported: a single-disk ZIP, no ZIP64, no encryption, stored or deflated entries.

const EOCD = 0x06054b50;
const ZIP64_LOCATOR = 0x07064b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;
const EOCD_LENGTH = 22;
const MAX_COMMENT = 0xffff;
const FLAG_ENCRYPTED = 0x0001;
const FLAG_DESCRIPTOR = 0x0008;

const invalid = (detailKey, params) => new ViewerError('invalid_pptx', 'error.invalidPptx', { detail: msg(detailKey, params) });

/** Finds the entries of the ZIP without decompressing anything. */
export function readZipIndex(buf) {
  if (buf.length < EOCD_LENGTH) throw invalid('zip.tooSmall');
  let eocd = -1;
  for (let i = buf.length - EOCD_LENGTH, stop = Math.max(0, buf.length - EOCD_LENGTH - MAX_COMMENT); i >= stop; i--) {
    if (buf.readUInt32LE(i) === EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw invalid('zip.notZip');
  if (eocd >= 20 && buf.readUInt32LE(eocd - 20) === ZIP64_LOCATOR) throw invalid('zip.zip64');
  const disk = buf.readUInt16LE(eocd + 4);
  const cdDisk = buf.readUInt16LE(eocd + 6);
  const onDisk = buf.readUInt16LE(eocd + 8);
  const count = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  const commentLength = buf.readUInt16LE(eocd + 20);
  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) throw invalid('zip.zip64');
  if (disk !== 0 || cdDisk !== 0 || onDisk !== count) throw invalid('zip.split');
  if (cdOffset + cdSize > eocd) throw invalid('zip.directoryPosition');

  const entries = [];
  let at = cdOffset;
  for (let i = 0; i < count; i++) {
    if (at + 46 > cdOffset + cdSize || buf.readUInt32LE(at) !== CENTRAL) throw invalid('zip.directoryCorrupt');
    const nameLength = buf.readUInt16LE(at + 28);
    const length = 46 + nameLength + buf.readUInt16LE(at + 30) + buf.readUInt16LE(at + 32);
    if (at + length > cdOffset + cdSize) throw invalid('zip.directoryCorrupt');
    const entry = {
      name: buf.toString('utf8', at + 46, at + 46 + nameLength),
      flags: buf.readUInt16LE(at + 8),
      method: buf.readUInt16LE(at + 10),
      time: buf.readUInt16LE(at + 12),
      date: buf.readUInt16LE(at + 14),
      crc: buf.readUInt32LE(at + 16),
      compressedSize: buf.readUInt32LE(at + 20),
      size: buf.readUInt32LE(at + 24),
      localOffset: buf.readUInt32LE(at + 42),
      cenOffset: at,
      cenLength: length,
      nameLength,
    };
    if (entry.compressedSize === 0xffffffff || entry.size === 0xffffffff || entry.localOffset === 0xffffffff) throw invalid('zip.zip64');
    entries.push(entry);
    at += length;
  }
  return { entries, cdOffset, cdSize, eocdOffset: eocd, eocdLength: EOCD_LENGTH + commentLength };
}

export function findEntry(index, name) {
  const found = index.entries.filter((entry) => entry.name === name);
  if (found.length > 1) throw invalid('zip.duplicate', { name });
  return found[0] ?? null;
}

function dataRange(buf, entry) {
  const at = entry.localOffset;
  if (at + 30 > buf.length || buf.readUInt32LE(at) !== LOCAL) throw invalid('zip.entryPosition', { name: entry.name });
  const start = at + 30 + buf.readUInt16LE(at + 26) + buf.readUInt16LE(at + 28);
  const end = start + entry.compressedSize;
  if (end > buf.length) throw invalid('zip.entryTruncated', { name: entry.name });
  return { start, end };
}

/** The uncompressed bytes of an entry, checked against its CRC. */
export function readEntry(buf, entry) {
  if (entry.flags & FLAG_ENCRYPTED) throw invalid('zip.entryEncrypted', { name: entry.name });
  const { start, end } = dataRange(buf, entry);
  const raw = buf.subarray(start, end);
  let data;
  if (entry.method === 0) data = Buffer.from(raw);
  else if (entry.method === 8) {
    try {
      data = inflateRawSync(raw);
    } catch {
      throw invalid('zip.entryInflate', { name: entry.name });
    }
  } else throw invalid('zip.entryMethod', { name: entry.name });
  if (data.length !== entry.size || crc32(data) !== entry.crc) throw invalid('zip.entryCorrupt', { name: entry.name });
  return data;
}

/**
 * Returns the ZIP with `entry` holding `data`. The other entries are copied as they are, so only the central directory
 * (the offsets of the entries after the changed one) is rewritten.
 */
export function replaceEntry(buf, index, entry, data) {
  const compressed = deflateRawSync(data, { level: 6 });
  const crc = crc32(data);
  const flags = entry.flags & ~FLAG_DESCRIPTOR;
  const start = entry.localOffset;
  const end = index.entries.map((other) => other.localOffset).filter((offset) => offset > start).sort((a, b) => a - b)[0] ?? index.cdOffset;
  const nameBytes = buf.subarray(entry.cenOffset + 46, entry.cenOffset + 46 + entry.nameLength);

  const local = Buffer.alloc(30 + entry.nameLength);
  local.writeUInt32LE(LOCAL, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(flags, 6);
  local.writeUInt16LE(8, 8);
  local.writeUInt16LE(entry.time, 10);
  local.writeUInt16LE(entry.date, 12);
  local.writeUInt32LE(crc, 14);
  local.writeUInt32LE(compressed.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(entry.nameLength, 26);
  local.writeUInt16LE(0, 28);
  nameBytes.copy(local, 30);

  const delta = local.length + compressed.length - (end - start);
  const central = Buffer.from(buf.subarray(index.cdOffset, index.cdOffset + index.cdSize));
  for (const other of index.entries) {
    const at = other.cenOffset - index.cdOffset;
    if (other === entry) {
      central.writeUInt16LE(flags, at + 8);
      central.writeUInt16LE(8, at + 10);
      central.writeUInt32LE(crc, at + 16);
      central.writeUInt32LE(compressed.length, at + 20);
      central.writeUInt32LE(data.length, at + 24);
    } else if (other.localOffset > start) {
      central.writeUInt32LE(other.localOffset + delta, at + 42);
    }
  }
  const tail = Buffer.from(buf.subarray(index.eocdOffset, index.eocdOffset + index.eocdLength));
  tail.writeUInt32LE(index.cdOffset + delta, 16);

  return Buffer.concat([
    buf.subarray(0, start),
    local,
    compressed,
    buf.subarray(end, index.cdOffset),
    central,
    buf.subarray(index.cdOffset + index.cdSize, index.eocdOffset),
    tail,
  ]);
}

/**
 * Returns the ZIP without the entries named in `names`. The other entries are copied as they are; only the central
 * directory (which lists them) and the offsets after each removed entry change.
 */
export function removeEntries(buf, index, names) {
  const drop = new Set(names);
  const sorted = [...index.entries].sort((a, b) => a.localOffset - b.localOffset);
  const blockEnd = new Map(sorted.map((entry, i) => [entry, sorted[i + 1]?.localOffset ?? index.cdOffset]));
  const head = buf.subarray(0, sorted[0]?.localOffset ?? index.cdOffset);
  const body = [head];
  const moved = new Map();
  let offset = head.length;
  for (const entry of sorted) {
    if (drop.has(entry.name)) continue;
    const block = buf.subarray(entry.localOffset, blockEnd.get(entry));
    moved.set(entry, offset);
    body.push(block);
    offset += block.length;
  }
  const kept = index.entries.filter((entry) => !drop.has(entry.name));
  const central = Buffer.concat(
    kept.map((entry) => {
      const record = Buffer.from(buf.subarray(entry.cenOffset, entry.cenOffset + entry.cenLength));
      record.writeUInt32LE(moved.get(entry), 42);
      return record;
    }),
  );
  const tail = Buffer.from(buf.subarray(index.eocdOffset, index.eocdOffset + index.eocdLength));
  tail.writeUInt16LE(kept.length, 8);
  tail.writeUInt16LE(kept.length, 10);
  tail.writeUInt32LE(central.length, 12);
  tail.writeUInt32LE(offset, 16);
  return Buffer.concat([...body, central, tail]);
}

/**
 * Checks the structure of a ZIP without decompressing it: every entry starts with its own local header, its data lies inside
 * the file, and the entries do not overlap. Throws when something is off.
 */
export function validateZip(buf) {
  const index = readZipIndex(buf);
  const spans = [];
  for (const entry of index.entries) {
    const at = entry.localOffset;
    if (at + 30 > index.cdOffset || buf.readUInt32LE(at) !== LOCAL) throw invalid('zip.entryPosition', { name: entry.name });
    const nameLength = buf.readUInt16LE(at + 26);
    if (buf.toString('utf8', at + 30, at + 30 + nameLength) !== entry.name) throw invalid('zip.entryHeader', { name: entry.name });
    const { end } = dataRange(buf, entry);
    if (end > index.cdOffset) throw invalid('zip.entryOverlapsDirectory', { name: entry.name });
    spans.push([at, end]);
  }
  spans.sort((a, b) => a[0] - b[0]);
  for (let i = 1; i < spans.length; i++) if (spans[i][0] < spans[i - 1][1]) throw invalid('zip.entriesOverlap');
  return index;
}
