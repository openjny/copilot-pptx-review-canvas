import fs from 'node:fs/promises';
import path from 'node:path';
import { ViewerError } from './errors.mjs';
import { msg } from './i18n.mjs';
import { findEntry, readEntry, readZipIndex, replaceEntry, validateZip } from './pptx-zip.mjs';
import { hashPath, writeFileAtomic } from './store.mjs';

// Rearranges the slides of a PowerPoint file by rewriting the order in ppt/presentation.xml (and the sections that group the slides).
// No slide, picture or relationship is touched, so the file stays exactly as it was apart from the order.

export const REORDERABLE_EXTENSIONS = new Set(['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm']);
const PRESENTATION_PART = 'ppt/presentation.xml';
const KEEP_BACKUPS = 5;
const LOCK_CODES = new Set(['EPERM', 'EBUSY', 'EACCES']);

export const SLIDE_LIST = /<((?:[\w-]+:)?)sldIdLst\b[^>]*>([\s\S]*?)<\/\1sldIdLst>/;
export const SLIDE_ITEM = /<(?:[\w-]+:)?sldId\b[^>]*?(?:\/>|>\s*<\/(?:[\w-]+:)?sldId>)/g;
export const SECTION_LIST = /<p14:sectionLst\b[\s\S]*?<\/p14:sectionLst>/;
const SECTION = /<p14:section\b[\s\S]*?<\/p14:section>/g;
const SECTION_SLIDES = /<p14:sldIdLst\b[^>]*?(?:\/>|>[\s\S]*?<\/p14:sldIdLst>)/;
const SECTION_ITEM = /<p14:sldId\b[^>]*?\bid="(\d+)"/g;

const unsupported = (detailKey) => new ViewerError('unsupported_pptx', 'error.unsupportedReorder', { detail: msg(detailKey) });

export const isReorderableExtension = (extension) => REORDERABLE_EXTENSIONS.has(String(extension).toLowerCase());

/** The order after moving the item at `from` to `to` (both 1-based): `order[p - 1]` is the old number of the item now at p. */
export function moveOrder(count, from, to) {
  const order = Array.from({ length: count }, (_, i) => i + 1);
  order.splice(to - 1, 0, ...order.splice(from - 1, 1));
  return order;
}

function rewriteSections(block, oldIds, newIds, movedId) {
  const sections = [...block.matchAll(SECTION)].map((match) => ({ text: match[0], ids: [...match[0].matchAll(SECTION_ITEM)].map((item) => item[1]) }));
  if (sections.flatMap((section) => section.ids).join() !== oldIds.join()) throw unsupported('pptx.sectionMismatch');

  const lists = sections.map((section) => section.ids.filter((id) => id !== movedId));
  const position = newIds.indexOf(movedId);
  const anchor = position > 0 ? newIds[position - 1] : newIds[1];
  const target = lists.findIndex((list) => list.includes(anchor));
  if (target < 0) throw unsupported('pptx.sectionTargetMissing');
  lists[target].splice(lists[target].indexOf(anchor) + (position > 0 ? 1 : 0), 0, movedId);
  if (lists.flat().join() !== newIds.join()) throw unsupported('pptx.sectionUpdateFailed');

  let at = 0;
  return block.replace(SECTION, (text) => {
    const before = sections[at].ids;
    const next = lists[at++];
    if (next.join() === before.join()) return text;
    if (!SECTION_SLIDES.test(text)) throw unsupported('pptx.sectionNoList');
    const replacement = next.length ? `<p14:sldIdLst>${next.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst>` : '<p14:sldIdLst/>';
    return text.replace(SECTION_SLIDES, () => replacement);
  });
}

/** Moves one slide in the text of ppt/presentation.xml. Returns the new text and the new order (see `moveOrder`). */
export function reorderPresentationXml(xml, from, to) {
  const listMatch = SLIDE_LIST.exec(xml);
  if (!listMatch || listMatch[1] === 'p14:') throw new ViewerError('invalid_pptx', 'error.slideListMissing');
  const inner = listMatch[2];
  const items = inner.match(SLIDE_ITEM) ?? [];
  if (inner.replace(SLIDE_ITEM, '').trim()) throw unsupported('pptx.slideListUnexpected');
  const count = items.length;
  if (count === 0) throw new ViewerError('invalid_pptx', 'error.noSlides');
  for (const value of [from, to]) {
    if (!Number.isInteger(value) || value < 1 || value > count) throw new ViewerError('slide_out_of_range', 'error.slideOutOfRange', { count });
  }
  const order = moveOrder(count, from, to);
  if (from === to) return { xml, order };

  const idOf = (item) => /\sid="(\d+)"/.exec(item)?.[1];
  const oldIds = items.map(idOf);
  if (oldIds.some((id) => id === undefined) || new Set(oldIds).size !== count) throw unsupported('pptx.slideIdUnreadable');
  const newIds = order.map((old) => oldIds[old - 1]);

  let k = 0;
  const reordered = inner.replace(SLIDE_ITEM, () => items[order[k++] - 1]);
  let next = xml.slice(0, listMatch.index) + listMatch[0].replace(inner, () => reordered) + xml.slice(listMatch.index + listMatch[0].length);

  if (SECTION_LIST.test(next)) next = next.replace(SECTION_LIST, (block) => rewriteSections(block, oldIds, newIds, oldIds[from - 1]));
  else if (/sectionLst/.test(next)) throw unsupported('pptx.sectionFormat');
  return { xml: next, order };
}

async function saveBackup(backupDir, file, original) {
  await fs.mkdir(backupDir, { recursive: true });
  await writeFileAtomic(path.join(backupDir, `${hashPath(file)}${path.extname(file).toLowerCase()}`), original);
  const names = await fs.readdir(backupDir);
  const stats = await Promise.all(names.map(async (name) => ({ name, mtimeMs: (await fs.stat(path.join(backupDir, name)).catch(() => ({ mtimeMs: 0 }))).mtimeMs })));
  stats.sort((a, b) => b.mtimeMs - a.mtimeMs);
  for (const { name } of stats.slice(KEEP_BACKUPS)) await fs.rm(path.join(backupDir, name), { force: true });
}

/**
 * Writes `output` over `file` after keeping `original` in `backupDir` (the latest for each file). The write is atomic; when another
 * app has the file open the error code is `file_locked` and the file is unchanged.
 */
export async function commitRewrite(file, original, output, backupDir) {
  if (backupDir) await saveBackup(backupDir, file, original);
  try {
    await writeFileAtomic(file, output);
  } catch (error) {
    if (LOCK_CODES.has(error.code)) {
      throw new ViewerError('file_locked', 'error.fileLocked');
    }
    throw error;
  }
}

/**
 * Moves the slide at `from` to position `to` in the file itself. The file as it was is kept in `backupDir` (the latest for each file).
 * Returns the new order. The write is atomic; when another app has the file open the error code is `file_locked` and nothing changed.
 */
export async function reorderSlidesInFile(file, from, to, { backupDir } = {}) {
  if (!isReorderableExtension(path.extname(file))) {
    throw new ViewerError('not_reorderable', 'error.legacyNoReorder');
  }
  const original = await fs.readFile(file);
  const index = readZipIndex(original);
  const entry = findEntry(index, PRESENTATION_PART);
  if (!entry) throw new ViewerError('invalid_pptx', 'error.partMissing', { part: PRESENTATION_PART });
  const { xml, order } = reorderPresentationXml(readEntry(original, entry).toString('utf8'), from, to);
  if (from === to) return { order, changed: false };

  const output = replaceEntry(original, index, entry, Buffer.from(xml, 'utf8'));
  const check = validateZip(output);
  const written = findEntry(check, PRESENTATION_PART);
  if (!written || readEntry(output, written).toString('utf8') !== xml || check.entries.length !== index.entries.length) {
    throw new ViewerError('internal_error', 'error.verifyFailed');
  }

  await commitRewrite(file, original, output, backupDir);
  return { order, changed: true };
}
