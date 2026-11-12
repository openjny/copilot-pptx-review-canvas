import fs from 'node:fs/promises';
import path from 'node:path';
import { ViewerError } from './errors.mjs';
import { msg } from './i18n.mjs';
import { SECTION_LIST, SLIDE_ITEM, SLIDE_LIST, commitRewrite, isReorderableExtension } from './pptx-reorder.mjs';
import { findEntry, readEntry, readZipIndex, removeEntries, replaceEntry, validateZip } from './pptx-zip.mjs';

// Deletes a slide from a PowerPoint file: the slide leaves the slide list (and the sections and custom shows that name it),
// its relationship is dropped, and the parts that only that slide used (its notes, its pictures) are removed from the ZIP.
// Everything else, including the layouts and pictures that other slides share, is copied as it is.

const PRESENTATION_PART = 'ppt/presentation.xml';
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';
const CUSTOM_SHOWS = /<(?:[\w-]+:)?custShowLst\b[\s\S]*?<\/(?:[\w-]+:)?custShowLst>/;
const RELATIONSHIP = /<Relationship\b[^>]*?(?:\/>|>\s*<\/Relationship>)/g;
const OVERRIDE = /<Override\b[^>]*?(?:\/>|>\s*<\/Override>)/g;

const unsupported = (detailKey) => new ViewerError('unsupported_pptx', 'error.unsupportedDelete', { detail: msg(detailKey) });
const escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const unescapeXml = (text) => text.replace(/&(lt|gt|quot|apos|amp);/g, (_, name) => ({ lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' })[name]);

function attributes(tag) {
  const found = {};
  for (const match of tag.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) found[match[1]] = unescapeXml(match[2] ?? match[3]);
  return found;
}

/**
 * Takes the slide at `slide` (1-based) out of the text of ppt/presentation.xml. Returns the new text, the relationship id the
 * slide is stored under, and `order`: the old number of each slide that is left, in the new order.
 */
export function removeSlideFromPresentationXml(xml, slide) {
  const listMatch = SLIDE_LIST.exec(xml);
  if (!listMatch || listMatch[1] === 'p14:') throw new ViewerError('invalid_pptx', 'error.slideListMissing');
  const inner = listMatch[2];
  const items = inner.match(SLIDE_ITEM) ?? [];
  if (inner.replace(SLIDE_ITEM, '').trim()) throw unsupported('pptx.slideListUnexpected');
  const count = items.length;
  if (count === 0) throw new ViewerError('invalid_pptx', 'error.noSlides');
  if (!Number.isInteger(slide) || slide < 1 || slide > count) throw new ViewerError('slide_out_of_range', 'error.slideOutOfRange', { count });
  if (count === 1) throw new ViewerError('last_slide', 'error.lastSlide');

  const id = /\sid="(\d+)"/.exec(items[slide - 1])?.[1];
  const rId = /\s[\w-]+:id="([^"]+)"/.exec(items[slide - 1])?.[1];
  if (!id || !rId) throw unsupported('pptx.slideIdUnreadable');

  let k = 0;
  const kept = inner.replace(SLIDE_ITEM, (item) => (++k === slide ? '' : item));
  let next = xml.slice(0, listMatch.index) + listMatch[0].replace(inner, () => kept) + xml.slice(listMatch.index + listMatch[0].length);

  if (SECTION_LIST.test(next)) {
    const entry = new RegExp(`<p14:sldId\\b[^>]*?\\bid="${id}"[^>]*?(?:/>|>\\s*</p14:sldId>)`, 'g');
    next = next.replace(SECTION_LIST, (block) => {
      if ((block.match(entry)?.length ?? 0) !== 1) throw unsupported('pptx.sectionMismatch');
      return block.replace(entry, '').replace(/<p14:sldIdLst>\s*<\/p14:sldIdLst>/g, '<p14:sldIdLst/>');
    });
  } else if (/sectionLst/.test(next)) throw unsupported('pptx.sectionFormat');

  const shown = new RegExp(`<(?:[\\w-]+:)?sld\\b[^>]*?\\br:id="${escapeRegExp(rId)}"[^>]*?/>`, 'g');
  next = next.replace(CUSTOM_SHOWS, (block) => block.replace(shown, ''));
  if (next.includes(`"${rId}"`)) throw unsupported('pptx.slideReferenced');

  return { xml: next, rId, order: Array.from({ length: count }, (_, i) => i + 1).filter((n) => n !== slide) };
}

function readRelationships(xml) {
  return [...xml.matchAll(RELATIONSHIP)].map((match) => {
    const found = attributes(match[0]);
    return { id: found.Id, type: found.Type ?? '', target: found.Target ?? '', external: found.TargetMode === 'External' };
  });
}

/** Takes the relationship `rId` (which must be a slide) out of the text of a .rels part. Returns the new text and where it pointed. */
function removeRelationship(xml, rId) {
  let found = null;
  const next = xml.replace(RELATIONSHIP, (tag) => {
    const attrs = attributes(tag);
    if (attrs.Id !== rId) return tag;
    found = attrs;
    return '';
  });
  if (!found) throw new ViewerError('invalid_pptx', 'error.slideRelMissing');
  if (!/\/slide$/.test(found.Type ?? '') || found.TargetMode === 'External') throw unsupported('pptx.slideRelType');
  return { xml: next, target: found.Target };
}

const relsPartOf = (part) => (part === '' ? '_rels/.rels' : path.posix.join(path.posix.dirname(part), '_rels', `${path.posix.basename(part)}.rels`));

function resolveTarget(source, target) {
  let clean = target.split('#')[0];
  try {
    clean = decodeURI(clean);
  } catch {
    // A target that is not valid percent-encoding is used as it is.
  }
  return clean.startsWith('/') ? clean.slice(1) : path.posix.join(path.posix.dirname(source || '.'), clean);
}

/** Every part that can be reached from the package's root by following relationships. `relationshipsOf(relsPart)` lists those of one .rels part. */
function reachableParts(relationshipsOf) {
  const seen = new Set();
  const queue = [''];
  while (queue.length) {
    const source = queue.pop();
    for (const rel of relationshipsOf(relsPartOf(source))) {
      if (rel.external) continue;
      const part = resolveTarget(source, rel.target);
      if (seen.has(part)) continue;
      seen.add(part);
      queue.push(part);
    }
  }
  return seen;
}

/**
 * Works out what deleting the slide at `slide` changes in the package: the new text of the three parts that mention it, and the
 * parts to remove (the slide, its .rels, and everything that was reachable only through it).
 */
export function planSlideDeletion(buf, index, slide) {
  const text = (name) => {
    const entry = findEntry(index, name);
    if (!entry) throw new ViewerError('invalid_pptx', 'error.partMissing', { part: name });
    return readEntry(buf, entry).toString('utf8');
  };
  const presentation = removeSlideFromPresentationXml(text(PRESENTATION_PART), slide);
  const rels = removeRelationship(text(PRESENTATION_RELS), presentation.rId);
  const slidePart = resolveTarget(PRESENTATION_PART, rels.target);

  const cache = new Map();
  const stored = (name) => {
    if (!cache.has(name)) {
      const entry = findEntry(index, name);
      cache.set(name, entry ? readRelationships(readEntry(buf, entry).toString('utf8')) : []);
    }
    return cache.get(name);
  };
  const changed = readRelationships(rels.xml);
  const before = reachableParts(stored);
  const after = reachableParts((name) => (name === PRESENTATION_RELS ? changed : stored(name)));
  if (after.has(slidePart)) {
    throw new ViewerError('slide_linked', 'error.slideLinked');
  }

  const names = new Set(index.entries.map((entry) => entry.name));
  const remove = [...before].filter((part) => !after.has(part) && names.has(part));
  if (!remove.includes(slidePart)) throw unsupported('pptx.slidePartMissing');
  for (const part of [...remove]) {
    const own = relsPartOf(part);
    if (names.has(own)) remove.push(own);
  }

  const gone = new Set(remove.map((name) => `/${name}`));
  const contentTypes = text(CONTENT_TYPES).replace(OVERRIDE, (tag) => (gone.has(attributes(tag).PartName) ? '' : tag));
  return { presentation: presentation.xml, rels: rels.xml, contentTypes, remove, order: presentation.order, reachable: after, names };
}

/**
 * Deletes the slide at `slide` from the file itself. The file as it was is kept in `backupDir` (the latest for each file).
 * Returns the old numbers of the slides that are left, in order. The write is atomic; when another app has the file open the
 * error code is `file_locked` and nothing changed.
 */
export async function deleteSlideInFile(file, slide, { backupDir } = {}) {
  if (!isReorderableExtension(path.extname(file))) {
    throw new ViewerError('not_reorderable', 'error.legacyNoDelete');
  }
  const original = await fs.readFile(file);
  const index = readZipIndex(original);
  const plan = planSlideDeletion(original, index, slide);

  let output = original;
  let current = index;
  for (const [name, content] of [[PRESENTATION_PART, plan.presentation], [PRESENTATION_RELS, plan.rels], [CONTENT_TYPES, plan.contentTypes]]) {
    output = replaceEntry(output, current, findEntry(current, name), Buffer.from(content, 'utf8'));
    current = readZipIndex(output);
  }
  output = removeEntries(output, current, plan.remove);

  const check = validateZip(output);
  const written = findEntry(check, PRESENTATION_PART);
  const names = new Set(check.entries.map((entry) => entry.name));
  const intact =
    written &&
    readEntry(output, written).toString('utf8') === plan.presentation &&
    check.entries.length === index.entries.length - plan.remove.length &&
    plan.remove.every((name) => !names.has(name)) &&
    [...plan.reachable].every((part) => !plan.names.has(part) || names.has(part));
  if (!intact) throw new ViewerError('internal_error', 'error.verifyFailed');

  await commitRewrite(file, original, output, backupDir);
  return { order: plan.order, changed: true, removedParts: plan.remove.length };
}
