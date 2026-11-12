import fs from 'node:fs/promises';
import path from 'node:path';
import { ViewerError } from './errors.mjs';
import { tEn } from './i18n.mjs';
import { whitePng } from './png.mjs';
import { normalizePathKey, writeFileAtomic } from './store.mjs';

export const SKETCHBOOK_EXTENSION = '.sketch';
const SKETCHBOOK_FORMAT = 'pptx-slide-viewer.sketchbook';
const MAX_NAME_LENGTH = 60;
const RESERVED_NAMES = /^(con|prn|aux|nul|com\d|lpt\d)$/i;
// A sketchbook has no .pptx, so it is shown as a deck of one white 16:9 page (the size PowerPoint uses by default).
const PAGE_WIDTH = 960;
const PAGE_HEIGHT = 540;

export function sketchbooksDir(root) {
  return path.join(root, 'sketchbooks');
}

/** A file-name-safe name; empty input gives a name made from the time (in the language of `t`), so each new sketchbook is separate. */
export function sketchbookName(input, now = new Date(), t = tEn) {
  const cleaned = String(input ?? '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, MAX_NAME_LENGTH);
  if (cleaned) return RESERVED_NAMES.test(cleaned) ? `${cleaned}-` : cleaned;
  const pad = (value) => String(value).padStart(2, '0');
  return t('sketchbook.defaultName', { stamp: `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}` });
}

function sourceFor(file) {
  return { path: file, name: path.basename(file, SKETCHBOOK_EXTENSION), extension: SKETCHBOOK_EXTENSION, size: 0, mtimeMs: 0, sketchbook: true };
}

/** Creates the sketchbook, or finds it again when the name is already taken (the same name continues the same sketchbook). */
export async function createSketchbook(root, name, t = tEn) {
  const file = path.join(sketchbooksDir(root), `${sketchbookName(name, new Date(), t)}${SKETCHBOOK_EXTENSION}`);
  const exists = await fs.stat(file).then((stat) => stat.isFile(), () => false);
  if (!exists) await writeFileAtomic(file, `${JSON.stringify({ format: SKETCHBOOK_FORMAT, version: 1, createdAt: Date.now() })}\n`);
  return sourceFor(file);
}

/** The sketchbook `input` names, or null when it is not a sketchbook (so the caller treats it as a PowerPoint file). */
export async function resolveSketchbook(root, input) {
  const raw = String(input ?? '').trim().replace(/^"(.*)"$/, '$1').trim();
  if (!raw || !path.isAbsolute(raw) || path.extname(raw).toLowerCase() !== SKETCHBOOK_EXTENSION) return null;
  const file = path.normalize(raw);
  if (normalizePathKey(path.dirname(file)) !== normalizePathKey(sketchbooksDir(root))) return null;
  const exists = await fs.stat(file).then((stat) => stat.isFile(), () => false);
  if (!exists) throw new ViewerError('file_not_found', 'error.sketchbookNotFound', { path: file });
  return sourceFor(file);
}

/**
 * Writes the one white page and its thumbnail, in the shape the PowerPoint renderer returns. The page has no title of its own: the
 * viewer shows it under the name of the language in use, so the cached manifest does not depend on the language.
 */
export async function renderSketchbook({ outDir, width, thumbWidth }) {
  const height = Math.round((width * PAGE_HEIGHT) / PAGE_WIDTH);
  const thumbHeight = Math.round((thumbWidth * PAGE_HEIGHT) / PAGE_WIDTH);
  await fs.writeFile(path.join(outDir, 's0001.png'), whitePng(width, height));
  await fs.writeFile(path.join(outDir, 't0001.png'), whitePng(thumbWidth, thumbHeight));
  return { slideWidth: PAGE_WIDTH, slideHeight: PAGE_HEIGHT, slides: [{ n: 1, title: '', notes: '' }] };
}
