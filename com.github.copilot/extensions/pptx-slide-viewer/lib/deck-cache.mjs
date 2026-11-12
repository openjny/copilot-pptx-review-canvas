import { createHash, randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { ViewerError } from './errors.mjs';
import { createI18n, msg } from './i18n.mjs';
import { renderSketchbook } from './sketchbook.mjs';
import { normalizePathKey, readRecord, samePath, writeRecord } from './store.mjs';

/** Bump when the renderer changes what it writes, so older caches are not reused. */
export const RENDER_VERSION = 2;
export const IMAGE_WIDTH = 1600;
export const THUMB_WIDTH = 360;
export const ALLOWED_EXTENSIONS = new Set(['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm', '.ppt', '.pps', '.pot']);
const MAX_SOURCE_BYTES = 1024 ** 3;
const MANIFEST_FORMAT = 'pptx-slide-viewer.deck';
const MANIFEST_VERSION = 1;
const KEY_PATTERN = /^[0-9a-f]{24}$/;
const KEEP_DECKS = 30;
const SUPERSEDED_GRACE_MS = 10 * 60 * 1000;
const STALE_WORK_MS = 24 * 60 * 60 * 1000;

export function pad(n) {
  return String(n).padStart(4, '0');
}

/** Checks the file a caller named and returns its identity. Relative paths resolve against `baseDir`. */
export async function resolveDeckSource(input, baseDir) {
  const raw = String(input ?? '').trim().replace(/^"(.*)"$/, '$1').trim();
  if (!raw) throw new ViewerError('invalid_path', 'error.pathEmpty');
  let full;
  if (path.isAbsolute(raw)) full = path.normalize(raw);
  else if (baseDir) full = path.resolve(baseDir, raw);
  else throw new ViewerError('invalid_path', 'error.pathNotAbsolute');
  const extension = path.extname(full).toLowerCase();
  if (!ALLOWED_EXTENSIONS.has(extension)) {
    throw new ViewerError('unsupported_file', 'error.unsupportedFile', { extensions: [...ALLOWED_EXTENSIONS] });
  }
  let stat;
  try {
    stat = await fs.stat(full);
  } catch (error) {
    if (error.code === 'ENOENT' || error.code === 'ENOTDIR') throw new ViewerError('file_not_found', 'error.fileNotFound', { path: full });
    throw new ViewerError('file_unreadable', 'error.fileUnreadable', { detail: error.message });
  }
  if (!stat.isFile()) throw new ViewerError('file_not_found', 'error.notAFile', { path: full });
  if (stat.size > MAX_SOURCE_BYTES) throw new ViewerError('file_too_large', 'error.fileTooLarge');
  return { path: full, name: path.basename(full), extension, size: stat.size, mtimeMs: Math.floor(stat.mtimeMs) };
}

export function deckKey(source) {
  const identity = [RENDER_VERSION, IMAGE_WIDTH, THUMB_WIDTH, normalizePathKey(source.path), source.size, source.mtimeMs].join('|');
  return createHash('sha256').update(identity).digest('hex').slice(0, 24);
}

function validateManifest(data) {
  if (!data || typeof data !== 'object') return msg('record.notObject');
  if (typeof data.key !== 'string' || !KEY_PATTERN.test(data.key)) return msg('manifest.keyInvalid');
  if (!data.source || typeof data.source.path !== 'string') return msg('manifest.sourcePathMissing');
  if (!Number.isInteger(data.slideCount) || data.slideCount < 1) return msg('manifest.slideCountInvalid');
  if (!Array.isArray(data.slides) || data.slides.length !== data.slideCount) return msg('manifest.slidesMismatch');
  for (const [index, slide] of data.slides.entries()) {
    if (!slide || slide.n !== index + 1 || typeof slide.title !== 'string' || typeof slide.notes !== 'string') return msg('manifest.slideShape', { index });
    if (typeof slide.hash !== 'string' || slide.hash.length < 8) return msg('manifest.slideHashMissing', { index });
  }
  if (!Number.isFinite(data.slideWidth) || !Number.isFinite(data.slideHeight) || data.slideWidth <= 0 || data.slideHeight <= 0) return msg('manifest.sizeInvalid');
  return null;
}

function buildManifest(key, source, raw, hashes) {
  const slides = Array.isArray(raw?.slides) ? raw.slides : [];
  const data = {
    key,
    ...(source.sketchbook ? { kind: 'sketchbook' } : {}),
    source: { path: source.path, size: source.size, mtimeMs: source.mtimeMs },
    renderedAt: Date.now(),
    slideWidth: Number(raw?.slideWidth),
    slideHeight: Number(raw?.slideHeight),
    imageWidth: IMAGE_WIDTH,
    thumbWidth: THUMB_WIDTH,
    slideCount: slides.length,
    slides: slides.map((slide, index) => ({
      n: index + 1,
      title: String(slide?.title ?? '').replace(/\s+/g, ' ').trim(),
      notes: String(slide?.notes ?? '').replace(/\r\n|\r|\v/g, '\n').trim(),
      hash: hashes[index] ?? '',
    })),
  };
  const problem = validateManifest(data);
  if (problem) throw new ViewerError('render_failed', 'error.renderResultInvalid', { detail: problem });
  return data;
}

async function exists(target) {
  return fs.access(target).then(() => true, () => false);
}

/** Renders decks into a cache keyed by the file's identity, so reopening an unchanged deck is instant. */
export class DeckCache {
  constructor({ root, renderer, log = () => {}, i18n = createI18n() }) {
    this.root = root;
    this.decksDir = path.join(root, 'decks');
    this.workDir = path.join(root, 'work');
    this.renderer = renderer;
    this.log = log;
    this.i18n = i18n;
    this._inFlight = new Map();
    this._pins = new Map();
  }

  /** Keeps the render under `key` from being cleaned up (an undo may need its pictures). Every `pin` needs one `unpin`. */
  pin(key) {
    this._pins.set(key, (this._pins.get(key) ?? 0) + 1);
  }

  unpin(key) {
    const count = (this._pins.get(key) ?? 0) - 1;
    if (count > 0) this._pins.set(key, count);
    else this._pins.delete(key);
  }

  keyDir(key) {
    return path.join(this.decksDir, key);
  }

  /** The shapes of each slide (name, kind, box in points, text) that the renderer recorded, or null when there are none. */
  async readShapes(key) {
    try {
      const parsed = JSON.parse(await fs.readFile(path.join(this.keyDir(key), 'shapes.json'), 'utf8'));
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
    } catch {
      return null;
    }
  }

  /** Returns the deck's manifest, rendering it when it is not cached (or when `force` is set). */
  async resolve(source, { force = false, signal, onProgress } = {}) {
    const key = deckKey(source);
    if (!force) {
      const cached = await this._readCached(key);
      if (cached) return cached;
    }
    const flightKey = `${key}:${force ? 'force' : 'normal'}`;
    let flight = this._inFlight.get(flightKey);
    if (!flight) {
      flight = { flightKey, waiters: 0, listeners: new Set(), controller: new AbortController(), promise: null };
      const current = flight;
      flight.promise = this._render(source, key, {
        force,
        signal: current.controller.signal,
        onProgress: (done, total) => {
          for (const listener of current.listeners) listener(done, total);
        },
      }).finally(() => this._leaveFlight(current));
      this._inFlight.set(flightKey, flight);
    }
    return this._join(flight, { signal, onProgress });
  }

  _leaveFlight(flight) {
    if (this._inFlight.get(flight.flightKey) === flight) this._inFlight.delete(flight.flightKey);
  }

  /**
   * Waits for a shared render. The render belongs to all its waiters: one waiter giving up does not stop it, and it is
   * cancelled only when the last waiter has given up.
   */
  _join(flight, { signal, onProgress }) {
    flight.waiters += 1;
    if (onProgress) flight.listeners.add(onProgress);
    return new Promise((resolve, reject) => {
      let settled = false;
      const finish = (action) => {
        if (settled) return;
        settled = true;
        if (onProgress) flight.listeners.delete(onProgress);
        signal?.removeEventListener('abort', onAbort);
        flight.waiters -= 1;
        action();
      };
      const onAbort = () => finish(() => {
        if (flight.waiters === 0) {
          this._leaveFlight(flight);
          flight.controller.abort();
        }
        reject(new ViewerError('cancelled', 'error.cancelled'));
      });
      if (signal?.aborted) return onAbort();
      signal?.addEventListener('abort', onAbort, { once: true });
      flight.promise.then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
    });
  }

  /**
   * Publishes the render of `source` for a deck whose slides were only rearranged (or some of them deleted): the pictures did not
   * change, so they are copied under the new order instead of being drawn again. `order[p - 1]` is the old number of the slide now
   * at position p. Pass `{ dropping: true }` when `order` leaves some slides out (it then lists only the slides that are kept).
   * While it runs, anyone else asking for the same file waits for it instead of starting a render of their own.
   */
  derive(fromKey, source, order, { dropping = false } = {}) {
    const key = deckKey(source);
    const flightKey = `${key}:normal`;
    const flight = { flightKey, waiters: 0, listeners: new Set(), controller: new AbortController(), promise: null };
    flight.promise = this._derive(fromKey, source, key, order, dropping).finally(() => this._leaveFlight(flight));
    if (!this._inFlight.has(flightKey)) this._inFlight.set(flightKey, flight);
    return flight.promise;
  }

  async _derive(fromKey, source, key, order, dropping) {
    const previous = await this._readCached(fromKey);
    if (!previous) throw new ViewerError('render_failed', 'error.deriveSourceMissing');
    const before = previous.slideCount;
    const sorted = Array.isArray(order) ? [...order].sort((a, b) => a - b) : [];
    const valid = dropping
      ? sorted.length >= 1 && sorted.length < before && sorted.every((value, i) => Number.isInteger(value) && value >= 1 && value <= before && value !== sorted[i - 1])
      : sorted.length === before && sorted.every((value, i) => value === i + 1);
    if (!valid) throw new ViewerError('invalid_input', 'error.invalidOrder');
    const count = order.length;
    const work = path.join(this.workDir, randomUUID());
    await fs.mkdir(work, { recursive: true });
    try {
      const from = this.keyDir(fromKey);
      for (let p = 1; p <= count; p++) {
        for (const prefix of ['s', 't']) await fs.copyFile(path.join(from, `${prefix}${pad(order[p - 1])}.png`), path.join(work, `${prefix}${pad(p)}.png`));
      }
      const shapes = await this.readShapes(fromKey);
      if (shapes) {
        const moved = {};
        order.forEach((old, i) => {
          if (shapes[String(old)]) moved[String(i + 1)] = shapes[String(old)];
        });
        await fs.writeFile(path.join(work, 'shapes.json'), JSON.stringify(moved));
      }
      const manifest = {
        ...previous,
        key,
        source: { path: source.path, size: source.size, mtimeMs: source.mtimeMs },
        renderedAt: Date.now(),
        slideCount: count,
        slides: order.map((old, i) => ({ ...previous.slides[old - 1], n: i + 1 })),
      };
      const problem = validateManifest(manifest);
      if (problem) throw new ViewerError('render_failed', 'error.rewriteResultInvalid', { detail: problem });
      await writeRecord(path.join(work, 'manifest.json'), MANIFEST_FORMAT, MANIFEST_VERSION, manifest);
      const published = await this._publish(work, key, false);
      this._prune(source, key).catch((error) => this.log(this.i18n.t('log.cachePruneFailed', { detail: error.message })));
      return published ?? manifest;
    } finally {
      await fs.rm(work, { recursive: true, force: true }).catch(() => {});
    }
  }

  async _readCached(key) {
    const dir = this.keyDir(key);
    const manifestFile = path.join(dir, 'manifest.json');
    let record;
    try {
      record = await readRecord(manifestFile, { format: MANIFEST_FORMAT, version: MANIFEST_VERSION, validate: validateManifest });
    } catch (error) {
      this.log(this.i18n.t('log.cacheUnreadable', { detail: error.message }));
      return null;
    }
    if (record.status !== 'ok' || record.data.key !== key) {
      if (record.status === 'invalid') this.log(this.i18n.t('log.cacheInvalid', { reason: record.reason }));
      return null;
    }
    for (let n = 1; n <= record.data.slideCount; n++) {
      if (!(await exists(path.join(dir, `s${pad(n)}.png`))) || !(await exists(path.join(dir, `t${pad(n)}.png`)))) {
        this.log(this.i18n.t('log.cacheImagesMissing'));
        return null;
      }
    }
    const now = new Date();
    await fs.utimes(manifestFile, now, now).catch(() => {});
    return record.data;
  }

  async _render(source, key, { force, signal, onProgress }) {
    const work = path.join(this.workDir, randomUUID());
    await fs.mkdir(work, { recursive: true });
    try {
      const copy = source.sketchbook ? null : path.join(work, `source${source.extension}`);
      if (copy) await fs.copyFile(source.path, copy);
      const raw = source.sketchbook
        ? await renderSketchbook({ outDir: work, width: IMAGE_WIDTH, thumbWidth: THUMB_WIDTH })
        : await this.renderer({ pptx: copy, outDir: work, width: IMAGE_WIDTH, thumbWidth: THUMB_WIDTH, signal, onProgress });
      const count = Array.isArray(raw?.slides) ? raw.slides.length : 0;
      const hashes = [];
      for (let n = 1; n <= count; n++) {
        const file = path.join(work, `s${pad(n)}.png`);
        if (!(await exists(file)) || !(await exists(path.join(work, `t${pad(n)}.png`)))) {
          throw new ViewerError('render_failed', 'error.slideImageMissing', { n });
        }
        hashes.push(createHash('sha1').update(await fs.readFile(file)).digest('hex').slice(0, 16));
      }
      const manifest = buildManifest(key, source, raw, hashes);
      if (copy) await fs.rm(copy, { force: true });
      await fs.rm(path.join(work, 'slides.json'), { force: true });
      await fs.rm(path.join(work, 'cancel.flag'), { force: true });
      await writeRecord(path.join(work, 'manifest.json'), MANIFEST_FORMAT, MANIFEST_VERSION, manifest);
      const published = await this._publish(work, key, force);
      this._prune(source, key).catch((error) => this.log(this.i18n.t('log.cachePruneFailed', { detail: error.message })));
      return published ?? manifest;
    } finally {
      await fs.rm(work, { recursive: true, force: true }).catch(() => {});
    }
  }

  /** Moves the finished render into place. Returns the manifest already there when another render won the race. */
  async _publish(work, key, force) {
    const finalDir = this.keyDir(key);
    await fs.mkdir(this.decksDir, { recursive: true });
    try {
      await fs.rename(work, finalDir);
      return null;
    } catch (error) {
      if (!(await exists(finalDir))) throw error;
    }
    if (!force) {
      const existing = await this._readCached(key);
      if (existing) return existing;
    }
    const aside = `${finalDir}.old-${randomUUID()}`;
    await fs.rename(finalDir, aside);
    await fs.rename(work, finalDir);
    await fs.rm(aside, { recursive: true, force: true }).catch(() => {});
    return null;
  }

  /** Drops superseded renders of the same file and the oldest decks beyond the keep limit. */
  async _prune(source, currentKey) {
    const now = Date.now();
    const entries = [];
    for (const entry of await fs.readdir(this.decksDir, { withFileTypes: true }).catch(() => [])) {
      if (!entry.isDirectory()) continue;
      const dir = path.join(this.decksDir, entry.name);
      if (!KEY_PATTERN.test(entry.name)) {
        if (/\.old-/.test(entry.name)) await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
        continue;
      }
      const manifestFile = path.join(dir, 'manifest.json');
      const stat = await fs.stat(manifestFile).catch(() => null);
      const record = stat ? await readRecord(manifestFile, { format: MANIFEST_FORMAT, version: MANIFEST_VERSION, validate: validateManifest }).catch(() => null) : null;
      entries.push({ key: entry.name, dir, usedAt: stat?.mtimeMs ?? 0, sourcePath: record?.status === 'ok' ? record.data.source.path : null });
    }
    const doomed = new Set();
    for (const entry of entries) {
      if (entry.key === currentKey || this._pins.has(entry.key)) continue;
      if (entry.sourcePath === null) doomed.add(entry);
      else if (samePath(entry.sourcePath, source.path) && now - entry.usedAt > SUPERSEDED_GRACE_MS) doomed.add(entry);
    }
    const keepers = entries.filter((entry) => !doomed.has(entry)).sort((a, b) => b.usedAt - a.usedAt);
    for (const entry of keepers.slice(KEEP_DECKS)) if (entry.key !== currentKey && !this._pins.has(entry.key)) doomed.add(entry);
    for (const entry of doomed) await fs.rm(entry.dir, { recursive: true, force: true }).catch(() => {});
    for (const entry of await fs.readdir(this.workDir, { withFileTypes: true }).catch(() => [])) {
      const dir = path.join(this.workDir, entry.name);
      const stat = await fs.stat(dir).catch(() => null);
      if (stat && now - stat.mtimeMs > STALE_WORK_MS) await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
