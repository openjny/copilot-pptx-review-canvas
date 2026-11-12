import { createHash, randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createI18n, msg, normalizeLanguage } from './i18n.mjs';

export const CANVAS_ID = 'pptx-slide-viewer';

/** `COPILOT_HOME` exactly as given when it is non-empty, otherwise `~/.copilot`. */
export function resolveCopilotHome(env = process.env, homedir = os.homedir()) {
  const configured = env.COPILOT_HOME;
  return typeof configured === 'string' && configured.length > 0 ? configured : path.join(homedir, '.copilot');
}

/** Private data of this extension; every repository and host on the machine shares it. */
export function dataRootFor(copilotHome = resolveCopilotHome()) {
  return path.join(copilotHome, 'extension-data', CANVAS_ID);
}

/** Paths compare case-insensitively on Windows. */
export function normalizePathKey(filePath) {
  const resolved = path.resolve(filePath);
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

export function samePath(a, b) {
  return normalizePathKey(a) === normalizePathKey(b);
}

export function hashPath(filePath) {
  return createHash('sha1').update(normalizePathKey(filePath)).digest('hex');
}

const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EBUSY', 'EACCES']);

async function renameWithRetry(from, to) {
  for (let attempt = 0; ; attempt++) {
    try {
      await fs.rename(from, to);
      return;
    } catch (error) {
      if (attempt >= 4 || !RETRYABLE_RENAME_CODES.has(error.code)) throw error;
      await new Promise((resolve) => setTimeout(resolve, 25 * (attempt + 1)));
    }
  }
}

/** Writes a uniquely named temporary file next to the target, then renames it over the target. */
export async function writeFileAtomic(target, contents) {
  const dir = path.dirname(target);
  await fs.mkdir(dir, { recursive: true });
  const temporary = path.join(dir, `.${path.basename(target)}.${process.pid}.${randomUUID()}.tmp`);
  try {
    await fs.writeFile(temporary, contents);
    await renameWithRetry(temporary, target);
  } catch (error) {
    await fs.rm(temporary, { force: true }).catch(() => {});
    throw error;
  }
}

export async function writeRecord(file, format, version, data) {
  await writeFileAtomic(file, `${JSON.stringify({ format, version, data }, null, 1)}\n`);
}

/**
 * Reads a `{ format, version, data }` record.
 * Returns `{ status: 'missing' }`, `{ status: 'ok', data }` or `{ status: 'invalid', reason }`.
 * `validate(data)` returns a problem (a `msg()`), or a falsy value when the shape is right. `reason` is such a message too.
 */
export async function readRecord(file, { format, version, validate }) {
  let text;
  try {
    text = await fs.readFile(file, 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') return { status: 'missing' };
    throw error;
  }
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    return { status: 'invalid', reason: msg('record.invalidJson') };
  }
  if (!parsed || typeof parsed !== 'object' || parsed.format !== format) {
    return { status: 'invalid', reason: msg('record.formatMismatch') };
  }
  if (parsed.version !== version) {
    return { status: 'invalid', reason: msg('record.unsupportedVersion', { version: String(parsed.version) }) };
  }
  const problem = validate(parsed.data);
  if (problem) return { status: 'invalid', reason: problem };
  return { status: 'ok', data: parsed.data };
}

/** Keeps an unreadable record for inspection instead of overwriting it. */
export async function quarantine(file) {
  const target = `${file}.corrupt-${Date.now()}`;
  await fs.rename(file, target);
  return target;
}

const POSITION_FORMAT = 'pptx-slide-viewer.position';
const POSITION_VERSION = 1;
const MAX_POSITION_FILES = 100;

function validatePosition(data) {
  if (!data || typeof data !== 'object') return msg('record.notObject');
  if (typeof data.path !== 'string' || data.path.length === 0) return msg('position.pathMissing');
  if (!Number.isInteger(data.slide) || data.slide < 1) return msg('position.slideInvalid');
  if (!Number.isFinite(data.updatedAt)) return msg('position.updatedAtInvalid');
  return null;
}

/** Remembers the slide each deck was left on; one file per deck path, so concurrent sessions do not overwrite each other. */
export class PositionStore {
  constructor(root, { log = () => {}, i18n = createI18n() } = {}) {
    this.dir = path.join(root, 'positions');
    this.log = log;
    this.i18n = i18n;
    this._saves = 0;
    this._writes = new Map();
  }

  fileFor(deckPath) {
    return path.join(this.dir, `${hashPath(deckPath)}.json`);
  }

  async load(deckPath) {
    const file = this.fileFor(deckPath);
    const record = await readRecord(file, { format: POSITION_FORMAT, version: POSITION_VERSION, validate: validatePosition });
    if (record.status === 'ok') return record.data;
    if (record.status === 'invalid') {
      const kept = await quarantine(file).catch(() => null);
      this.log(this.i18n.t(kept ? 'log.positionInvalidKept' : 'log.positionInvalid', { reason: record.reason, kept }));
    }
    return null;
  }

  async save(deckPath, slide) {
    const file = this.fileFor(deckPath);
    const data = { path: path.resolve(deckPath), slide, updatedAt: Date.now() };
    // Saves of one deck run in the order they were asked for, so a slow earlier save cannot overwrite a later one.
    const previous = this._writes.get(file) ?? Promise.resolve();
    const write = previous.catch(() => {}).then(() => writeRecord(file, POSITION_FORMAT, POSITION_VERSION, data));
    this._writes.set(file, write);
    try {
      await write;
    } finally {
      if (this._writes.get(file) === write) this._writes.delete(file);
    }
    if (this._saves++ % 20 === 0) await this._prune().catch(() => {});
  }

  /** The decks most recently left, newest first; files that no longer exist are skipped. */
  async recent(limit = 8) {
    let names;
    try {
      names = (await fs.readdir(this.dir)).filter((name) => name.endsWith('.json'));
    } catch (error) {
      if (error.code === 'ENOENT') return [];
      throw error;
    }
    const entries = [];
    for (const name of names) {
      const record = await readRecord(path.join(this.dir, name), { format: POSITION_FORMAT, version: POSITION_VERSION, validate: validatePosition }).catch(() => null);
      if (record?.status === 'ok') entries.push(record.data);
    }
    entries.sort((a, b) => b.updatedAt - a.updatedAt);
    const result = [];
    for (const entry of entries) {
      if (result.length >= limit) break;
      const exists = await fs.stat(entry.path).then((stat) => stat.isFile(), () => false);
      if (exists) result.push({ path: entry.path, name: path.basename(entry.path), slide: entry.slide, updatedAt: entry.updatedAt });
    }
    return result;
  }

  async _prune() {
    const names = (await fs.readdir(this.dir)).filter((name) => name.endsWith('.json'));
    if (names.length <= MAX_POSITION_FILES) return;
    const stats = await Promise.all(names.map(async (name) => ({ name, mtimeMs: (await fs.stat(path.join(this.dir, name))).mtimeMs })));
    stats.sort((a, b) => b.mtimeMs - a.mtimeMs);
    for (const { name } of stats.slice(MAX_POSITION_FILES)) await fs.rm(path.join(this.dir, name), { force: true });
  }
}

const SETTINGS_FORMAT = 'pptx-slide-viewer.settings';
const SETTINGS_VERSION = 1;

function validateSettings(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) return msg('record.notObject');
  if (normalizeLanguage(data.language) === null) return msg('settings.languageInvalid');
  return null;
}

/** The user's preferences, one small file in the data root that every viewer of every session shares. */
export class SettingsStore {
  constructor(root, { log = () => {}, i18n = createI18n() } = {}) {
    this.file = path.join(root, 'settings.json');
    this.log = log;
    this.i18n = i18n;
    this._write = Promise.resolve();
  }

  /** `{ language }`: the saved language ('en' or 'ja'), or null when nothing usable is saved. An unusable file is kept aside, not overwritten. */
  async load() {
    try {
      const record = await readRecord(this.file, { format: SETTINGS_FORMAT, version: SETTINGS_VERSION, validate: validateSettings });
      if (record.status === 'ok') return { language: normalizeLanguage(record.data.language) };
      if (record.status === 'invalid') {
        const kept = await quarantine(this.file).catch(() => null);
        this.log(this.i18n.t(kept ? 'log.settingsInvalidKept' : 'log.settingsInvalid', { reason: record.reason, kept }));
      }
    } catch (error) {
      this.log(this.i18n.t('log.settingsUnreadable', { detail: error.message }));
    }
    return { language: null };
  }

  /** Saves the language. Saves run in the order they were asked for, so the last choice is the one on disk. */
  save({ language }) {
    const data = { language: normalizeLanguage(language) };
    const write = this._write.catch(() => {}).then(() => writeRecord(this.file, SETTINGS_FORMAT, SETTINGS_VERSION, data));
    this._write = write;
    return write;
  }
}
