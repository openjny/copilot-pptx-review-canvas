import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import http from 'node:http';
import path from 'node:path';
import { validateStroke } from './annotations.mjs';
import { commentKind } from './comments.mjs';
import { pad, resolveDeckSource } from './deck-cache.mjs';
import { ViewerError, isAbortError } from './errors.mjs';
import { UndoJournal } from './history.mjs';
import { SUPPORTED_LANGUAGES, createI18n, normalizeLanguage } from './i18n.mjs';
import { analyzeComments, analyzeSlide } from './marks.mjs';
import { createPresentationWithPowerPoint, suggestNewDeckPath } from './new-deck.mjs';
import { deleteSlideInFile } from './pptx-delete.mjs';
import { commitRewrite, isReorderableExtension, reorderSlidesInFile } from './pptx-reorder.mjs';
import { whitePng } from './png.mjs';
import { compositeAnnotations } from './raster.mjs';
import { SKETCHBOOK_EXTENSION, createSketchbook, resolveSketchbook } from './sketchbook.mjs';
import { CANVAS_ID, THEMES, isTheme, samePath } from './store.mjs';

const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '[::1]']);
const MAX_BODY_BYTES = 16 * 1024;
const MAX_STROKE_BODY_BYTES = 512 * 1024;
const DECK_FILE_PATTERN = /^\/deck\/([0-9a-f]{24})\/([st]\d{4}\.png)$/;
const HEARTBEAT_MS = 20_000;
const WATCH_INTERVAL_MS = 1500;
const POSITION_SAVE_DELAY_MS = 600;
const CONTENT_SECURITY_POLICY = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; base-uri 'none'; form-action 'none'";
const ASSET_TYPES = {
  'index.html': 'text/html; charset=utf-8',
  'styles.css': 'text/css; charset=utf-8',
  'main.js': 'text/javascript; charset=utf-8',
  'ink.js': 'text/javascript; charset=utf-8',
  'drag.js': 'text/javascript; charset=utf-8',
  'icons.js': 'text/javascript; charset=utf-8',
  'comments.js': 'text/javascript; charset=utf-8',
  'resize.js': 'text/javascript; charset=utf-8',
  'review.js': 'text/javascript; charset=utf-8',
  'bubble.js': 'text/javascript; charset=utf-8',
  'place.js': 'text/javascript; charset=utf-8',
  'menu.js': 'text/javascript; charset=utf-8',
  'shape.js': 'text/javascript; charset=utf-8',
  'i18n.js': 'text/javascript; charset=utf-8',
  'locales/en.js': 'text/javascript; charset=utf-8',
  'locales/ja.js': 'text/javascript; charset=utf-8',
};

const CONFLICT_CODES = new Set(['no_deck', 'not_ready', 'busy', 'file_changed', 'file_locked', 'file_exists', 'history_changed', 'history_conflict', 'nothing_to_submit', 'already_submitted', 'agent_unavailable', 'comment_not_editable']);
const UNDO_FILE_MAX_AGE_MS = 24 * 60 * 60 * 1000;

/** An error that answers a request with a given status; it is a ViewerError, so its text is rendered in the current language like any other. */
class HttpError extends ViewerError {
  constructor(status, code, messageKey, params) {
    super(code, messageKey, params);
    this.status = status;
  }
}

function statusFor(error) {
  if (error instanceof HttpError) return error.status;
  if (error instanceof ViewerError) {
    if (error.code === 'file_not_found' || error.code === 'draft_not_found' || error.code === 'comment_not_found') return 404;
    return CONFLICT_CODES.has(error.code) ? 409 : 400;
  }
  return 500;
}

function sendJson(res, status, body) {
  const payload = Buffer.from(JSON.stringify(body));
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': payload.length,
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
  });
  res.end(payload);
}

async function readJsonBody(req, limit = MAX_BODY_BYTES) {
  const type = String(req.headers['content-type'] ?? '').split(';')[0].trim().toLowerCase();
  if (type !== 'application/json') throw new HttpError(415, 'unsupported_media_type', 'http.contentType');
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw new HttpError(413, 'body_too_large', 'http.bodyTooLarge');
    chunks.push(chunk);
  }
  let value;
  try {
    value = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
  } catch {
    throw new HttpError(400, 'invalid_json', 'http.invalidJson');
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new HttpError(400, 'invalid_json', 'http.jsonObject');
  return value;
}

function parseHost(value) {
  try {
    const url = new URL(`http://${value}`);
    return { hostname: url.hostname, port: url.port };
  } catch {
    return null;
  }
}

const sourceStamp = (source) => `${source.size}:${source.mtimeMs}`;

/** A request names either a slide (`slide`) or a draft page (`draft`). */
function targetOf(body) {
  if (body.draft === undefined) return body.slide;
  if (typeof body.draft !== 'string') throw new HttpError(400, 'invalid_input', 'input.draftString');
  return body.draft;
}
const isSketchbook = (deck) => deck?.kind === 'sketchbook';
/** The threads that are still open, and how many are resolved. Resolved ones are done with: the agent does not need to read them again. */
const splitThreads = (comments) => ({ open: comments.filter((comment) => !comment.resolved), resolved: comments.filter((comment) => comment.resolved).length });
const isReorderable = (deck) => Boolean(deck) && !isSketchbook(deck) && isReorderableExtension(path.extname(deck.source.path));
const deckName = (deck) => path.basename(deck.source.path, isSketchbook(deck) ? SKETCHBOOK_EXTENSION : undefined);

/** One slide viewer: an HTTP server on loopback that shows one deck, lets the user mark it up, and keeps its page in sync over server-sent events. */
export class ViewerInstance {
  /**
   * `sendToAgent({ prompt, displayPrompt })` sends a message to the agent (without it a review cannot be submitted);
   * `getPreferences()`, `changeCommentsWidth(width)`, `changeRailHeight(height)` and `changeTheme(theme)` read and save the user's preferences, which the registry owns.
   */
  constructor({ instanceId, deckCache, positions, annotationStore, assetsDir, baseDir, log = () => {}, watchIntervalMs = WATCH_INTERVAL_MS, creator = createPresentationWithPowerPoint, i18n = createI18n(), changeLanguage, sendToAgent, getPreferences, changeCommentsWidth, changeRailHeight, changeTheme }) {
    this.instanceId = instanceId;
    this.deckCache = deckCache;
    this.positions = positions;
    this.annotationStore = annotationStore;
    this.creator = creator;
    this.assetsDir = assetsDir;
    this.baseDir = baseDir;
    this.log = log;
    this.i18n = i18n;
    this.t = (key, params) => this.i18n.t(key, params);
    this.sendToAgent = sendToAgent;
    this.getPreferences = getPreferences ?? (() => ({ commentsWidth: null, railHeight: null, theme: 'auto' }));
    this.changeCommentsWidth = changeCommentsWidth ?? (async (width) => width);
    this.changeRailHeight = changeRailHeight ?? (async (height) => height);
    this.changeTheme = changeTheme ?? (async (theme) => theme);
    // The version of the marks that was last submitted, so the same marks are not submitted twice; any change to them makes a new version.
    this._submittedVersion = null;
    // Saving the language and telling every viewer is the registry's job; on its own, a viewer only switches itself.
    this.changeLanguage = changeLanguage ?? (async (language) => {
      this.i18n.setLanguage(language);
      this.republish();
      return language;
    });
    this.watchIntervalMs = watchIntervalMs;
    this.url = null;
    this._server = null;
    this._port = 0;
    this._assets = new Map();
    this._clients = new Set();
    this._heartbeat = null;
    this._watcher = null;
    this._positionTimer = null;
    this._abort = null;
    this._seq = 0;
    this._rev = 0;
    this._closed = false;
    this._settled = Promise.resolve(null);
    this._book = null;
    this._unsubscribeBook = null;
    this._notice = null;
    this._noticeId = 0;
    this._shapes = { key: null, value: null };
    this._pendingStamp = null;
    this._attemptedStamp = null;
    this._checking = false;
    this._moving = false;
    this._blank = null;
    this.undoDir = path.join(deckCache.root, 'undo');
    this._history = new UndoJournal({ release: (entry) => this._disposeEntry(entry), describe: (entry) => this._labelOf(entry) });
    this._state = { phase: 'empty', deck: null, current: 1, draft: null, progress: null, error: null, pending: null, recent: [] };
  }

  async start() {
    for (const name of Object.keys(ASSET_TYPES)) {
      this._assets.set(name, await fsp.readFile(path.join(this.assetsDir, name)));
    }
    this._server = http.createServer((req, res) => {
      this._handle(req, res).catch((error) => {
        const status = statusFor(error);
        if (status === 500) this.log(this.t('log.requestFailed', { detail: error.stack ?? error.message }));
        if (res.headersSent) return void res.destroy();
        sendJson(res, status, { error: { code: error.code ?? 'internal_error', message: status === 500 ? this.t('http.internal') : this.i18n.messageOf(error) } });
      });
    });
    this._server.on('error', (error) => this.log(this.t('log.serverError', { detail: error.message })));
    await new Promise((resolve, reject) => {
      this._server.once('error', reject);
      this._server.listen(0, '127.0.0.1', () => {
        this._server.off('error', reject);
        resolve();
      });
    });
    this._port = this._server.address().port;
    this.url = `http://127.0.0.1:${this._port}/`;
    this._heartbeat = setInterval(() => this._write(': ping\n\n'), HEARTBEAT_MS);
    this._heartbeat.unref?.();
    if (this.watchIntervalMs > 0) {
      this._watcher = setInterval(() => this._checkFile(), this.watchIntervalMs);
      this._watcher.unref?.();
    }
    this._state.recent = await this.positions.recent().catch(() => []);
    this._purgeUndoDir().catch(() => {});
    return this.url;
  }

  /** The copies kept for undo belong to a viewer that is running; whatever a viewer that has gone left behind is removed after a day. */
  async _purgeUndoDir() {
    const now = Date.now();
    for (const name of await fsp.readdir(this.undoDir).catch(() => [])) {
      const file = path.join(this.undoDir, name);
      const stat = await fsp.stat(file).catch(() => null);
      if (stat && now - stat.mtimeMs > UNDO_FILE_MAX_AGE_MS) await fsp.rm(file, { force: true }).catch(() => {});
    }
  }

  /** Frees what an undo entry holds (the copy of the file as it was, and the pinned render). Safe to call more than once. */
  _disposeEntry(entry) {
    if (entry.pinned) {
      this.deckCache.unpin(entry.beforeKey);
      entry.pinned = false;
    }
    if (entry.undoFile) {
      fsp.rm(entry.undoFile, { force: true }).catch(() => {});
      entry.undoFile = null;
    }
  }

  get title() {
    return this._state.deck ? deckName(this._state.deck) : this._state.pending?.name ?? this.t('app.title');
  }

  get statusText() {
    const { phase, deck, progress, error } = this._state;
    if (phase === 'rendering') return progress?.total ? this.t('status.renderingProgress', { done: progress.done, total: progress.total }) : this.t('status.rendering');
    if (phase === 'error') return this.t('status.error', { message: error ? this.i18n.messageOf(error) : this.t('status.unknownError') });
    if (phase === 'ready') return isSketchbook(deck) ? this.t('status.sketchbook') : this.t(deck.slideCount === 1 ? 'status.slideCountOne' : 'status.slideCount', { count: deck.slideCount });
    return this.t('status.noDeck');
  }

  get hasDeck() {
    return Boolean(this._state.deck);
  }

  _annotationInfo() {
    const book = this._state.deck ? this._book : null;
    if (!book) return { version: '', counts: {}, commentCounts: {}, unread: [], drafts: [] };
    return { version: `${book.deckKey}:${book.rev}`, counts: book.counts(), commentCounts: book.commentCounts(), unread: book.unreadSlides(), drafts: book.drafts() };
  }

  snapshot() {
    const { phase, deck, current, draft, progress, error, pending, recent } = this._state;
    return {
      rev: this._rev,
      language: this.i18n.language,
      commentsWidth: this.getPreferences().commentsWidth ?? null,
      railHeight: this.getPreferences().railHeight ?? null,
      theme: this.getPreferences().theme ?? 'auto',
      phase,
      current,
      draft,
      progress,
      error: error && { code: error.code, message: this.i18n.messageOf(error) },
      pending,
      recent,
      notice: this._notice && { id: this._notice.id, text: this._noticeText(this._notice.parts) },
      annotations: this._annotationInfo(),
      review: this.reviewInfo(),
      history: this._history.view(),
      deck: deck && {
        key: deck.key,
        kind: deck.kind ?? 'pptx',
        reorderable: isReorderable(deck),
        name: deckName(deck),
        path: deck.source.path,
        slideCount: deck.slideCount,
        slideWidth: deck.slideWidth,
        slideHeight: deck.slideHeight,
        renderedAt: deck.renderedAt,
        slides: deck.slides.map(({ n, title, notes }) => ({ n, title: isSketchbook(deck) ? this.t('sketchbook.pageTitle') : title, notes })),
      },
    };
  }

  /** The title of a slide as a person or the agent reads it: a sketchbook page is named in the current language, and a slide without a title is called by its number. */
  _pageTitle(deck, slide) {
    if (isSketchbook(deck)) return this.t('sketchbook.pageTitle');
    return slide.title || this.t('slide.untitled', { n: slide.n });
  }

  /** The notice is kept as messages and rendered here, so it follows a change of language. */
  _noticeText(parts) {
    return parts.map((part) => this.t(part.key, part.params)).reduce((text, next) => this.t('notice.join', { a: text, b: next }));
  }

  _labelOf(entry) {
    return this.t(entry.labelKey, entry.labelParams);
  }

  /** Sends the state again, for when something it is made of changed without the deck changing (the language). */
  republish() {
    this._publish();
  }

  slideInfo(slide = this._state.current) {
    const deck = this._requireDeck();
    this._assertSlide(slide, deck);
    const { notes } = deck.slides[slide - 1];
    return { slide, slideCount: deck.slideCount, title: this._pageTitle(deck, deck.slides[slide - 1]), notes };
  }

  outline({ includeNotes = false } = {}) {
    const deck = this._requireDeck();
    const counts = this._book?.counts() ?? {};
    const notes = this._book?.commentCounts() ?? {};
    const drafts = this._book?.drafts() ?? [];
    return {
      name: deckName(deck),
      path: deck.source.path,
      ...(isSketchbook(deck) ? { kind: 'sketchbook' } : {}),
      slideCount: deck.slideCount,
      current: this._state.current,
      slides: deck.slides.map((meta) => ({
        slide: meta.n,
        title: this._pageTitle(deck, meta),
        ...(counts[meta.n] ? { strokes: counts[meta.n] } : {}),
        ...(notes[meta.n] ? { comments: notes[meta.n] } : {}),
        ...(includeNotes ? { notes: meta.notes } : {}),
      })),
      ...(drafts.length
        ? { draftPages: drafts.map(({ id, title, after, count, comments }) => ({ id, title, insertAfterSlide: after, strokes: count, ...(comments ? { comments } : {}) })) }
        : {}),
    };
  }

  async slideImage(slide = this._state.current) {
    const deck = this._requireDeck();
    this._assertSlide(slide, deck);
    return fsp.readFile(path.join(this.deckCache.keyDir(deck.key), `s${pad(slide)}.png`));
  }

  setCurrent(slide, { persist = true } = {}) {
    const deck = this._requireDeck();
    this._assertSlide(slide, deck);
    if (slide !== this._state.current || this._state.draft !== null) {
      this._state.current = slide;
      this._state.draft = null;
      this._publish();
    }
    if (persist) this._schedulePositionSave();
    return this.slideInfo(slide);
  }

  /** Shows a draft page. It stays where it is in the deck: `current` keeps the slide that was shown before. */
  setDraft(id) {
    if (!this._requireBook().draft(id)) throw new ViewerError('draft_not_found', 'error.draftNotFound');
    if (this._state.draft !== id) {
      this._state.draft = id;
      this._publish();
    }
    return { draft: id };
  }

  /** Validates the file, then starts loading it. `settled` resolves with a summary once the deck is ready or has failed. */
  async openDeck(input, { slide, force = false } = {}) {
    const source = (await resolveSketchbook(this.deckCache.root, input)) ?? (await resolveDeckSource(input, this.baseDir));
    const deck = this._state.deck;
    const unchanged = deck && samePath(deck.source.path, source.path) && deck.source.size === source.size && deck.source.mtimeMs === source.mtimeMs;
    if (!force && unchanged && this._state.phase === 'ready') {
      if (slide) this.setCurrent(Math.min(slide, deck.slideCount));
      return { settled: Promise.resolve(this.summary()) };
    }
    return { settled: this._load(source, { slide, force }) };
  }

  /** Starts (or continues, when the name exists) a sketchbook: pages to sketch ideas on, with no PowerPoint file behind them. */
  async newSketchbook(name) {
    const source = await createSketchbook(this.deckCache.root, name, this.t);
    return this.openDeck(source.path);
  }

  /** A free file name for a new presentation, in the folder the session works in. */
  suggestNewDeckPath() {
    return suggestNewDeckPath(this.baseDir, this.t);
  }

  /** Creates a presentation with one slide (blank, or from `template`) and opens it. Never overwrites a file. */
  async newDeck(input, { template } = {}) {
    const file = await this.creator({ path: input, template, baseDir: this.baseDir });
    return this.openDeck(file);
  }

  /**
   * Moves a slide to a new position in the file itself (only the order is rewritten), then shows the result at once: the
   * pictures are reused rather than drawn again, and the marks and draft pages follow their slides.
   */
  async moveSlide(from, to) {
    const deck = this._requireDeck();
    this._requireBook();
    this._assertRewritable(deck, 'error.sketchbookNoReorder', 'error.legacyNoReorder');
    this._assertSlide(from, deck);
    this._assertSlide(to, deck);
    if (from === to) return { slide: to, changed: false };
    await this._rewriteDeck(deck, (file, backupDir) => reorderSlidesInFile(file, from, to, { backupDir }), () => to);
    // The file is no longer what an earlier deletion left, so that deletion can not be undone any more.
    this._history.clear();
    this._publish();
    return { slide: to, changed: true };
  }

  /**
   * Deletes a slide from the file itself (the file as it was is kept as a backup), then shows the result at once. The slide
   * that follows takes its place on screen. Its marks are set aside in the archive; draft pages that followed it follow the slide before.
   * The deletion can be undone (see `historyStep`). With `entry`, it is the redo of that undone deletion.
   */
  async deleteSlide(slide, { entry = null } = {}) {
    const deck = this._requireDeck();
    const book = this._requireBook();
    this._assertRewritable(deck, 'error.sketchbookNoDelete', 'error.legacyNoDelete');
    this._assertSlide(slide, deck);
    if (deck.slideCount === 1) throw new ViewerError('last_slide', 'error.lastSlide');

    const saved = { strokes: [...book.strokes(slide)], comments: [...book.comments(slide)], unread: book.isUnread(slide), draftIds: book.drafts().filter((draft) => draft.after === slide).map((draft) => draft.id) };
    let undoFile = path.join(this.undoDir, `${randomUUID()}.pptx`);
    const { removedParts, next } = await this._rewriteDeck(
      deck,
      async (file, backupDir) => {
        // The copy is made before the file changes, so nothing else slows down the moment between the write and showing the result.
        try {
          await fsp.mkdir(this.undoDir, { recursive: true });
          await fsp.copyFile(file, undoFile);
        } catch (error) {
          this.log(this.t('log.undoCopyFailed', { detail: error.message }));
          undoFile = null;
        }
        try {
          return await deleteSlideInFile(file, slide, { backupDir });
        } catch (error) {
          if (undoFile) fsp.rm(undoFile, { force: true }).catch(() => {});
          throw error;
        }
      },
      (shown) => Math.min(slide, shown.slideCount),
      { dropping: true },
    );
    if (undoFile) {
      this.deckCache.pin(deck.key);
      const data = {
        type: 'slide',
        labelKey: 'history.deleteSlide',
        labelParams: { slide },
        slide,
        count: deck.slideCount,
        beforeKey: deck.key,
        pinned: true,
        undoFile,
        after: { size: next.source.size, mtimeMs: next.source.mtimeMs },
        saved,
      };
      if (entry) Object.assign(entry, data);
      else this._history.record(data);
    } else if (!entry) {
      this._history.clear();
    }
    this._publish();
    return { slide: Math.min(slide, deck.slideCount - 1), slideCount: deck.slideCount - 1, changed: true, removedParts };
  }

  /**
   * Undoes or redoes the action the page asked for. `id` must be the action `kind` ('undo' or 'redo') would act on now; the page
   * got it from the state it showed, and a different one means the history moved on in the meantime.
   */
  async historyStep(kind, id) {
    if (kind !== 'undo' && kind !== 'redo') throw new ViewerError('invalid_input', 'error.historyKind');
    const entry = this._history.peek(kind);
    if (!entry || entry.id !== id) throw new ViewerError('history_changed', 'error.historyChanged');
    if (entry.type === 'draft') this._stepDraft(kind, entry);
    else if (entry.type === 'comment') this._stepComment(kind, entry);
    else if (kind === 'undo') await this._undoSlideDeletion(entry);
    else await this.deleteSlide(entry.slide, { entry });
    if (this._history.peek(kind)?.id === entry.id) this._history.commit(kind);
    this._publish();
    return { ok: true, kind: entry.type, label: this._labelOf(entry) };
  }

  /** Brings back a deleted comment (undo) or deletes it again (redo), and shows the page it belongs to. */
  _stepComment(kind, entry) {
    const book = this._requireBook();
    this._assertTarget(entry.target);
    if (kind === 'undo') {
      book.reinsertComment(entry.target, entry.comment, entry.index);
    } else {
      const removed = book.removeComment(entry.target, entry.comment.id);
      if (!removed) throw new ViewerError('comment_not_found', 'error.commentNotFound');
      entry.comment = removed.comment;
      entry.index = removed.index;
    }
    if (typeof entry.target === 'string') this._state.draft = entry.target;
    else {
      this._state.draft = null;
      this._state.current = entry.target;
    }
    this._afterAnnotationChange(book);
  }

  /** Brings back a deleted new slide (undo) or deletes it again (redo). */
  _stepDraft(kind, entry) {
    const book = this._requireBook();
    if (kind === 'undo') {
      book.reinsertDraft(entry.draft, entry.index, entry.unread);
      this._state.draft = entry.draft.id;
      this._saveBook(book);
      return;
    }
    entry.index = book.draftIndex(entry.draft.id);
    entry.unread = book.isUnread(entry.draft.id);
    const draft = this._dropDraft(book, entry.draft.id);
    if (!draft) throw new ViewerError('draft_not_found', 'error.draftNotFound');
    entry.draft = draft;
    this._archiveDrafts(book, [draft]);
  }

  /**
   * Puts the file back as it was before the slide was deleted (from the copy kept for this), then brings the slide, its marks and the
   * draft pages that followed it back. Only possible while the file is still the one the deletion wrote.
   */
  async _undoSlideDeletion(entry) {
    const deck = this._requireDeck();
    this._requireBook();
    this._assertRewritable(deck, 'error.sketchbookNoUndo', 'error.legacyNoUndo');
    if (this._moving) throw new ViewerError('busy', 'error.rewriteBusy');
    this._moving = true;
    const seq = this._seq;
    try {
      const file = deck.source.path;
      const onDisk = await resolveDeckSource(file);
      const unchanged = (stamp) => stamp.size === entry.after.size && stamp.mtimeMs === entry.after.mtimeMs;
      if (!unchanged(onDisk) || !unchanged(deck.source) || deck.slideCount !== entry.count - 1) {
        this._history.clear();
        throw new ViewerError('history_conflict', 'error.historyConflictChanged');
      }
      const restored = await fsp.readFile(entry.undoFile).catch(() => null);
      if (!restored) {
        this._history.clear();
        throw new ViewerError('history_conflict', 'error.historyConflictNoCopy');
      }
      await commitRewrite(file, await fsp.readFile(file), restored, path.join(this.deckCache.root, 'backups'));
      const source = await resolveDeckSource(file);
      const order = Array.from({ length: entry.count }, (_, i) => i + 1);
      // The pictures of the file as it was are still in the cache; if they are not, the file is drawn again.
      const next = await this.deckCache.derive(entry.beforeKey, source, order).catch((error) => {
        if (error.code !== 'render_failed') throw error;
        return this.deckCache.resolve(source);
      });
      await this.annotationStore.attach(next, { insert: { slide: entry.slide, strokes: entry.saved.strokes, comments: entry.saved.comments, unread: entry.saved.unread, draftIds: entry.saved.draftIds } });
      if (seq === this._seq) {
        this._state = { ...this._state, deck: next, current: Math.min(entry.slide, next.slideCount), draft: null };
        this._reconcileDraft();
        this._publish();
        this._schedulePositionSave();
      }
      this._disposeEntry(entry);
    } finally {
      this._moving = false;
    }
  }

  _assertRewritable(deck, sketchbookMessage, legacyMessage) {
    if (isReorderable(deck)) return;
    throw new ViewerError('not_reorderable', isSketchbook(deck) ? sketchbookMessage : legacyMessage);
  }

  /**
   * Rewrites the slides of the open file with `rewrite(file, backupDir)` (which returns the new `order`), then publishes the
   * result without drawing again: the pictures are reused, and the marks and draft pages follow their slides.
   */
  async _rewriteDeck(deck, rewrite, showSlide, { dropping = false } = {}) {
    if (this._moving) throw new ViewerError('busy', 'error.rewriteBusy');
    this._moving = true;
    const seq = this._seq;
    try {
      const onDisk = await resolveDeckSource(deck.source.path);
      if (onDisk.size !== deck.source.size || onDisk.mtimeMs !== deck.source.mtimeMs) {
        throw new ViewerError('file_changed', 'error.fileChanged');
      }
      const result = await rewrite(deck.source.path, path.join(this.deckCache.root, 'backups'));
      const next = await this.deckCache.derive(deck.key, await resolveDeckSource(deck.source.path), result.order, { dropping });
      await this.annotationStore.attach(next, { order: result.order });
      // Something else started loading a deck while the file was being rewritten: that load shows the result.
      if (seq === this._seq) {
        this._state = { ...this._state, deck: next, current: showSlide(next) };
        this._reconcileDraft();
        this._publish();
        this._schedulePositionSave();
      }
      return { ...result, next };
    } finally {
      this._moving = false;
    }
  }

  async refresh() {
    const deck = this._requireDeck();
    if (isSketchbook(deck)) throw new ViewerError('not_supported', 'error.sketchbookNoRefresh');
    const source = await resolveDeckSource(deck.source.path);
    return { settled: this._load(source, { slide: this._state.current, force: true }) };
  }

  /** Resolves with the viewer's summary once no load is running (or with `timedOut` after `timeoutMs`). */
  whenSettled(timeoutMs) {
    let timer;
    const timeout = new Promise((resolve) => {
      timer = setTimeout(() => resolve({ ...this.summary(), timedOut: true }), timeoutMs);
    });
    return Promise.race([this._awaitSettled(), timeout]).finally(() => clearTimeout(timer));
  }

  /** Waits for a load that is still running, then returns the viewer. Throws when the load failed with nothing to show or did not finish within `timeoutMs`. */
  async whenReady(timeoutMs) {
    const result = await this.whenSettled(timeoutMs);
    if (result.timedOut) throw new ViewerError('still_rendering', 'error.stillRendering');
    if (result.phase === 'error' && !this.hasDeck && this._state.error) throw new ViewerError(this._state.error.code, this._state.error.messageKey, this._state.error.params);
    return this;
  }

  async _awaitSettled() {
    for (;;) {
      const current = this._settled;
      await current;
      if (current === this._settled) return this.summary();
    }
  }

  summary() {
    const { phase, deck, current, error, progress } = this._state;
    return { phase, name: deck ? deckName(deck) : null, slideCount: deck?.slideCount ?? 0, current, error: error && { code: error.code, message: this.i18n.messageOf(error) }, progress };
  }

  // ---- annotations (pen and highlighter strokes) -------------------------------------------------------------

  _requireBook() {
    if (!this._state.deck || !this._book || this._state.phase !== 'ready') throw new ViewerError('not_ready', 'error.notReadyToDraw');
    return this._book;
  }

  _saveBook(book) {
    this.annotationStore.save(book).catch((error) => this.log(this.t('log.annotationsSaveFailed', { detail: error.message })));
  }

  /** Follows the book of the open deck: every viewer of that deck shares it, and each republishes its state when it changes. */
  _useBook(book) {
    if (book === this._book) return;
    this._unsubscribeBook?.();
    this._book = book;
    this._unsubscribeBook = book
      ? book.subscribe(() => {
        this._reconcileDraft();
        this._publish();
      })
      : null;
  }

  /** The page on screen may be a draft that was removed elsewhere (another viewer, or the agent); fall back to the slide. */
  _reconcileDraft() {
    const { draft } = this._state;
    if (draft && !this._book?.draft(draft)) this._state.draft = null;
  }

  _afterAnnotationChange(book) {
    this._saveBook(book);
  }

  /** A target is a slide number or a draft id; this checks it exists. */
  _assertTarget(target) {
    if (typeof target === 'string') {
      if (!this._book?.draft(target)) throw new ViewerError('draft_not_found', 'error.draftNotFound');
      return;
    }
    this._assertSlide(target, this._state.deck);
  }

  annotationsFor(target) {
    const book = this._requireBook();
    this._assertTarget(target);
    return { version: this._annotationInfo().version, ...(typeof target === 'string' ? { draft: target } : { slide: target }), strokes: book.strokes(target) };
  }

  addStroke(target, input) {
    const book = this._requireBook();
    this._assertTarget(target);
    const stroke = validateStroke(input);
    if (book.add(target, stroke)) this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, count: book.strokes(target).length };
  }

  removeStrokes(target, ids) {
    const book = this._requireBook();
    this._assertTarget(target);
    if (!Array.isArray(ids) || ids.length > 1000 || ids.some((id) => typeof id !== 'string')) throw new ViewerError('invalid_input', 'input.idsStrings');
    const removed = book.remove(target, ids);
    if (removed) this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, count: book.strokes(target).length, removed };
  }

  clearAnnotations(target) {
    const book = this._requireBook();
    this._assertTarget(target);
    const removed = book.clear(target);
    if (removed) this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, count: 0, removed };
  }

  // ---- comments about a page ------------------------------------------------------------------------------------

  commentsFor(target) {
    const book = this._requireBook();
    this._assertTarget(target);
    return { version: this._annotationInfo().version, ...(typeof target === 'string' ? { draft: target } : { slide: target }), comments: book.comments(target) };
  }

  /**
   * Adds a comment about the whole page (no `anchor`), or about a point or an area of it (`anchor`: `{ x, y, w, h }`, fractions of the page).
   * This is the user writing it in the page, so it is always the user's.
   */
  addComment(target, { text, anchor } = {}) {
    const book = this._requireBook();
    this._assertTarget(target);
    const comment = book.addComment(target, { text, anchor, author: 'user' });
    this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, comment };
  }

  updateComment(target, id, text) {
    const book = this._requireBook();
    this._assertTarget(target);
    if (typeof id !== 'string') throw new ViewerError('invalid_input', 'input.idString');
    const updated = book.updateComment(target, id, text);
    if (updated) this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, comment: updated ?? book.comments(target).find((comment) => comment.id === id) };
  }

  /** The user replying in a thread. */
  replyComment(target, id, text) {
    const book = this._requireBook();
    this._assertTarget(target);
    if (typeof id !== 'string') throw new ViewerError('invalid_input', 'input.idString');
    const { comment, reply } = book.replyComment(target, id, { text, author: 'user' });
    this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, comment, reply };
  }

  /** The user resolving a thread, or opening it again. */
  resolveComment(target, id, resolved) {
    const book = this._requireBook();
    this._assertTarget(target);
    if (typeof id !== 'string') throw new ViewerError('invalid_input', 'input.idString');
    if (typeof resolved !== 'boolean') throw new ViewerError('invalid_input', 'input.resolvedBoolean');
    const updated = book.setResolved(target, id, resolved);
    if (updated) this._afterAnnotationChange(book);
    return { version: this._annotationInfo().version, comment: updated ?? book.comments(target).find((comment) => comment.id === id) };
  }

  /**
   * The agent starting a thread on a slide (`slide`) or a draft page (`draft`): about the whole page, or at `anchor` (`{ x, y, w, h }`
   * in percent of the page), or around the shape called `shape` (a name from `onShapes`). It is not "unread": the agent wrote it.
   */
  async addAgentComment({ slide, draft, text, anchor, shape } = {}) {
    const book = this._requireBook();
    const deck = this._requireDeck();
    if ((slide === undefined) === (draft === undefined)) throw new ViewerError('invalid_input', 'input.slideOrDraft');
    const target = draft ?? slide;
    this._assertTarget(target);
    if (anchor !== undefined && shape !== undefined) throw new ViewerError('invalid_input', 'input.anchorOrShape');
    let place = null;
    if (shape !== undefined) {
      if (typeof shape !== 'string' || !shape) throw new ViewerError('invalid_input', 'input.shapeString');
      if (draft !== undefined) throw new ViewerError('invalid_input', 'input.shapeOnSlide');
      const found = ((await this._shapesFor(deck.key))?.[String(slide)] ?? []).find((entry) => entry.name === shape);
      if (!found) throw new ViewerError('invalid_input', 'error.shapeNotFound', { shape });
      place = { x: found.x / deck.slideWidth, y: found.y / deck.slideHeight, w: found.w / deck.slideWidth, h: found.h / deck.slideHeight };
    } else if (anchor !== undefined && anchor !== null) {
      if (typeof anchor !== 'object' || Array.isArray(anchor) || ![anchor.x, anchor.y].every(Number.isFinite)) throw new ViewerError('invalid_input', 'input.anchorPercent');
      place = { x: anchor.x / 100, y: anchor.y / 100, w: (anchor.w ?? 0) / 100, h: (anchor.h ?? 0) / 100 };
    }
    const comment = book.addComment(target, { text, anchor: place, author: 'agent' });
    this._afterAnnotationChange(book);
    return { id: comment.id, ...(typeof target === 'string' ? { draft: target } : { slide: target }), kind: commentKind(comment) };
  }

  /** The agent answering in a thread (`id` is the comment's id from get_annotations). */
  replyAsAgent({ id, text } = {}) {
    const book = this._requireBook();
    this._requireDeck();
    if (typeof id !== 'string') throw new ViewerError('invalid_input', 'input.idString');
    const found = book.findComment(id);
    if (!found) throw new ViewerError('comment_not_found', 'error.commentNotFound');
    const { reply } = book.replyComment(found.target, id, { text, author: 'agent' });
    this._afterAnnotationChange(book);
    return { id: reply.id, comment: id, ...(typeof found.target === 'string' ? { draft: found.target } : { slide: found.target }), replies: book.findComment(id).comment.replies.length };
  }

  /** The user deleting a comment. It can be undone (see `historyStep`). */
  removeComment(target, id) {
    const book = this._requireBook();
    this._assertTarget(target);
    if (typeof id !== 'string') throw new ViewerError('invalid_input', 'input.idString');
    const removed = book.removeComment(target, id);
    if (!removed) throw new ViewerError('comment_not_found', 'error.commentNotFound');
    this._history.record({ type: 'comment', labelKey: 'history.deleteComment', target, comment: removed.comment, index: removed.index });
    this._afterAnnotationChange(book);
    this._publish();
    return { version: this._annotationInfo().version };
  }

  // ---- draft pages (ideas for new slides) ---------------------------------------------------------------------

  /** Adds a white page right after the page on screen (or after `after` / `afterDraft`) and shows it. */
  addDraft({ after, afterDraft } = {}) {
    const book = this._requireBook();
    const where = afterDraft !== undefined || after !== undefined ? { after, afterDraft } : this._state.draft ? { afterDraft: this._state.draft } : { after: this._state.current };
    const draft = book.addDraft(where);
    this._state.draft = draft.id;
    this._publish();
    this._saveBook(book);
    return { version: this._annotationInfo().version, draft };
  }

  _dropDraft(book, id) {
    const info = book.drafts().find((draft) => draft.id === id);
    if (!info) return null;
    if (this._state.draft === id) {
      this._state.draft = null;
      this._state.current = Math.min(Math.max(1, info.after), this._state.deck.slideCount);
    }
    return book.removeDraft(id);
  }

  async _archiveDrafts(book, drafts) {
    await this.annotationStore.archiveDrafts(book, drafts).catch((error) => this.log(this.t('log.draftArchiveFailed', { detail: error.message })));
    this._saveBook(book);
  }

  /** Moves a draft page to follow slide `after` or the draft `afterDraft`; the viewer keeps showing it. */
  moveDraft(id, { after, afterDraft } = {}) {
    const book = this._requireBook();
    const result = book.moveDraft(id, { after, afterDraft });
    if (result.changed) this._saveBook(book);
    return { version: this._annotationInfo().version, ...result };
  }

  /** The user deleting a draft page. It can be undone (see `historyStep`). */
  removeDraft(id) {
    const book = this._requireBook();
    const index = book.draftIndex(id);
    const unread = book.isUnread(id);
    const draft = this._dropDraft(book, id);
    if (!draft) throw new ViewerError('draft_not_found', 'error.draftNotFound');
    this._history.record({ type: 'draft', labelKey: 'history.deleteNewSlide', draft, index, unread });
    this._publish();
    this._archiveDrafts(book, [draft]);
    return { version: this._annotationInfo().version, removed: draft.strokes.length };
  }

  /** The agent saying it has turned these drafts into real slides: they are taken off the page and set aside. */
  async resolveDrafts(ids) {
    const book = this._requireBook();
    if (!Array.isArray(ids) || ids.some((id) => typeof id !== 'string')) throw new ViewerError('invalid_input', 'input.idsStrings');
    const resolved = [];
    const missing = [];
    for (const id of ids) {
      const draft = this._dropDraft(book, id);
      if (draft) resolved.push(draft);
      else missing.push(id);
    }
    await this._archiveDrafts(book, resolved);
    return { resolved: resolved.map((draft) => ({ id: draft.id, title: draft.title })), missing };
  }

  /** The slides and draft pages whose marks the agent has not looked at yet, or null when there is nothing new. */
  unreadAnnotations() {
    if (!this._state.deck || !this._book) return null;
    const slides = this._book.unreadSlides();
    const drafts = this._book.unreadDrafts().map((draft) => draft.title);
    if (!slides.length && !drafts.length) return null;
    return { deck: deckName(this._state.deck), path: this._state.deck.source.path, slides, drafts };
  }

  /**
   * What the page needs for the submit button: whether there is something the agent has not read (`slides`, `drafts`), and whether it
   * has been submitted already (`submitted`) or can be (`canSubmit`). Reading the marks, or changing them, ends "submitted".
   */
  reviewInfo() {
    const unread = this.unreadAnnotations();
    if (!unread) return { canSubmit: false, submitted: false, slides: [], drafts: 0 };
    const submitted = this._submittedVersion === this._annotationInfo().version;
    return { canSubmit: Boolean(this.sendToAgent) && !submitted, submitted, slides: unread.slides, drafts: unread.drafts.length };
  }

  /** The message that asks the agent to act on the review: what it says to the agent, and the shorter text the chat shows in its place. */
  _reviewMessage(unread) {
    const scope = [];
    if (unread.slides.length) scope.push(this.t('unread.slides', { list: unread.slides }));
    if (unread.drafts.length) scope.push(this.t('unread.drafts', { count: unread.drafts.length }));
    const params = { deck: unread.deck, canvasId: CANVAS_ID, instanceId: this.instanceId, scope: scope.join(' / ') };
    return {
      prompt: [this.t('submit.prompt', params), this.t('submit.outro')].join('\n\n'),
      displayPrompt: this.t('submit.display', params),
    };
  }

  /**
   * The user finishing the review: asks the agent, in the session this viewer belongs to, to read the marks, comments and new slides
   * and to make the changes. The message is queued behind whatever the agent is doing. Only marks the agent has not read can be
   * submitted, and the same marks only once.
   */
  async submitReview() {
    if (!this.sendToAgent) throw new ViewerError('agent_unavailable', 'error.agentUnavailable');
    this._requireBook();
    const unread = this.unreadAnnotations();
    if (!unread) throw new ViewerError('nothing_to_submit', 'error.nothingToSubmit');
    const version = this._annotationInfo().version;
    if (this._submittedVersion === version) throw new ViewerError('already_submitted', 'error.alreadySubmitted');
    try {
      await this.sendToAgent(this._reviewMessage(unread));
    } catch (error) {
      throw new ViewerError('agent_unavailable', 'error.agentSendFailed', { detail: error?.message ?? String(error) });
    }
    this._submittedVersion = version;
    this._publish();
    return { submitted: true };
  }

  async _shapesFor(key) {
    if (this._shapes.key !== key) this._shapes = { key, value: await this.deckCache.readShapes(key) };
    return this._shapes.value;
  }

  _blankPage(deck) {
    const width = 1600;
    const height = Math.round((width * deck.slideHeight) / deck.slideWidth);
    if (this._blank?.key !== `${width}x${height}`) this._blank = { key: `${width}x${height}`, png: whitePng(width, height) };
    return this._blank.png;
  }

  /**
   * What the agent needs to act on the marks: per slide, the marks (where, what kind, which shapes lie under them) and the slide
   * image with the strokes and a number next to each mark; and the draft pages (new slide ideas), with the slide each follows.
   * Reading it clears the unread flag.
   */
  async annotationReport({ slide, draft, includeImages = true, includeClean = false, maxImages = 8 } = {}) {
    const deck = this._requireDeck();
    const book = this._book;
    if (!book) throw new ViewerError('not_ready', 'error.notReadyToRun');
    if (slide !== undefined) this._assertSlide(slide, deck);
    if (draft !== undefined) this._assertTarget(draft);
    const slideTargets = draft !== undefined ? [] : slide !== undefined ? [slide] : book.annotatedSlides();
    const draftTargets = slide !== undefined && draft === undefined ? [] : book.drafts().filter((entry) => draft === undefined || entry.id === draft);
    const shapes = await this._shapesFor(deck.key);
    const aspect = deck.slideWidth / deck.slideHeight;
    const slides = [];
    const draftPages = [];
    const images = [];
    for (const n of slideTargets) {
      const strokes = book.strokes(n);
      const meta = deck.slides[n - 1];
      const slideShapes = shapes?.[String(n)];
      const analysis = analyzeSlide({ strokes, aspect, shapes: slideShapes, slideW: deck.slideWidth, slideH: deck.slideHeight, t: this.t });
      const { open, resolved } = splitThreads(book.comments(n));
      const notes = analyzeComments({ comments: open, aspect, shapes: slideShapes, slideW: deck.slideWidth, slideH: deck.slideHeight, t: this.t });
      slides.push({
        slide: n,
        title: this._pageTitle(deck, meta),
        strokeCount: strokes.length,
        marks: analysis.marks,
        ...(notes.comments.length ? { commentCount: notes.comments.length, comments: notes.comments } : {}),
        ...(resolved ? { resolvedCommentCount: resolved } : {}),
        ...(meta.notes ? { notes: meta.notes } : {}),
      });
      if (includeImages && (strokes.length || notes.badges.length) && images.length < maxImages) {
        const png = await this.slideImage(n);
        if (includeClean) images.push({ slide: n, kind: 'clean', data: png });
        images.push({ slide: n, kind: 'annotated', data: compositeAnnotations(png, strokes, analysis.badges, notes) });
      }
    }
    for (const entry of draftTargets) {
      const strokes = book.strokes(entry.id);
      const analysis = analyzeSlide({ strokes, aspect, shapes: null, slideW: deck.slideWidth, slideH: deck.slideHeight, t: this.t });
      const { open, resolved } = splitThreads(book.comments(entry.id));
      const notes = analyzeComments({ comments: open, aspect, shapes: null, slideW: deck.slideWidth, slideH: deck.slideHeight, t: this.t });
      const followed = entry.after > 0 ? deck.slides[entry.after - 1] : null;
      draftPages.push({
        id: entry.id,
        title: entry.title,
        insertAfterSlide: entry.after,
        ...(followed ? { afterSlideTitle: this._pageTitle(deck, followed) } : {}),
        strokeCount: strokes.length,
        marks: analysis.marks,
        ...(notes.comments.length ? { commentCount: notes.comments.length, comments: notes.comments } : {}),
        ...(resolved ? { resolvedCommentCount: resolved } : {}),
      });
      if (includeImages && (strokes.length || notes.badges.length) && images.length < maxImages) {
        images.push({ draft: entry.id, title: entry.title, after: entry.after, kind: 'annotated', data: compositeAnnotations(this._blankPage(deck), strokes, analysis.badges, notes) });
      }
    }
    if (book.markRead([...slideTargets, ...draftTargets.map((entry) => entry.id)])) this._saveBook(book);
    const hasImage = (entry) => entry.strokeCount > 0 || (entry.comments ?? []).some((comment) => comment.kind !== 'page');
    const withImages = slides.filter(hasImage).length + draftPages.filter(hasImage).length;
    const shown = images.filter((image) => image.kind === 'annotated').length;
    return {
      info: {
        deck: deckName(deck),
        path: deck.source.path,
        ...(isSketchbook(deck) ? { kind: 'sketchbook' } : {}),
        slideCount: deck.slideCount,
        currentSlide: this._state.current,
        annotatedSlides: slides.filter((entry) => entry.strokeCount > 0 || entry.commentCount > 0).length,
        slides,
        draftPages,
        shapeNamesAvailable: Boolean(shapes),
        ...(includeImages && withImages > shown ? { imagesOmitted: this.t('report.imagesOmitted') } : {}),
      },
      images,
    };
  }

  // ---- lifecycle ----------------------------------------------------------------------------------------------

  async close() {
    this._closed = true;
    clearInterval(this._heartbeat);
    clearInterval(this._watcher);
    clearTimeout(this._positionTimer);
    this._seq += 1;
    this._abort?.abort();
    this._history.clear();
    this._useBook(null);
    await this._flushPosition().catch(() => {});
    await this.annotationStore.flush().catch(() => {});
    for (const client of this._clients) client.end();
    this._clients.clear();
    if (this._server) {
      await new Promise((resolve) => {
        this._server.close(() => resolve());
        this._server.closeAllConnections?.();
      });
    }
  }

  _requireDeck() {
    if (!this._state.deck) throw new ViewerError('no_deck', 'error.noDeck');
    return this._state.deck;
  }

  _assertSlide(slide, deck) {
    if (!Number.isInteger(slide) || slide < 1 || slide > deck.slideCount) {
      throw new ViewerError('slide_out_of_range', 'error.slideOutOfRange', { count: deck.slideCount });
    }
  }

  /** Reloads the deck by itself when the file on disk changes, once it has stopped changing (a build writes the file over a moment). */
  async _checkFile() {
    const deck = this._state.deck;
    if (this._checking || this._moving || this._closed || !deck || isSketchbook(deck) || this._state.phase === 'rendering') return;
    this._checking = true;
    try {
      const source = await resolveDeckSource(deck.source.path);
      const stamp = sourceStamp(source);
      if (stamp === sourceStamp(deck.source)) {
        this._pendingStamp = null;
        return;
      }
      if (stamp === this._attemptedStamp) return;
      if (stamp !== this._pendingStamp) {
        this._pendingStamp = stamp;
        return;
      }
      this._attemptedStamp = stamp;
      this._pendingStamp = null;
      this._load(source, { slide: this._state.current, force: false, auto: true });
    } catch {
      // The file may be missing for a moment while it is rewritten; try again at the next tick.
    } finally {
      this._checking = false;
    }
  }

  _load(source, { slide, force, auto = false }) {
    this._abort?.abort();
    this._history.clear();
    const controller = new AbortController();
    this._abort = controller;
    const seq = ++this._seq;
    const sameDeck = this._state.deck && samePath(this._state.deck.source.path, source.path);
    if (!sameDeck) this._useBook(null);
    this._state = {
      ...this._state,
      phase: 'rendering',
      error: null,
      progress: { done: 0, total: 0 },
      pending: { name: source.name },
      deck: sameDeck ? this._state.deck : null,
      draft: sameDeck ? this._state.draft : null,
    };
    this._publish();
    this._settled = this._run(source, { slide, force, seq, controller, auto });
    return this._settled;
  }

  _noticeFor(previousKey, deck, attached, auto) {
    const rebuilt = previousKey ? previousKey !== deck.key : attached.changed;
    if (!rebuilt) return null;
    const parts = [];
    if (previousKey) parts.push({ key: auto ? 'notice.reloadedAuto' : 'notice.reloaded' });
    if (attached.kept > 0) parts.push({ key: 'notice.kept', params: { count: attached.kept } });
    if (attached.archived > 0) parts.push({ key: 'notice.archived', params: { count: attached.archived } });
    return parts.length ? parts : null;
  }

  async _run(source, { slide, force, seq, controller, auto }) {
    try {
      const previousKey = this._state.deck?.key ?? null;
      const deck = await this.deckCache.resolve(source, {
        force,
        signal: controller.signal,
        onProgress: (done, total) => {
          if (seq !== this._seq) return;
          this._state.progress = { done, total };
          this._send('progress', { done, total });
        },
      });
      if (seq !== this._seq) return this.summary();
      let start = slide;
      if (!start) {
        const saved = await this.positions.load(source.path).catch((error) => {
          this.log(this.t('log.positionReadFailed', { detail: error.message }));
          return null;
        });
        start = saved?.slide ?? 1;
      }
      if (seq !== this._seq) return this.summary();
      const attached = await this.annotationStore.attach(deck);
      if (seq !== this._seq) return this.summary();
      this._useBook(attached.book);
      this._reconcileDraft();
      const parts = this._noticeFor(previousKey, deck, attached, auto);
      if (parts) this._notice = { id: ++this._noticeId, parts };
      const current = Math.min(Math.max(1, start), deck.slideCount);
      this._state = { ...this._state, phase: 'ready', deck, current, progress: null, error: null, pending: null };
      this._publish();
      this.positions.save(source.path, current).catch((error) => this.log(this.t('log.positionSaveFailed', { detail: error.message }))).finally(() => this._refreshRecent());
    } catch (error) {
      if (seq !== this._seq || isAbortError(error)) return this.summary();
      const known = error instanceof ViewerError;
      if (!known) this.log(this.t('log.renderUnexpected', { detail: error.stack ?? error.message }));
      this._state = {
        ...this._state,
        phase: 'error',
        error: known ? error : new ViewerError('internal_error', 'error.unexpected', { detail: error.message }),
        progress: null,
        pending: null,
      };
      this._publish();
    }
    return this.summary();
  }

  _refreshRecent() {
    this.positions.recent().then((recent) => {
      this._state.recent = recent;
      this._publish();
    }, () => {});
  }

  _schedulePositionSave() {
    clearTimeout(this._positionTimer);
    this._positionTimer = setTimeout(() => this._flushPosition().catch((error) => this.log(this.t('log.positionSaveFailed', { detail: error.message }))), POSITION_SAVE_DELAY_MS);
    this._positionTimer.unref?.();
  }

  async _flushPosition() {
    clearTimeout(this._positionTimer);
    const { deck, current } = this._state;
    if (deck) await this.positions.save(deck.source.path, current);
  }

  _write(frame) {
    for (const client of this._clients) {
      try {
        client.write(frame);
      } catch {
        this._clients.delete(client);
      }
    }
  }

  _send(event, data) {
    this._write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  }

  _publish() {
    this._rev += 1;
    this._send('state', this.snapshot());
  }

  /** Only loopback hosts and the page's own origin may talk to the server (blocks DNS rebinding and cross-site requests). */
  _guard(req) {
    const host = parseHost(req.headers.host ?? '');
    if (!host || !LOOPBACK_HOSTS.has(host.hostname) || host.port !== String(this._port)) throw new HttpError(403, 'forbidden_host', 'http.forbiddenHost');
    const origin = req.headers.origin;
    if (origin !== undefined) {
      const parsed = parseHost(origin.replace(/^https?:\/\//, ''));
      if (!parsed || !LOOPBACK_HOSTS.has(parsed.hostname) || parsed.port !== String(this._port)) throw new HttpError(403, 'forbidden_origin', 'http.forbiddenOrigin');
    }
  }

  async _handle(req, res) {
    this._guard(req);
    const { pathname, searchParams } = new URL(req.url, `http://127.0.0.1:${this._port}`);
    const method = req.method;

    const asset = pathname === '/' ? 'index.html' : pathname.slice(1);
    if (Object.hasOwn(ASSET_TYPES, asset)) return this._sendAsset(req, res, asset);

    const deckFile = DECK_FILE_PATTERN.exec(pathname);
    if (deckFile) return this._sendDeckFile(req, res, deckFile[1], deckFile[2]);

    if (pathname === '/api/state' && method === 'GET') return sendJson(res, 200, this.snapshot());
    if (pathname === '/api/events' && method === 'GET') return this._sendEvents(req, res);
    if (pathname === '/api/annotations' && method === 'GET') {
      const draft = searchParams.get('draft');
      return sendJson(res, 200, this.annotationsFor(draft ?? Number(searchParams.get('slide'))));
    }
    if (pathname === '/api/annotations/add' && method === 'POST') {
      const body = await readJsonBody(req, MAX_STROKE_BODY_BYTES);
      return sendJson(res, 200, this.addStroke(targetOf(body), body.stroke));
    }
    if (pathname === '/api/annotations/remove' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.removeStrokes(targetOf(body), body.ids));
    }
    if (pathname === '/api/annotations/clear' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.clearAnnotations(targetOf(body)));
    }
    if (pathname === '/api/comments' && method === 'GET') {
      const draft = searchParams.get('draft');
      return sendJson(res, 200, this.commentsFor(draft ?? Number(searchParams.get('slide'))));
    }
    if (pathname === '/api/comments/add' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.addComment(targetOf(body), { text: body.text, anchor: body.anchor }));
    }
    if (pathname === '/api/comments/update' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.updateComment(targetOf(body), body.id, body.text));
    }
    if (pathname === '/api/comments/reply' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.replyComment(targetOf(body), body.id, body.text));
    }
    if (pathname === '/api/comments/resolve' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.resolveComment(targetOf(body), body.id, body.resolved));
    }
    if (pathname === '/api/comments/remove' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.removeComment(targetOf(body), body.id));
    }
    if (pathname === '/api/drafts/add' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, this.addDraft({ after: body.after, afterDraft: body.afterDraft }));
    }
    if (pathname === '/api/drafts/move' && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.id !== 'string') throw new HttpError(400, 'invalid_input', 'input.idString');
      if (body.afterDraft !== undefined && typeof body.afterDraft !== 'string') throw new HttpError(400, 'invalid_input', 'input.afterDraftString');
      return sendJson(res, 200, this.moveDraft(body.id, { after: body.after, afterDraft: body.afterDraft }));
    }
    if (pathname === '/api/drafts/remove' && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.id !== 'string') throw new HttpError(400, 'invalid_input', 'input.idString');
      return sendJson(res, 200, this.removeDraft(body.id));
    }
    if (pathname === '/api/sketchbook' && method === 'POST') {
      const body = await readJsonBody(req);
      await this.newSketchbook(typeof body.name === 'string' ? body.name : undefined);
      return sendJson(res, 202, { ok: true });
    }
    if (pathname === '/api/new-deck/suggest' && method === 'GET') return sendJson(res, 200, { path: await this.suggestNewDeckPath() });
    if (pathname === '/api/new-deck' && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.path !== 'string') throw new HttpError(400, 'invalid_input', 'input.pathString');
      if (body.template !== undefined && typeof body.template !== 'string') throw new HttpError(400, 'invalid_input', 'input.templateString');
      await this.newDeck(body.path, { template: body.template });
      return sendJson(res, 202, { ok: true });
    }
    if (pathname === '/api/slides/move' && method === 'POST') {
      const body = await readJsonBody(req);
      if (!Number.isInteger(body.from) || !Number.isInteger(body.to)) throw new HttpError(400, 'invalid_input', 'input.fromToInts');
      return sendJson(res, 200, await this.moveSlide(body.from, body.to));
    }
    if (pathname === '/api/slides/delete' && method === 'POST') {
      const body = await readJsonBody(req);
      if (!Number.isInteger(body.slide)) throw new HttpError(400, 'invalid_input', 'input.slideInt');
      return sendJson(res, 200, await this.deleteSlide(body.slide));
    }
    if ((pathname === '/api/history/undo' || pathname === '/api/history/redo') && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.id !== 'string') throw new HttpError(400, 'invalid_input', 'input.idString');
      return sendJson(res, 200, await this.historyStep(pathname.endsWith('undo') ? 'undo' : 'redo', body.id));
    }
    if (pathname === '/api/open' && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.path !== 'string') throw new HttpError(400, 'invalid_input', 'input.pathString');
      await this.openDeck(body.path);
      return sendJson(res, 202, { ok: true });
    }
    if (pathname === '/api/current' && method === 'POST') {
      const body = await readJsonBody(req);
      if (typeof body.draft === 'string') return sendJson(res, 200, this.setDraft(body.draft));
      if (!Number.isInteger(body.slide)) throw new HttpError(400, 'invalid_input', 'input.slideOrDraft');
      return sendJson(res, 200, this.setCurrent(body.slide));
    }
    if (pathname === '/api/refresh' && method === 'POST') {
      await readJsonBody(req);
      await this.refresh();
      return sendJson(res, 202, { ok: true });
    }
    if (pathname === '/api/open-external' && method === 'POST') {
      await readJsonBody(req);
      await this._openExternal();
      return sendJson(res, 202, { ok: true });
    }
    if (pathname === '/api/language' && method === 'POST') {
      const body = await readJsonBody(req);
      const language = normalizeLanguage(body.language);
      if (!language) throw new HttpError(400, 'invalid_input', 'input.languageInvalid', { languages: SUPPORTED_LANGUAGES });
      return sendJson(res, 200, { language: await this.changeLanguage(language) });
    }
    if (pathname === '/api/comments-width' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, { width: await this.changeCommentsWidth(body.width ?? null) });
    }
    if (pathname === '/api/rail-height' && method === 'POST') {
      const body = await readJsonBody(req);
      return sendJson(res, 200, { height: await this.changeRailHeight(body.height ?? null) });
    }
    if (pathname === '/api/theme' && method === 'POST') {
      const body = await readJsonBody(req);
      if (!isTheme(body.theme)) throw new HttpError(400, 'invalid_input', 'input.themeInvalid', { themes: THEMES });
      return sendJson(res, 200, { theme: await this.changeTheme(body.theme) });
    }
    if (pathname === '/api/review/submit' && method === 'POST') {
      await readJsonBody(req);
      return sendJson(res, 200, await this.submitReview());
    }
    throw new HttpError(404, 'not_found', 'http.notFound');
  }

  _sendAsset(req, res, name) {
    if (req.method !== 'GET' && req.method !== 'HEAD') throw new HttpError(405, 'method_not_allowed', 'http.getOnly');
    const body = this._assets.get(name);
    const headers = {
      'Content-Type': ASSET_TYPES[name],
      'Content-Length': body.length,
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer',
    };
    if (name === 'index.html') headers['Content-Security-Policy'] = CONTENT_SECURITY_POLICY;
    res.writeHead(200, headers);
    res.end(req.method === 'HEAD' ? undefined : body);
  }

  _sendDeckFile(req, res, key, file) {
    if (req.method !== 'GET') throw new HttpError(405, 'method_not_allowed', 'http.getOnly');
    const target = path.join(this.deckCache.keyDir(key), file);
    fs.stat(target, (error, stat) => {
      if (error || !stat.isFile()) return sendJson(res, 404, { error: { code: 'not_found', message: this.t('http.imageNotFound') } });
      res.writeHead(200, {
        'Content-Type': 'image/png',
        'Content-Length': stat.size,
        'Cache-Control': 'private, max-age=31536000, immutable',
        'X-Content-Type-Options': 'nosniff',
      });
      fs.createReadStream(target).on('error', () => res.destroy()).pipe(res);
    });
  }

  _sendEvents(req, res) {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-store',
      Connection: 'keep-alive',
      'X-Content-Type-Options': 'nosniff',
    });
    res.write(`retry: 2000\nevent: state\ndata: ${JSON.stringify(this.snapshot())}\n\n`);
    this._clients.add(res);
    req.on('close', () => this._clients.delete(res));
  }

  async _openExternal() {
    const deck = this._requireDeck();
    if (isSketchbook(deck)) throw new ViewerError('not_supported', 'error.sketchbookNoExternal');
    const source = await resolveDeckSource(deck.source.path);
    const [command, args] = process.platform === 'win32' ? ['explorer.exe', [source.path]] : process.platform === 'darwin' ? ['open', [source.path]] : ['xdg-open', [source.path]];
    const child = spawn(command, args, { detached: true, stdio: 'ignore' });
    child.once('error', (error) => this.log(this.t('log.openExternalFailed', { detail: error.message })));
    child.unref();
  }
}
