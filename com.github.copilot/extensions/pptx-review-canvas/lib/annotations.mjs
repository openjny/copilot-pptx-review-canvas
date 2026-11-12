import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { ViewerError } from './errors.mjs';
import { MAX_COMMENTS_PER_PAGE, MAX_REPLIES_PER_COMMENT, newComment, newReply, validateComment, validateCommentText } from './comments.mjs';
import { createI18n, msg } from './i18n.mjs';
import { hashPath, normalizePathKey, quarantine, readRecord, writeRecord } from './store.mjs';

export const TOOLS = ['pen', 'highlighter'];
export const MAX_STROKES_PER_SLIDE = 400;
export const MAX_POINTS_PER_SLIDE = 80_000;
export const MAX_DRAFTS = 40;
const MAX_POINTS_PER_STROKE = 4000;
const MIN_WIDTH = 0.0005;
const MAX_WIDTH = 0.06;
const DRAFT_ID = /^d[0-9a-f]{12}$/;
const FORMAT = 'pptx-review-canvas.annotations';
const VERSION = 1;
const MAX_ARCHIVES = 200;
const DRAFT_TITLE = 'New slide';
// Records written by the Japanese-only versions name draft pages "新規スライド案 N" (the oldest) or "新しいスライド N"; both are read as "New slide N".
const LEGACY_DRAFT_TITLE = /^(?:新規スライド案|新しいスライド)(?= |$)/;

const invalid = (messageKey, params) => new ViewerError('invalid_input', messageKey, params);

/** How much a slide's record holds that the user wrote: strokes and comments. */
const weigh = (entry) => entry.strokes.length + (entry.comments?.length ?? 0);

/** Checks a stroke from the page and returns it in canonical form: points are fractions of the slide image, clamped to 0..1. */
export function validateStroke(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw invalid('stroke.notObject');
  if (typeof input.id !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(input.id)) throw invalid('stroke.idInvalid');
  if (!TOOLS.includes(input.tool)) throw invalid('stroke.toolInvalid', { tools: TOOLS });
  if (typeof input.color !== 'string' || !/^#[0-9a-f]{6}$/i.test(input.color)) throw invalid('stroke.colorInvalid');
  if (!Number.isFinite(input.width) || input.width < MIN_WIDTH || input.width > MAX_WIDTH) throw invalid('stroke.widthRange');
  if (!Array.isArray(input.points) || input.points.length < 1 || input.points.length > MAX_POINTS_PER_STROKE) {
    throw invalid('stroke.pointsRange', { max: MAX_POINTS_PER_STROKE });
  }
  const points = input.points.map((point) => {
    if (!Array.isArray(point) || point.length !== 2 || !Number.isFinite(point[0]) || !Number.isFinite(point[1])) throw invalid('stroke.pointsShape');
    return [Math.round(Math.min(1, Math.max(0, point[0])) * 1e5) / 1e5, Math.round(Math.min(1, Math.max(0, point[1])) * 1e5) / 1e5];
  });
  return { id: input.id, tool: input.tool, color: input.color.toLowerCase(), width: Math.round(input.width * 1e5) / 1e5, points };
}

function validateStrokes(strokes, where) {
  for (const stroke of strokes) {
    try {
      validateStroke(stroke);
    } catch (error) {
      return msg('record.strokesInvalid', { where, detail: error });
    }
  }
  return null;
}

function validateComments(comments, where) {
  if (comments === undefined) return null;
  if (!Array.isArray(comments)) return msg('record.commentsNotArray', { where });
  for (const comment of comments) {
    try {
      validateComment(comment);
    } catch (error) {
      return msg('record.commentsInvalid', { where, detail: error });
    }
  }
  return null;
}

function validateRecordData(data) {
  if (!data || typeof data !== 'object') return msg('record.notObject');
  if (typeof data.path !== 'string' || typeof data.deckKey !== 'string') return msg('annotations.pathMissing');
  if (!data.slides || typeof data.slides !== 'object' || Array.isArray(data.slides)) return msg('annotations.slidesMissing');
  for (const [n, entry] of Object.entries(data.slides)) {
    if (!/^\d+$/.test(n) || !entry || typeof entry.hash !== 'string' || !Array.isArray(entry.strokes)) return msg('annotations.slideShape', { n });
    const problem = validateStrokes(entry.strokes, `slides[${n}]`) ?? validateComments(entry.comments, `slides[${n}]`);
    if (problem) return problem;
  }
  if (data.drafts !== undefined) {
    if (!Array.isArray(data.drafts)) return msg('annotations.draftsNotArray');
    for (const draft of data.drafts) {
      if (!draft || !DRAFT_ID.test(draft.id) || typeof draft.title !== 'string' || !Number.isInteger(draft.after) || draft.after < 0) return msg('annotations.draftsShape');
      if (typeof draft.afterHash !== 'string' || !Array.isArray(draft.strokes)) return msg('annotations.draftShape', { id: draft.id });
      const problem = validateStrokes(draft.strokes, `drafts[${draft.id}]`) ?? validateComments(draft.comments, `drafts[${draft.id}]`);
      if (problem) return problem;
    }
  }
  return null;
}

/**
 * The strokes of one open deck: on each slide, and on its draft pages (ideas for new slides that are not in the file).
 * A target is a slide number or a draft id. Every change bumps `rev` and marks the target as not yet read by the agent.
 */
export class AnnotationBook {
  constructor({ path: deckPath, deckKey, slideCount, hashes }) {
    this.path = deckPath;
    this.deckKey = deckKey;
    this.slideCount = slideCount;
    this.hashes = hashes;
    this.rev = 0;
    this.draftSeq = 0;
    this._slides = new Map();
    this._comments = new Map();
    this._drafts = [];
    this._unread = new Set();
    this._listeners = new Set();
  }

  /** Calls `listener` after every change (strokes added or removed, or slides read). Returns a function that stops it. */
  subscribe(listener) {
    this._listeners.add(listener);
    return () => this._listeners.delete(listener);
  }

  _emit() {
    for (const listener of this._listeners) {
      try {
        listener();
      } catch {
        // One failing listener must not keep the others from hearing about the change.
      }
    }
  }

  draft(id) {
    return this._drafts.find((draft) => draft.id === id) ?? null;
  }

  _requireDraft(id) {
    const draft = this.draft(id);
    if (!draft) throw new ViewerError('draft_not_found', 'error.draftNotFound');
    return draft;
  }

  strokes(target) {
    return typeof target === 'string' ? this.draft(target)?.strokes ?? [] : this._slides.get(target) ?? [];
  }

  _setStrokes(target, list) {
    if (typeof target === 'string') this._requireDraft(target).strokes = list;
    else if (list.length) this._slides.set(target, list);
    else this._slides.delete(target);
  }

  counts() {
    const counts = {};
    for (const [slide, strokes] of [...this._slides].sort((a, b) => a[0] - b[0])) if (strokes.length) counts[slide] = strokes.length;
    return counts;
  }

  total() {
    let total = 0;
    for (const strokes of this._slides.values()) total += strokes.length;
    return total;
  }

  // ---- comments about a page ---------------------------------------------------------------------------------

  comments(target) {
    return typeof target === 'string' ? this.draft(target)?.comments ?? [] : this._comments.get(target) ?? [];
  }

  _setComments(target, list) {
    if (typeof target === 'string') this._requireDraft(target).comments = list;
    else if (list.length) this._comments.set(target, list);
    else this._comments.delete(target);
  }

  /** How many open (not resolved) comments each slide has; slides without any are left out. */
  commentCounts() {
    const counts = {};
    for (const [slide, list] of [...this._comments].sort((a, b) => a[0] - b[0])) {
      const open = list.filter((comment) => !comment.resolved).length;
      if (open) counts[slide] = open;
    }
    return counts;
  }

  /**
   * Adds a comment (`{ text, anchor, author }`) to a slide or draft page and returns it. What the agent writes is not "unread" for the
   * agent: only the user's own comments make a page one the agent has not looked at.
   */
  addComment(target, input) {
    if (typeof target === 'string') this._requireDraft(target);
    const list = this.comments(target);
    if (list.length >= MAX_COMMENTS_PER_PAGE) throw new ViewerError('annotation_limit', 'error.commentsLimit');
    const comment = newComment(input);
    this._setComments(target, [...list, comment]);
    this._touch(target, { unread: comment.author === 'user' });
    return comment;
  }

  _requireComment(target, id) {
    const list = this.comments(target);
    const found = list.find((comment) => comment.id === id);
    if (!found) throw new ViewerError('comment_not_found', 'error.commentNotFound');
    return { list, found };
  }

  _replace(target, list, found, next) {
    this._setComments(target, list.map((comment) => (comment === found ? next : comment)));
  }

  /** Where a comment is, wherever it is: `{ target, comment }` (a slide number or a draft id), or null. */
  findComment(id) {
    for (const [slide, list] of this._comments) {
      const comment = list.find((entry) => entry.id === id);
      if (comment) return { target: slide, comment };
    }
    for (const draft of this._drafts) {
      const comment = (draft.comments ?? []).find((entry) => entry.id === id);
      if (comment) return { target: draft.id, comment };
    }
    return null;
  }

  /** Changes the text of the user's own comment. Returns the comment, or null when the text is the same. */
  updateComment(target, id, text) {
    if (typeof target === 'string') this._requireDraft(target);
    const { list, found } = this._requireComment(target, id);
    if (found.author !== 'user') throw new ViewerError('comment_not_editable', 'error.commentNotEditable');
    const next = validateCommentText(text);
    if (next === found.text) return null;
    const updated = { ...found, text: next, updatedAt: Date.now() };
    this._replace(target, list, found, updated);
    this._touch(target);
    return updated;
  }

  /** Adds a reply (`{ text, author }`) to a comment. A reply from the user is something the agent has not read; one from the agent is not. */
  replyComment(target, id, input) {
    if (typeof target === 'string') this._requireDraft(target);
    const { list, found } = this._requireComment(target, id);
    if (found.replies.length >= MAX_REPLIES_PER_COMMENT) throw new ViewerError('annotation_limit', 'error.repliesLimit');
    const reply = newReply(input);
    const updated = { ...found, replies: [...found.replies, reply], updatedAt: reply.createdAt };
    this._replace(target, list, found, updated);
    this._touch(target, { unread: reply.author === 'user' });
    return { comment: updated, reply };
  }

  /** Resolves a thread, or opens it again. Returns the comment, or null when it already was in that state. Nothing is unread by it. */
  setResolved(target, id, resolved) {
    if (typeof target === 'string') this._requireDraft(target);
    const { list, found } = this._requireComment(target, id);
    if (found.resolved === resolved) return null;
    const { resolvedAt, ...rest } = found;
    const updated = { ...rest, resolved, ...(resolved ? { resolvedAt: Date.now() } : {}) };
    this._replace(target, list, found, updated);
    this._touch(target, { unread: false });
    return updated;
  }

  /** Removes a comment and returns it with its place in the list (so an undo can put it back), or null when there is no such comment. */
  removeComment(target, id) {
    if (typeof target === 'string' && !this.draft(target)) throw new ViewerError('draft_not_found', 'error.draftNotFound');
    const list = this.comments(target);
    const index = list.findIndex((comment) => comment.id === id);
    if (index < 0) return null;
    this._setComments(target, list.filter((_, at) => at !== index));
    this._touch(target);
    return { comment: list[index], index };
  }

  /** Puts a removed comment back at `index` (an undo). */
  reinsertComment(target, comment, index) {
    if (typeof target === 'string') this._requireDraft(target);
    const list = this.comments(target);
    if (list.some((entry) => entry.id === comment.id)) throw new ViewerError('invalid_input', 'error.commentExists');
    if (list.length >= MAX_COMMENTS_PER_PAGE) throw new ViewerError('annotation_limit', 'error.commentsLimit');
    const next = [...list];
    next.splice(Math.min(Math.max(index, 0), next.length), 0, comment);
    this._setComments(target, next);
    this._touch(target);
  }

  /** The slides that have strokes or comments. */
  annotatedSlides() {
    return [...new Set([...Object.keys(this.counts()), ...Object.keys(this.commentCounts())].map(Number))].sort((a, b) => a - b);
  }

  unreadSlides() {
    return [...this._unread]
      .filter((target) => typeof target === 'number' && (this.strokes(target).length > 0 || this.comments(target).length > 0))
      .sort((a, b) => a - b);
  }

  /** The draft pages in the order they appear: after the slide they follow, and in the order they were inserted there. */
  drafts() {
    return [...this._drafts]
      .sort((a, b) => a.after - b.after)
      .map((draft) => ({ id: draft.id, title: draft.title, after: draft.after, count: draft.strokes.length, comments: (draft.comments ?? []).filter((comment) => !comment.resolved).length, unread: this._unread.has(draft.id) }));
  }

  unreadDrafts() {
    return this.drafts().filter((draft) => draft.unread);
  }

  /** Puts strokes and comments in place without marking them changed (used when a record is loaded). */
  restore(slide, strokes, unread, comments = []) {
    if (strokes.length) this._slides.set(slide, strokes);
    if (comments.length) this._comments.set(slide, comments.map(validateComment));
    if (unread) this._unread.add(slide);
  }

  restoreDraft(draft, unread) {
    const title = String(draft.title).replace(LEGACY_DRAFT_TITLE, DRAFT_TITLE);
    this._drafts.push({ id: draft.id, title, after: draft.after, afterHash: draft.afterHash, createdAt: draft.createdAt ?? Date.now(), strokes: draft.strokes, comments: (draft.comments ?? []).map(validateComment) });
    if (unread) this._unread.add(draft.id);
  }

  _touch(target, { unread = true } = {}) {
    this.rev += 1;
    if (unread) this._unread.add(target);
    this._emit();
  }

  /**
   * Adds a blank draft page right after slide `after` (0 puts it before the first slide), or right after the draft `afterDraft`.
   * The agent builds a real slide from it later, so it records which slide it follows.
   */
  addDraft({ after = 0, afterDraft } = {}) {
    if (this._drafts.length >= MAX_DRAFTS) throw new ViewerError('annotation_limit', 'error.draftsLimit');
    const { follows, position } = this._place(this._drafts, { after, afterDraft });
    const draft = {
      id: `d${randomUUID().replace(/-/g, '').slice(0, 12)}`,
      title: `${DRAFT_TITLE} ${++this.draftSeq}`,
      after: follows,
      afterHash: follows > 0 ? this.hashes[follows - 1] ?? '' : '',
      createdAt: Date.now(),
      strokes: [],
      comments: [],
    };
    this._drafts.splice(position, 0, draft);
    this.rev += 1;
    this._unread.add(draft.id);
    this._emit();
    return { id: draft.id, title: draft.title, after: draft.after };
  }

  /** Where in `list` a draft goes when it follows slide `after` (before the drafts already there) or the draft `afterDraft`. */
  _place(list, { after = 0, afterDraft }) {
    if (afterDraft !== undefined) {
      const anchor = list.find((draft) => draft.id === afterDraft);
      if (!anchor) throw new ViewerError('draft_not_found', 'error.draftNotFound');
      return { follows: anchor.after, position: list.indexOf(anchor) + 1 };
    }
    if (!Number.isInteger(after) || after < 0 || after > this.slideCount) throw new ViewerError('slide_out_of_range', 'error.afterRange', { max: this.slideCount });
    const first = list.findIndex((draft) => draft.after === after);
    return { follows: after, position: first >= 0 ? first : list.length };
  }

  /** Moves a draft page to follow slide `after` or the draft `afterDraft`. Returns `{ changed }`; nothing happens when it is already there. */
  moveDraft(id, { after, afterDraft } = {}) {
    const draft = this._requireDraft(id);
    if (afterDraft === id) throw new ViewerError('invalid_input', 'error.moveDraftSelf');
    const rest = this._drafts.filter((entry) => entry !== draft);
    const { follows, position } = this._place(rest, { after, afterDraft });
    const before = this.drafts().map((entry) => entry.id).join();
    const previous = { after: draft.after, afterHash: draft.afterHash };
    rest.splice(position, 0, draft);
    draft.after = follows;
    draft.afterHash = follows > 0 ? this.hashes[follows - 1] ?? '' : '';
    const original = this._drafts;
    this._drafts = rest;
    if (this.drafts().map((entry) => entry.id).join() === before && previous.after === follows) {
      this._drafts = original;
      Object.assign(draft, previous);
      return { changed: false, after: follows };
    }
    this.rev += 1;
    this._unread.add(id);
    this._emit();
    return { changed: true, after: follows };
  }

  /** Removes the draft and returns it (with its strokes), or null when there is no such draft. */
  removeDraft(id) {
    const index = this._drafts.findIndex((draft) => draft.id === id);
    if (index < 0) return null;
    const [draft] = this._drafts.splice(index, 1);
    this._unread.delete(id);
    this.rev += 1;
    this._emit();
    return draft;
  }

  /** Where the draft sits among all drafts (the order they were inserted in), or -1. Used to put it back in the same place. */
  draftIndex(id) {
    return this._drafts.findIndex((draft) => draft.id === id);
  }

  isUnread(target) {
    return this._unread.has(target);
  }

  /** Puts back a draft that was removed (an undo), at `index` among the drafts and after the same slide, or the last one if that slide is gone. */
  reinsertDraft(draft, index, unread) {
    if (this.draft(draft.id)) throw new ViewerError('invalid_input', 'error.draftExists');
    if (this._drafts.length >= MAX_DRAFTS) throw new ViewerError('annotation_limit', 'error.draftsLimit');
    const after = Math.min(draft.after, this.slideCount);
    this._drafts.splice(Math.min(Math.max(index, 0), this._drafts.length), 0, { ...draft, after, afterHash: after > 0 ? this.hashes[after - 1] ?? '' : '' });
    if (unread) this._unread.add(draft.id);
    this.rev += 1;
    this._emit();
  }

  add(target, stroke) {
    const list = this.strokes(target);
    if (typeof target === 'string') this._requireDraft(target);
    if (list.some((existing) => existing.id === stroke.id)) return false;
    const points = list.reduce((sum, existing) => sum + existing.points.length, 0) + stroke.points.length;
    if (list.length >= MAX_STROKES_PER_SLIDE || points > MAX_POINTS_PER_SLIDE) {
      throw new ViewerError('annotation_limit', 'error.strokesLimit');
    }
    this._setStrokes(target, [...list, stroke]);
    this._touch(target);
    return true;
  }

  remove(target, ids) {
    const list = this.strokes(target);
    const wanted = new Set(ids);
    const kept = list.filter((stroke) => !wanted.has(stroke.id));
    if (kept.length === list.length) return 0;
    this._setStrokes(target, kept);
    this._touch(target);
    return list.length - kept.length;
  }

  clear(target) {
    const list = this.strokes(target);
    if (!list.length) return 0;
    this._setStrokes(target, []);
    this._touch(target);
    return list.length;
  }

  markRead(targets) {
    const list = targets ?? [...this._unread];
    let changed = false;
    for (const target of list) changed = this._unread.delete(target) || changed;
    if (changed) this._emit();
    return changed;
  }

  /**
   * The deck's slides were rearranged (`order[p - 1]` is the old number of the slide now at p), possibly with some deleted (then
   * `order` is shorter): every stroke, unread mark and draft page goes with the slide it belonged to. A draft page that followed
   * a deleted slide follows the slide before it now.
   */
  rearrange({ order, deckKey, hashes }) {
    const placed = new Map(order.map((old, i) => [old, i + 1]));
    const deleted = order.length < this.slideCount;
    const slides = new Map();
    for (const [n, strokes] of this._slides) if (placed.has(n)) slides.set(placed.get(n), strokes);
    this._slides = slides;
    const comments = new Map();
    for (const [n, list] of this._comments) if (placed.has(n)) comments.set(placed.get(n), list);
    this._comments = comments;
    const unread = new Set();
    for (const target of this._unread) {
      if (typeof target === 'string') unread.add(target);
      else if (placed.has(target)) unread.add(placed.get(target));
    }
    this._unread = unread;
    if (deleted) this._drafts = [...this._drafts].sort((a, b) => a.after - b.after);
    for (const draft of this._drafts) {
      const before = draft.after;
      draft.after = before > 0 ? (deleted ? order.filter((old) => old <= before).length : placed.get(before) ?? 0) : 0;
      draft.afterHash = draft.after > 0 ? hashes[draft.after - 1] ?? '' : '';
    }
    this.deckKey = deckKey;
    this.hashes = hashes;
    this.slideCount = hashes.length;
    this.rev += 1;
    this._emit();
  }

  /**
   * The slide at `slide` is back (a deletion was undone): the slides from there on move down one, it gets its marks back, and the
   * draft pages in `draftIds` (which followed it before it was deleted) follow it again.
   */
  insertSlide({ slide, strokes, comments = [], unread, draftIds, deckKey, hashes }) {
    const slides = new Map();
    for (const [n, list] of this._slides) slides.set(n >= slide ? n + 1 : n, list);
    if (strokes.length) slides.set(slide, strokes);
    this._slides = slides;
    const notes = new Map();
    for (const [n, list] of this._comments) notes.set(n >= slide ? n + 1 : n, list);
    if (comments.length) notes.set(slide, comments);
    this._comments = notes;
    const marked = new Set();
    for (const target of this._unread) marked.add(typeof target === 'string' || target < slide ? target : target + 1);
    if (unread && (strokes.length || comments.length)) marked.add(slide);
    this._unread = marked;
    const back = new Set(draftIds);
    for (const draft of this._drafts) {
      if (back.has(draft.id) && draft.after === slide - 1) draft.after = slide;
      else if (draft.after >= slide) draft.after += 1;
      draft.afterHash = draft.after > 0 ? hashes[draft.after - 1] ?? '' : '';
    }
    this.deckKey = deckKey;
    this.hashes = hashes;
    this.slideCount = hashes.length;
    this.rev += 1;
    this._emit();
  }

  toRecord() {
    const slides = {};
    const numbers = new Set([...this._slides.keys(), ...this._comments.keys()]);
    for (const slide of [...numbers].sort((a, b) => a - b)) {
      const comments = this._comments.get(slide) ?? [];
      slides[slide] = { hash: this.hashes[slide - 1] ?? '', strokes: this._slides.get(slide) ?? [], unread: this._unread.has(slide), ...(comments.length ? { comments } : {}) };
    }
    const drafts = this._drafts.map((draft) => ({ ...draft, unread: this._unread.has(draft.id) }));
    return { path: this.path, deckKey: this.deckKey, updatedAt: Date.now(), draftSeq: this.draftSeq, slides, drafts };
  }
}

/** Keeps each deck's strokes on disk, one file per deck path, and decides which strokes survive when the deck is rebuilt. */
export class AnnotationStore {
  constructor(root, { log = () => {}, i18n = createI18n() } = {}) {
    this.dir = path.join(root, 'annotations');
    this.archiveDir = path.join(this.dir, 'archive');
    this.log = log;
    this.i18n = i18n;
    this._writes = new Map();
    this._live = new Map();
    this._attaching = new Map();
  }

  fileFor(deckPath) {
    return path.join(this.dir, `${hashPath(deckPath)}.json`);
  }

  /**
   * Returns the book for a freshly rendered deck. All viewers that show the same deck share one book, so what one of them draws
   * is seen by the others and no viewer overwrites another's strokes. `order` (see `DeckCache.derive`) says the deck was only
   * rearranged, or had slides deleted (a shorter `order`): the shared book is then rearranged in place, so every stroke and draft
   * page follows its slide exactly, without comparing pictures, and every viewer keeps the same book. The strokes of a deleted
   * slide are set aside in the archive. `insert` (`{ slide, strokes, unread, draftIds }`) says a deleted slide is back (an undo):
   * the book then puts it and its marks in place.
   */
  attach(deck, { order, insert } = {}) {
    const id = normalizePathKey(deck.source.path);
    const previous = this._attaching.get(id) ?? Promise.resolve();
    const next = previous.catch(() => {}).then(async () => {
      const live = this._live.get(id);
      if (live?.deckKey === deck.key) return { book: live.book, kept: live.book.total(), moved: 0, archived: 0, changed: false };
      const known = live?.book.slideCount ?? 0;
      if (live && insert && deck.slideCount === known + 1 && insert.slide >= 1 && insert.slide <= deck.slideCount) {
        live.book.insertSlide({ ...insert, deckKey: deck.key, hashes: deck.slides.map((slide) => slide.hash) });
        this._live.set(id, { deckKey: deck.key, book: live.book });
        await this.save(live.book);
        return { book: live.book, kept: live.book.total(), moved: 0, archived: 0, changed: true };
      }
      if (live && order?.length === deck.slideCount && order.length <= known && order.every((old) => Number.isInteger(old) && old >= 1 && old <= known)) {
        const before = order.length < known ? live.book.toRecord() : null;
        live.book.rearrange({ order, deckKey: deck.key, hashes: deck.slides.map((slide) => slide.hash) });
        this._live.set(id, { deckKey: deck.key, book: live.book });
        const lost = Object.fromEntries(Object.entries(before?.slides ?? {}).filter(([n, entry]) => weigh(entry) > 0 && !order.includes(Number(n))));
        const archived = Object.values(lost).reduce((sum, entry) => sum + weigh(entry), 0);
        if (archived > 0) await this._archive(before, lost);
        await this.save(live.book);
        return { book: live.book, kept: live.book.total(), moved: 0, archived, changed: true };
      }
      const result = await this._build(deck, live?.book);
      this._live.set(id, { deckKey: deck.key, book: result.book });
      return result;
    });
    this._attaching.set(id, next);
    return next.finally(() => {
      if (this._attaching.get(id) === next) this._attaching.delete(id);
    });
  }

  /**
   * Builds the book from the previous one (or else from the stored record). Strokes stay on a slide whose picture did not change
   * (the agent has not fixed it yet), follow a slide that moved, and are archived when their slide changed.
   * `deck` is a manifest: `{ key, source, slideCount, slides: [{ hash }] }`.
   */
  async _build(deck, previousBook) {
    const hashes = deck.slides.map((slide) => slide.hash);
    const book = new AnnotationBook({ path: deck.source.path, deckKey: deck.key, slideCount: deck.slideCount, hashes });
    const file = this.fileFor(deck.source.path);
    let record;
    if (previousBook) {
      record = { status: 'ok', data: previousBook.toRecord() };
    } else {
      try {
        record = await readRecord(file, { format: FORMAT, version: VERSION, validate: validateRecordData });
      } catch (error) {
        this.log(this.i18n.t('log.annotationsUnreadable', { detail: error.message }));
        return { book, kept: 0, moved: 0, archived: 0, changed: false };
      }
    }
    if (record.status === 'invalid') {
      const aside = await quarantine(file).catch(() => null);
      this.log(this.i18n.t(aside ? 'log.annotationsInvalidKept' : 'log.annotationsInvalid', { reason: record.reason, kept: aside }));
      return { book, kept: 0, moved: 0, archived: 0, changed: false };
    }
    if (record.status === 'missing') return { book, kept: 0, moved: 0, archived: 0, changed: false };

    const stored = record.data;
    if (stored.deckKey === deck.key) {
      for (const [n, entry] of Object.entries(stored.slides)) {
        const slide = Number(n);
        if (slide <= deck.slideCount) book.restore(slide, entry.strokes, entry.unread === true, entry.comments ?? []);
      }
      this._restoreDrafts(book, stored);
      return { book, kept: book.total(), moved: 0, archived: 0, changed: false };
    }

    const taken = new Set();
    const lost = {};
    let kept = 0;
    let moved = 0;
    let archived = 0;
    for (const [n, entry] of Object.entries(stored.slides)) {
      const weight = weigh(entry);
      if (!weight) continue;
      const slide = Number(n);
      let target = hashes[slide - 1] === entry.hash && !taken.has(slide) ? slide : -1;
      if (target < 0) {
        const found = hashes.findIndex((hash, index) => hash === entry.hash && !taken.has(index + 1));
        target = found >= 0 ? found + 1 : -1;
      }
      if (target < 0 || entry.hash === '') {
        lost[n] = entry;
        archived += weight;
        continue;
      }
      taken.add(target);
      book.restore(target, entry.strokes, entry.unread === true, entry.comments ?? []);
      kept += weight;
      if (target !== slide) moved += weight;
    }
    if (archived > 0) await this._archive(stored, lost);
    this._restoreDrafts(book, stored);
    this._adopt(deck, book);
    await this.save(book);
    return { book, kept, moved, archived, changed: true };
  }

  /** Draft pages are ideas for new slides, so they never go away on a rebuild; each follows the slide it was placed after. */
  _restoreDrafts(book, stored) {
    book.draftSeq = Math.max(book.draftSeq, Number.isInteger(stored.draftSeq) ? stored.draftSeq : 0);
    for (const draft of stored.drafts ?? []) {
      let after = Math.min(draft.after, book.slideCount);
      if (draft.afterHash) {
        let best = -1;
        book.hashes.forEach((hash, index) => {
          if (hash === draft.afterHash && (best < 0 || Math.abs(index + 1 - draft.after) < Math.abs(best + 1 - draft.after))) best = index;
        });
        if (best >= 0) after = best + 1;
      }
      book.restoreDraft({ ...draft, after, afterHash: after > 0 ? book.hashes[after - 1] ?? '' : '' }, draft.unread === true);
    }
  }

  _adopt(deck, book) {
    this._live.set(normalizePathKey(deck.source.path), { deckKey: deck.key, book });
  }

  /** Writes the book. Writes to one deck's file are serialized, so a quick succession of strokes cannot interleave. */
  save(book) {
    const live = this._live.get(normalizePathKey(book.path));
    if (live && live.book !== book) return Promise.resolve();
    const file = this.fileFor(book.path);
    const record = book.toRecord();
    const previous = this._writes.get(file) ?? Promise.resolve();
    const next = previous.catch(() => {}).then(async () => {
      if (Object.keys(record.slides).length === 0 && record.drafts.length === 0) {
        const exists = await fs.access(file).then(() => true, () => false);
        if (!exists) return;
      }
      await writeRecord(file, FORMAT, VERSION, record);
    });
    this._writes.set(file, next);
    return next.finally(() => {
      if (this._writes.get(file) === next) this._writes.delete(file);
    });
  }

  /** Sets removed draft pages aside (they are not deleted outright, so they can be recovered by hand). */
  archiveDrafts(book, drafts) {
    if (!drafts.length) return Promise.resolve();
    return this._archive(book.toRecord(), {}, drafts);
  }

  async _archive(stored, lost, drafts = []) {
    const name = `${hashPath(stored.path)}-${Date.now()}-${randomUUID().slice(0, 6)}.json`;
    try {
      await writeRecord(path.join(this.archiveDir, name), FORMAT, VERSION, { ...stored, slides: lost, drafts });
      const names = (await fs.readdir(this.archiveDir)).sort();
      for (const old of names.slice(0, Math.max(0, names.length - MAX_ARCHIVES))) await fs.rm(path.join(this.archiveDir, old), { force: true });
    } catch (error) {
      this.log(this.i18n.t('log.annotationsArchiveFailed', { detail: error.message }));
    }
  }

  async flush() {
    await Promise.all([...this._writes.values()].map((write) => write.catch(() => {})));
  }
}
