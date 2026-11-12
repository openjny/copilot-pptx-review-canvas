import { createComments } from './comments.js';
import { createDrag } from './drag.js';
import { icon, mountIcons } from './icons.js';
import { applyTranslations, getLanguage, normalizeLanguage, setLanguage, t, tn } from './i18n.js';
import { createInk } from './ink.js';
import { createResize } from './resize.js';
import { createReview } from './review.js';

const $ = (id) => document.getElementById(id);

const el = {
  app: $('app'),
  title: $('deck-title'),
  meta: $('deck-meta'),
  notesButton: $('btn-notes'),
  refresh: $('btn-refresh'),
  external: $('btn-external'),
  newButton: $('btn-new'),
  openButton: $('btn-open'),
  helpButton: $('btn-help'),
  help: $('help'),
  langButtons: [...document.querySelectorAll('.lang__btn[data-lang]')],
  openPanel: $('open-panel'),
  openForm: $('open-form'),
  openPath: $('open-path'),
  newForm: $('new-form'),
  newPath: $('new-path'),
  newTemplate: $('new-template'),
  newSubmit: $('new-submit'),
  sketchForm: $('sketch-form'),
  sketchName: $('sketch-name'),
  recent: $('recent'),
  recentList: $('recent-list'),
  banner: $('banner'),
  notice: $('notice'),
  noticeText: $('notice-text'),
  noticeClose: $('notice-close'),
  status: $('status'),
  progress: $('progress'),
  statusText: $('status-text'),
  live: $('slide-live'),
  layout: $('layout'),
  rail: $('rail'),
  thumbs: $('thumbs'),
  stage: $('stage'),
  box: $('slide-box'),
  image: $('slide-img'),
  canvas: $('ink-canvas'),
  empty: $('stage-empty'),
  notes: $('notes'),
  notesBody: $('notes-body'),
  commentsPanel: $('comments'),
  commentsList: $('comments-list'),
  commentsTitle: $('comments-title'),
  commentsClose: $('comments-close'),
  commentsAdd: $('comments-add'),
  commentsToggle: $('btn-comments'),
  commentsCount: $('comments-count'),
  commentsResize: $('comments-resize'),
  commentLayer: $('comment-layer'),
  submitButton: $('btn-submit'),
  submitIcon: $('btn-submit-icon'),
  submitLabel: $('btn-submit-label'),
  submitPanel: $('submit-panel'),
  submitSummary: $('submit-summary'),
  submitNote: $('submit-note'),
  submitCancel: $('submit-cancel'),
  submitSend: $('submit-send'),
};

let state = null;
let thumbsKey = null;
let draftsKey = null;
let recentRendered = null;
let notesVisible = readStored('pptx-viewer.notes') === '1';
let openPanelRequested = false;
let panelWasVisible = false;
let suggestedPath = null;
let moving = false;
let creating = false;
let localError = null;
let navigationsInFlight = 0;
let lastActive = null;
let dismissedNotice = Number(readStored('pptx-viewer.notice') ?? 0);

const pad = (n) => String(n).padStart(4, '0');
const slideUrl = (deck, n) => `deck/${deck.key}/s${pad(n)}.png?v=${deck.renderedAt}`;
const thumbUrl = (deck, n) => `deck/${deck.key}/t${pad(n)}.png?v=${deck.renderedAt}`;
const slideTitle = (slide) => slide.title || t('slide.fallbackTitle', { n: slide.n });
// New slides are told apart by where they sit in the strip, so the screen does not number them.
const draftLabel = (draft) => (draft.after > 0 ? t('draft.labelAfter', { n: draft.after }) : t('draft.labelStart'));
const say = (text) => {
  el.live.textContent = text;
};

function readStored(key) {
  try {
    return sessionStorage.getItem(key);
  } catch {
    return null;
  }
}

function writeStored(key, value) {
  try {
    sessionStorage.setItem(key, value);
  } catch {
    // Storage may be blocked; the choice then simply is not remembered after a reload.
  }
}

async function getJson(path) {
  const response = await fetch(path, { cache: 'no-store' });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload?.error?.message ?? t('error.request', { status: response.status }));
  return payload;
}

async function post(path, body = {}) {
  const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload?.error?.message ?? t('error.request', { status: response.status }));
  return payload;
}

function reportError(error) {
  localError = error instanceof Error ? error.message : String(error);
  renderBanner();
}

mountIcons();

const deckHistory = {
  peek: (kind) => (state?.phase === 'ready' && !moving ? state.history?.[kind] ?? null : null),
  run: runHistory,
};

const ink = createInk({
  box: el.box,
  canvas: el.canvas,
  getState: () => state,
  getJson,
  post,
  onError: reportError,
  announce: say,
  history: deckHistory,
  onSelect: (anchor) => comments.startNew(anchor),
  onVisibleChange: (visible) => comments.setMarksVisible(visible),
});

const comments = createComments({
  elements: {
    panel: el.commentsPanel,
    list: el.commentsList,
    title: el.commentsTitle,
    close: el.commentsClose,
    add: el.commentsAdd,
    layer: el.commentLayer,
    toggle: el.commentsToggle,
    count: el.commentsCount,
  },
  getState: () => state,
  getJson,
  post,
  onError: reportError,
  announce: say,
  store: { read: () => readStored('pptx-viewer.comments'), write: (value) => writeStored('pptx-viewer.comments', value) },
});

const review = createReview({
  elements: {
    button: el.submitButton,
    buttonIcon: el.submitIcon,
    buttonLabel: el.submitLabel,
    panel: el.submitPanel,
    summary: el.submitSummary,
    note: el.submitNote,
    cancel: el.submitCancel,
    send: el.submitSend,
  },
  getState: () => state,
  post,
  onError: reportError,
  announce: say,
  onOpen: () => setHelp(false),
});

// The width the user chose for the comments pane is saved by the server (the address of this page changes every time, so the browser
// could not keep it). It is read from the first state only; later changes come from this page's own dragging.
const commentsResize = createResize({
  handle: el.commentsResize,
  layout: el.layout,
  onCommit: (width) => post('api/comments-width', { width }).catch(reportError),
});
let commentsWidthAdopted = false;

function adoptCommentsWidth(width) {
  if (commentsWidthAdopted) return;
  commentsWidthAdopted = true;
  if (Number.isInteger(width)) commentsResize.set(width);
}

function applyState(next) {
  if (state && next.rev < state.rev) return;
  const navigating = navigationsInFlight > 0 && state;
  state = { ...next, current: navigating ? state.current : next.current, draft: navigating ? state.draft : next.draft };
  if (next.phase !== 'error') localError = null;
  adoptLanguage(next.language);
  adoptCommentsWidth(next.commentsWidth);
  render();
}

/** Switches the page to `value` ('en' or 'ja') when it is a supported language other than the current one. The render that follows redraws the rest. */
function adoptLanguage(value) {
  const language = normalizeLanguage(value);
  if (!language || language === getLanguage()) return false;
  setLanguage(language);
  applyTranslations();
  renderLanguageSwitch();
  renderNewSubmit();
  // Text built in code is cached (the strip, the recent list, the comments pane); drop those caches so it is built again in the new language.
  thumbsKey = null;
  draftsKey = null;
  recentRendered = null;
  if (state) comments.refresh();
  return true;
}

function renderLanguageSwitch() {
  for (const button of el.langButtons) button.setAttribute('aria-pressed', String(button.dataset.lang === getLanguage()));
}

function chooseLanguage(language) {
  if (!language || language === getLanguage()) return;
  post('api/language', { language }).then((result) => {
    adoptLanguage(result?.language ?? language);
    if (state) render();
  }).catch(reportError);
}

function applyProgress(progress) {
  if (!state || state.phase !== 'rendering') return;
  state.progress = progress;
  renderStatus();
}

function render() {
  el.app.dataset.phase = state.phase;
  renderHeader();
  renderBanner();
  renderNotice();
  renderStatus();
  renderOpenPanel();
  renderThumbs();
  renderSlide();
  renderNotes();
  renderControls();
  ink.sync(state);
  comments.sync(state);
  review.sync();
}

function renderNotice() {
  const notice = state?.notice;
  const visible = Boolean(notice) && notice.id > dismissedNotice;
  el.notice.hidden = !visible;
  el.noticeText.textContent = visible ? notice.text : '';
}

function isSketchbook() {
  return state?.deck?.kind === 'sketchbook';
}

const recentLabel = (entry) => (/\.sketch$/i.test(entry.name) ? t('recent.sketchbook', { name: entry.name.replace(/\.sketch$/i, '') }) : entry.name);

function renderHeader() {
  const { deck, pending } = state;
  const defaultTitle = t('app.title');
  const title = deck?.name ?? pending?.name?.replace(/\.sketch$/i, '') ?? defaultTitle;
  el.title.textContent = title;
  el.meta.textContent = deck ? (isSketchbook() ? t('deck.sketchbook') : tn('deck.meta', deck.slideCount, { path: deck.path })) : '';
  el.meta.title = deck && !isSketchbook() ? deck.path : '';
  document.title = deck ? `${title} - ${defaultTitle}` : defaultTitle;
}

function renderBanner() {
  const message = localError ?? (state?.phase === 'error' ? state.error?.message : null);
  el.banner.hidden = !message;
  el.banner.textContent = message ?? '';
}

function renderStatus() {
  const rendering = state.phase === 'rendering';
  el.status.hidden = !rendering;
  if (!rendering) return;
  const { done = 0, total = 0 } = state.progress ?? {};
  el.progress.max = Math.max(total, 1);
  el.progress.value = done;
  el.statusText.textContent = total > 0 ? t('status.rendering', { done, total }) : t('status.preparing');
}

function renderOpenPanel() {
  const visible = !state.deck || openPanelRequested;
  el.openPanel.hidden = !visible;
  el.openButton.setAttribute('aria-expanded', String(visible));
  el.newButton.setAttribute('aria-expanded', String(visible));
  if (visible && !panelWasVisible) suggestNewPath();
  panelWasVisible = visible;
  const entries = state.recent ?? [];
  const recentKey = JSON.stringify(entries);
  if (recentKey === recentRendered) return;
  recentRendered = recentKey;
  el.recentList.replaceChildren();
  el.recent.hidden = entries.length === 0;
  for (const entry of entries) {
    const item = document.createElement('li');
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'recent__item';
    const name = document.createElement('span');
    name.className = 'recent__name';
    name.textContent = recentLabel(entry);
    const where = document.createElement('span');
    where.className = 'recent__path';
    where.textContent = entry.path;
    button.append(name, where);
    button.addEventListener('click', () => openDeck(entry.path));
    item.append(button);
    el.recentList.append(item);
  }
}

// ---- the strip of thumbnails ---------------------------------------------------------------------------------

/** A thin button between two thumbnails: a "+" shows there when the pointer rests on it, and adds a slide at that place. */
function gapButton(where, label) {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'gap';
  button.tabIndex = -1;
  button.dataset.where = JSON.stringify(where);
  button.setAttribute('aria-label', label);
  button.title = t('slide.addHere');
  const plus = document.createElement('span');
  plus.className = 'gap__plus';
  plus.innerHTML = icon('plus');
  button.append(plus);
  return button;
}

function slideItem(deck, slide) {
  const item = document.createElement('li');
  item.className = 'page';
  item.dataset.kind = 'slide';
  item.dataset.n = String(slide.n);
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'thumb';
  button.dataset.n = String(slide.n);
  button.dataset.label = t('slide.label', { n: slide.n, title: slide.title || t('slide.untitled') });
  button.setAttribute('aria-label', button.dataset.label);
  const image = document.createElement('img');
  image.src = thumbUrl(deck, slide.n);
  image.alt = '';
  image.loading = 'lazy';
  image.decoding = 'async';
  const number = document.createElement('span');
  number.className = 'thumb__n';
  number.textContent = String(slide.n);
  button.append(image, number);
  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'thumb__x thumb__x--slide';
  remove.dataset.n = String(slide.n);
  remove.title = t('slide.delete');
  remove.setAttribute('aria-label', t('slide.deleteN', { n: slide.n }));
  remove.innerHTML = icon('close');
  item.append(button, remove, gapButton({ after: slide.n }, t('slide.addAfter', { n: slide.n })));
  return item;
}

/** A slide that is not in the file yet. It looks like a slide in the strip; it only says "New" where a slide has its number. */
function draftItem(draft) {
  const item = document.createElement('li');
  item.className = 'page is-draft';
  item.dataset.kind = 'draft';
  item.dataset.draft = draft.id;
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'thumb thumb--draft';
  button.dataset.draft = draft.id;
  button.dataset.label = draftLabel(draft);
  button.setAttribute('aria-label', button.dataset.label);
  const page = document.createElement('span');
  page.className = 'thumb__page';
  const title = document.createElement('strong');
  title.textContent = t('slide.new');
  page.append(title);
  const tag = document.createElement('span');
  tag.className = 'thumb__n thumb__n--new';
  tag.textContent = t('draft.tag');
  button.append(page, tag);
  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'thumb__x';
  remove.dataset.draft = draft.id;
  remove.title = t('draft.delete');
  remove.setAttribute('aria-label', draft.after > 0 ? t('draft.deleteAfter', { n: draft.after }) : t('draft.deleteStart'));
  remove.innerHTML = icon('close');
  item.append(button, remove, gapButton({ afterDraft: draft.id }, draft.after > 0 ? t('draft.addAfter', { n: draft.after }) : t('draft.addStart')));
  return item;
}

function buildStrip(deck) {
  el.layout.style.setProperty('--ratio', `${deck.slideWidth} / ${deck.slideHeight}`);
  el.layout.style.setProperty('--aspect', String(deck.slideWidth / deck.slideHeight));
  const fragment = document.createDocumentFragment();
  const lead = document.createElement('li');
  lead.className = 'lead';
  lead.append(gapButton({ after: 0 }, t('slide.addStart')));
  fragment.append(lead);
  for (const slide of deck.slides) fragment.append(slideItem(deck, slide));
  const add = document.createElement('li');
  add.className = 'add';
  const tile = document.createElement('button');
  tile.type = 'button';
  tile.className = 'add-tile';
  tile.setAttribute('aria-label', t('slide.addEnd'));
  tile.title = t('slide.addEndTitle');
  tile.innerHTML = icon('plus');
  add.append(tile);
  fragment.append(add);
  el.thumbs.append(fragment);
}

function renderThumbs() {
  const { deck } = state;
  const addable = canEditPages();
  el.thumbs.dataset.reorderable = String(canMove());
  el.thumbs.dataset.removable = String(canRemoveSlides());
  el.thumbs.dataset.addable = String(addable);
  el.thumbs.classList.toggle('is-busy', moving);
  const key = deck ? `${deck.key}:${deck.renderedAt}` : null;
  if (key !== thumbsKey) {
    thumbsKey = key;
    draftsKey = null;
    lastActive = null;
    el.thumbs.replaceChildren();
    if (deck) buildStrip(deck);
  }
  const tile = el.thumbs.querySelector('.add-tile');
  if (tile) tile.disabled = !addable;
  renderDraftThumbs();
  let active = null;
  const marks = state.annotations ?? { counts: {}, unread: [], drafts: [] };
  const unread = new Set(marks.unread);
  for (const button of el.thumbs.querySelectorAll('.thumb')) {
    const isDraft = button.dataset.draft !== undefined;
    const draft = isDraft ? marks.drafts.find((entry) => entry.id === button.dataset.draft) : null;
    const n = Number(button.dataset.n);
    const isCurrent = isDraft ? state.draft === button.dataset.draft : !state.draft && n === state.current;
    if (isCurrent) {
      button.setAttribute('aria-current', 'true');
      active = button;
    } else {
      button.removeAttribute('aria-current');
    }
    if (isDraft) renderThumbInk(button, draft?.count ?? 0, draft?.unread ?? false, draft?.comments ?? 0);
    else renderThumbInk(button, marks.counts[n] ?? 0, unread.has(n), marks.commentCounts?.[n] ?? 0);
  }
  // Bring the thumbnail of the page on screen into view, but only when the page changes: not while the user is scrolling the strip.
  const activeKey = active ? `${key}:${active.dataset.draft ?? active.dataset.n}` : null;
  if (active && activeKey !== lastActive) active.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  lastActive = activeKey;
}

/** The thumbnails of new slides sit in the strip right after the slide they follow (or at the start). */
function renderDraftThumbs() {
  const drafts = state.annotations?.drafts ?? [];
  const key = JSON.stringify(drafts.map(({ id, title, after }) => [id, title, after]));
  if (key === draftsKey) return;
  const lead = el.thumbs.querySelector(':scope > li.lead');
  if (!lead) return;
  draftsKey = key;
  for (const item of el.thumbs.querySelectorAll(':scope > li.is-draft')) item.remove();
  const slideItems = [...el.thumbs.querySelectorAll(':scope > li.page[data-kind="slide"]')];
  const placed = new Map();
  for (const draft of drafts) {
    const item = draftItem(draft);
    const before = placed.get(draft.after) ?? (draft.after > 0 ? slideItems[draft.after - 1] : lead);
    if (!before) continue;
    before.after(item);
    placed.set(draft.after, item);
  }
}

/** The red badge on a thumbnail: how many strokes the page has, filled while the agent has not read them yet; and the blue one with the comments. */
function renderThumbInk(button, count, isUnread, notes = 0) {
  const status = isUnread ? t('thumb.unread') : t('thumb.read');
  let badge = button.querySelector('.thumb__ink');
  if (!count) {
    badge?.remove();
  } else {
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'thumb__ink';
      button.append(badge);
    }
    badge.textContent = String(count);
    badge.dataset.unread = String(isUnread);
    badge.title = tn('thumb.marks', count, { status });
  }
  let note = button.querySelector('.thumb__note');
  if (!notes) {
    note?.remove();
  } else {
    if (!note) {
      note = document.createElement('span');
      note.className = 'thumb__note';
      button.append(note);
    }
    note.textContent = String(notes);
    note.title = tn('thumb.comments', notes);
  }
  const parts = [button.dataset.label];
  if (count) parts.push(tn('thumb.marks', count, { status }));
  if (notes) parts.push(tn('thumb.comments', notes));
  button.setAttribute('aria-label', parts.join(t('common.listSeparator')));
}

function currentDraft() {
  return state?.draft ? state.annotations?.drafts?.find((entry) => entry.id === state.draft) ?? null : null;
}

function renderSlide() {
  const { deck, phase } = state;
  if (!deck) {
    el.box.hidden = true;
    el.empty.textContent = phase === 'rendering' ? t('stage.rendering') : t('stage.empty');
    return;
  }
  el.box.hidden = false;
  el.empty.textContent = '';
  const draft = currentDraft();
  if (draft) {
    el.image.hidden = true;
    const page = `d${draft.id}`;
    if (el.box.dataset.page !== page) {
      el.box.dataset.page = page;
      say(draftLabel(draft));
    }
    return;
  }
  el.image.hidden = false;
  el.box.dataset.page = `s${state.current}`;
  const slide = deck.slides[state.current - 1];
  const source = slideUrl(deck, state.current);
  if (el.image.dataset.source !== source) {
    el.image.dataset.source = source;
    el.image.src = source;
    say(t('slide.live', { n: state.current, count: deck.slideCount, title: slideTitle(slide) }));
  }
  el.image.alt = t('slide.label', { n: state.current, title: slideTitle(slide) });
  for (const n of [state.current - 1, state.current + 1]) {
    if (n >= 1 && n <= deck.slideCount) new Image().src = slideUrl(deck, n);
  }
}

function renderNotes() {
  const { deck } = state;
  el.notes.hidden = !deck || !notesVisible || Boolean(currentDraft()) || isSketchbook();
  el.notesButton.setAttribute('aria-pressed', String(notesVisible));
  if (!deck || currentDraft()) return;
  const notes = deck.slides[state.current - 1].notes;
  el.notesBody.textContent = notes || t('notes.empty');
  el.notesBody.classList.toggle('is-empty', !notes);
}

function renderControls() {
  const { deck } = state;
  el.refresh.hidden = isSketchbook();
  el.external.hidden = isSketchbook();
  el.refresh.disabled = !deck || state.phase === 'rendering';
  el.external.disabled = !deck;
  el.notesButton.hidden = isSketchbook();
}

// ---- moving between pages, adding, moving and removing slides ------------------------------------------------

/** The pages in the order the strip shows them: every slide, with the new slides that follow it. */
function pageList() {
  const deck = state?.deck;
  if (!deck) return [];
  const drafts = state.annotations?.drafts ?? [];
  const pages = [];
  const addDrafts = (after) => {
    for (const draft of drafts) if (draft.after === after) pages.push({ draft: draft.id });
  };
  addDrafts(0);
  for (let n = 1; n <= deck.slideCount; n++) {
    pages.push({ slide: n });
    addDrafts(n);
  }
  return pages;
}

const pageIndex = (pages) => pages.findIndex((page) => (state.draft ? page.draft === state.draft : page.slide === state.current));

function showSlide(n) {
  const deck = state?.deck;
  if (!deck || !Number.isFinite(n)) return;
  const target = Math.min(Math.max(1, Math.trunc(n)), deck.slideCount);
  if (target === state.current && !state.draft) return;
  state = { ...state, current: target, draft: null };
  render();
  navigationsInFlight += 1;
  post('api/current', { slide: target }).catch(reportError).finally(() => {
    navigationsInFlight -= 1;
  });
}

function showDraft(id) {
  if (!state?.deck || state.draft === id) return;
  state = { ...state, draft: id };
  render();
  navigationsInFlight += 1;
  post('api/current', { draft: id }).catch(reportError).finally(() => {
    navigationsInFlight -= 1;
  });
}

function showPage(page) {
  if (!page) return;
  if (page.draft) showDraft(page.draft);
  else showSlide(page.slide);
}

/** Moves through the pages in strip order, so the arrow keys visit new slides too. */
function step(delta) {
  const pages = pageList();
  const index = pageIndex(pages);
  if (index >= 0) showPage(pages[index + delta]);
}

/** Moving a saved slide needs a deck that can be rewritten (not a sketchbook or a legacy .ppt) and nothing else going on. */
function canMove() {
  return Boolean(state?.deck?.reorderable) && state.phase === 'ready' && !moving;
}

/** Adding, dragging and removing new slides only touch the viewer's own data, so any open deck allows them. */
function canEditPages() {
  return Boolean(state?.deck) && state.phase === 'ready' && !moving;
}

/** Deleting a saved slide rewrites the file, so it needs the same as moving one, and a slide that stays behind. */
function canRemoveSlides() {
  return canMove() && state.deck.slideCount > 1;
}

function requestMove(from, to) {
  if (!canMove() || from === to) return;
  moving = true;
  localError = null;
  renderBanner();
  renderThumbs();
  say(t('move.moving', { from, to }));
  post('api/slides/move', { from, to }).then(() => {
    say(t('move.moved', { from, to }));
  }).catch((error) => {
    say(t('move.failed'));
    reportError(error);
  }).finally(() => {
    moving = false;
    if (state) render();
  });
}

/** Moves a new slide so that `slot` other pages come before it. */
function moveDraftTo(id, slot) {
  if (!canEditPages()) return;
  const others = pageList().filter((page) => page.draft !== id);
  const anchor = slot <= 0 ? null : others[slot - 1];
  const where = !anchor ? { after: 0 } : anchor.draft ? { afterDraft: anchor.draft } : { after: anchor.slide };
  localError = null;
  post('api/drafts/move', { id, ...where }).then(() => say(t('draft.moved'))).catch(reportError);
}

/** Alt + arrow keys: one place earlier or later, for the slide or the new slide on screen. */
function moveCurrentBy(delta) {
  if (state?.draft) {
    const pages = pageList();
    const index = pageIndex(pages);
    const others = pages.length - 1;
    const slot = Math.min(Math.max(index + delta, 0), others);
    if (index >= 0 && slot !== index) moveDraftTo(state.draft, slot);
    return;
  }
  if (!canMove()) return;
  const to = state.current + delta;
  if (to >= 1 && to <= (state.deck?.slideCount ?? 0)) requestMove(state.current, to);
}

function addSlide(where) {
  if (!canEditPages()) return;
  localError = null;
  post('api/drafts/add', where).catch(reportError);
}

function addSlideAtEnd() {
  const last = pageList().at(-1);
  if (last) addSlide(last.draft ? { afterDraft: last.draft } : { after: last.slide });
}

function addSlideAfterCurrent() {
  addSlide(state.draft ? { afterDraft: state.draft } : { after: state.current });
}

/** Deleting a new slide happens at once; Ctrl+Z (the undo button) brings it back. */
function requestRemoveDraft(id) {
  if (!state?.annotations?.drafts?.some((entry) => entry.id === id) || !canEditPages()) return;
  localError = null;
  post('api/drafts/remove', { id }).then(() => say(t('draft.removed'))).catch(reportError);
}

/** Deleting a slide rewrites the PowerPoint file at once. The file as it was is kept, so the undo button (Ctrl+Z) brings the slide back. */
function requestRemoveSlide(n) {
  if (!canRemoveSlides() || !Number.isInteger(n)) return;
  moving = true;
  localError = null;
  renderBanner();
  renderThumbs();
  say(t('delete.deleting', { n }));
  post('api/slides/delete', { slide: n }).then(() => {
    say(t('delete.deleted', { n }));
  }).catch((error) => {
    say(t('delete.failed'));
    reportError(error);
  }).finally(() => {
    moving = false;
    if (state) render();
  });
}

/**
 * Undo and redo of what changed the deck or its new slides (a deleted slide). The undo of strokes is in ink.js, which asks here
 * which of the two comes first (`peek`) and hands over when it is the deletion (`run`).
 */
function runHistory(kind) {
  const entry = state?.history?.[kind];
  if (!entry || moving || state.phase !== 'ready') return;
  const rewrites = entry.kind === 'slide';
  const { label } = entry;
  if (rewrites) {
    moving = true;
    renderBanner();
    renderThumbs();
  }
  localError = null;
  say(kind === 'undo' ? t('history.undoing', { label }) : t('history.redoing', { label }));
  post(`api/history/${kind}`, { id: entry.id }).then(() => say(kind === 'undo' ? t('history.undone', { label }) : t('history.redone', { label }))).catch((error) => {
    say(t('history.failed'));
    reportError(error);
  }).finally(() => {
    if (rewrites) moving = false;
    if (state) render();
  });
}

// ---- opening, creating, and the small panels ------------------------------------------------------------------

function startSketchbook(name) {
  localError = null;
  post('api/sketchbook', { name: name.trim() }).then(() => {
    openPanelRequested = false;
    el.sketchName.value = '';
  }).catch(reportError);
}

/** A free file name for a new presentation. It only fills the field when the user has not typed anything of their own. */
function suggestNewPath() {
  getJson('api/new-deck/suggest').then(({ path }) => {
    if (!el.newPath.value.trim() || el.newPath.value === suggestedPath) el.newPath.value = path;
    suggestedPath = path;
  }).catch(() => {});
}

function renderNewSubmit() {
  el.newSubmit.textContent = creating ? t('open.creating') : t('open.createSubmit');
}

function startNewDeck() {
  const path = el.newPath.value.trim();
  if (creating || !path) return;
  creating = true;
  localError = null;
  renderBanner();
  el.newSubmit.disabled = true;
  renderNewSubmit();
  say(t('open.creatingStatus'));
  const template = el.newTemplate.value.trim();
  post('api/new-deck', { path, ...(template ? { template } : {}) }).then(() => {
    openPanelRequested = false;
    el.newPath.value = '';
    el.newTemplate.value = '';
    suggestedPath = null;
    if (state) renderOpenPanel();
  }).catch((error) => {
    say(t('open.createFailed'));
    reportError(error);
  }).finally(() => {
    creating = false;
    el.newSubmit.disabled = false;
    renderNewSubmit();
  });
}

function openDeck(path) {
  const value = path.trim();
  if (!value) return;
  localError = null;
  post('api/open', { path: value }).then(() => {
    openPanelRequested = false;
    el.openPath.value = '';
  }).catch(reportError);
}

function toggleNotes() {
  notesVisible = !notesVisible;
  writeStored('pptx-viewer.notes', notesVisible ? '1' : '0');
  if (state) renderNotes();
}

function setHelp(open) {
  if (open) review.close();
  el.help.hidden = !open;
  el.helpButton.setAttribute('aria-expanded', String(open));
}

createDrag({
  list: el.thumbs,
  scroller: el.rail,
  canDrag: (kind) => (kind === 'slide' ? canMove() : canEditPages()),
  onMoveSlide: requestMove,
  onMoveDraft: moveDraftTo,
  announce: say,
});

el.thumbs.addEventListener('click', (event) => {
  const target = event.target instanceof Element ? event.target : null;
  const gap = target?.closest('.gap');
  if (gap) {
    addSlide(JSON.parse(gap.dataset.where));
    return;
  }
  if (target?.closest('.add-tile')) {
    addSlideAtEnd();
    return;
  }
  const remove = target?.closest('.thumb__x');
  if (remove) {
    if (remove.dataset.draft) requestRemoveDraft(remove.dataset.draft);
    else requestRemoveSlide(Number(remove.dataset.n));
    return;
  }
  const button = target?.closest('.thumb');
  if (!button) return;
  if (button.dataset.draft) showDraft(button.dataset.draft);
  else showSlide(Number(button.dataset.n));
});

el.newForm.addEventListener('submit', (event) => {
  event.preventDefault();
  startNewDeck();
});
el.newButton.addEventListener('click', () => {
  openPanelRequested = true;
  if (state) renderOpenPanel();
  el.newPath.focus();
});
el.sketchForm.addEventListener('submit', (event) => {
  event.preventDefault();
  startSketchbook(el.sketchName.value);
});
el.openButton.addEventListener('click', () => {
  openPanelRequested = !openPanelRequested;
  if (state) renderOpenPanel();
  if (!el.openPanel.hidden) el.openPath.focus();
});
el.openForm.addEventListener('submit', (event) => {
  event.preventDefault();
  openDeck(el.openPath.value);
});
el.notesButton.addEventListener('click', toggleNotes);
el.refresh.addEventListener('click', () => post('api/refresh').catch(reportError));
el.external.addEventListener('click', () => post('api/open-external').catch(reportError));
el.helpButton.addEventListener('click', () => setHelp(el.help.hidden));
for (const button of el.langButtons) button.addEventListener('click', () => chooseLanguage(normalizeLanguage(button.dataset.lang)));
el.noticeClose.addEventListener('click', () => {
  if (!state?.notice) return;
  dismissedNotice = state.notice.id;
  writeStored('pptx-viewer.notice', String(dismissedNotice));
  renderNotice();
});
document.addEventListener('click', (event) => {
  if (el.help.hidden || !(event.target instanceof Node)) return;
  if (!el.help.contains(event.target) && !el.helpButton.contains(event.target)) setHelp(false);
});

document.addEventListener('keydown', (event) => {
  if (!state) return;
  if (event.key === 'Escape' && !el.help.hidden) {
    setHelp(false);
    el.helpButton.focus();
    event.preventDefault();
    return;
  }
  if (event.defaultPrevented) return;
  if (event.target instanceof Element && event.target.closest('input, textarea, select')) return;
  if (ink.handleKey(event)) {
    event.preventDefault();
    return;
  }
  if (event.ctrlKey && !event.altKey && !event.metaKey && event.key.toLowerCase() === 'm') {
    if (state.deck) addSlideAfterCurrent();
    event.preventDefault();
    return;
  }
  if (event.altKey && !event.ctrlKey && !event.metaKey && event.key.startsWith('Arrow')) {
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') moveCurrentBy(-1);
    else if (event.key === 'ArrowRight' || event.key === 'ArrowDown') moveCurrentBy(1);
    event.preventDefault();
    return;
  }
  if (event.altKey || event.ctrlKey || event.metaKey) return;
  switch (event.key) {
    case 'ArrowLeft':
    case 'ArrowUp':
    case 'PageUp':
      step(-1);
      break;
    case 'ArrowRight':
    case 'ArrowDown':
    case 'PageDown':
      step(1);
      break;
    case ' ':
      if (event.target !== document.body) return;
      step(1);
      break;
    case 'Home':
      showPage(pageList()[0]);
      break;
    case 'End':
      showPage(pageList().at(-1));
      break;
    case 'Delete':
      if (state.draft) requestRemoveDraft(state.draft);
      else requestRemoveSlide(state.current);
      break;
    case 'n':
    case 'N':
      if (isSketchbook()) return;
      toggleNotes();
      break;
    default:
      return;
  }
  event.preventDefault();
});

function startPolling() {
  setInterval(() => {
    fetch('api/state').then((response) => response.json()).then(applyState).catch(() => {});
  }, 2000);
}

function connect() {
  if (!('EventSource' in window)) return startPolling();
  let failures = 0;
  const source = new EventSource('api/events');
  source.addEventListener('state', (event) => {
    failures = 0;
    applyState(JSON.parse(event.data));
  });
  source.addEventListener('progress', (event) => applyProgress(JSON.parse(event.data)));
  source.addEventListener('error', () => {
    failures += 1;
    if (failures >= 3) {
      source.close();
      startPolling();
    }
  });
}

fetch('api/state')
  .then((response) => response.json())
  .then(applyState)
  .catch(reportError)
  .finally(connect);
