// Comments about the page on screen. The pane on the right lists them and is where they are written; the markers on the slide show where
// the ones tied to a point or an area are. The server holds the authoritative copy: this module sends every change and reloads the list
// when the page changes or someone else changed it (the comments of one page are fetched with the same version number the strokes use).

import { icon } from './icons.js';
import { t } from './i18n.js';

const AUTOSAVE_MS = 1000;

const kindOf = (anchor) => (!anchor ? 'page' : anchor.w === 0 && anchor.h === 0 ? 'point' : 'area');
const kindLabel = (anchor) => ({ page: t('comments.kind.page'), point: t('comments.kind.point'), area: t('comments.kind.area') })[kindOf(anchor)];
const percent = (value) => `${Math.round(value * 1e4) / 100}%`;

function parseVersion(version) {
  const cut = String(version ?? '').lastIndexOf(':');
  return cut < 0 ? { key: '', rev: -1 } : { key: version.slice(0, cut), rev: Number(version.slice(cut + 1)) };
}

/**
 * `elements`: `panel`, `list`, `title`, `close`, `add` (adds a comment about the whole page), `layer` (over the slide), `toggle` (header
 * button) and `count` (its badge). `getState()` is the viewer state; `getJson` and `post` talk to the server.
 */
export function createComments({ elements, getState, getJson, post, onError, announce, store = { read: () => null, write: () => {} } }) {
  const { panel, list, title, close, add, layer, toggle, count } = elements;
  const model = {
    open: store.read() === '1',
    scope: '',
    loadedScope: '',
    loading: false,
    retryAt: 0,
    epoch: 0,
    comments: [],
    known: { key: '', rev: -1 },
    pending: null,
    selected: null,
    visible: true,
  };
  let signature = '';

  const current = () => getState();
  /** "Comments on slide 3", or "Comments on the new slide" for a new slide. */
  const pageTitle = () => {
    const state = current();
    return state?.draft ? t('comments.pageTitleNew') : t('comments.pageTitleSlide', { n: state?.current ?? '' });
  };
  const query = () => (current().draft ? `draft=${encodeURIComponent(current().draft)}` : `slide=${current().current}`);
  const where = () => (current().draft ? { draft: current().draft } : { slide: current().current });

  // ---- talking to the server --------------------------------------------------------------------------------

  function adopt(version) {
    const next = parseVersion(version);
    if (next.key !== model.known.key || next.rev > model.known.rev) model.known = next;
  }

  function load() {
    const { scope, epoch } = model;
    model.loading = true;
    getJson(`api/comments?${query()}`).then((data) => {
      if (scope !== model.scope || epoch !== model.epoch) return;
      model.comments = data.comments;
      model.loadedScope = scope;
      adopt(data.version);
      render();
    }).catch(() => {
      model.retryAt = Date.now() + 1000;
      setTimeout(() => sync(current()), 1000);
    }).finally(() => {
      model.loading = false;
      sync(current());
    });
  }

  /** Runs `request` for the page that is on screen now, and ignores its answer when the page has changed since. */
  function mutate(path, body, apply) {
    const { scope } = model;
    return post(path, { ...where(), ...body }).then((data) => {
      adopt(data.version);
      if (scope === model.scope) {
        model.epoch += 1;
        apply(data);
        render();
      }
      return data;
    }).catch((error) => {
      onError(error);
      model.loadedScope = '';
      sync(current());
    });
  }

  // ---- what the user does -----------------------------------------------------------------------------------

  /** Starts a new comment: about the whole page (no anchor), or tied to a point or an area. It is saved when the user presses "Add". */
  function startNew(anchor = null) {
    if (!current()?.deck || current().phase !== 'ready') return;
    model.pending = { anchor, text: model.pending?.text ?? '' };
    model.selected = null;
    setOpen(true);
    render();
    list.querySelector('.comment--new .comment__text')?.focus();
  }

  function submitNew() {
    const pending = model.pending;
    const text = list.querySelector('.comment--new .comment__text')?.value ?? pending?.text ?? '';
    if (!pending || !text.trim()) return;
    mutate('api/comments/add', { text, anchor: pending.anchor }, ({ comment }) => {
      model.comments = [...model.comments, comment];
      model.pending = null;
      model.selected = comment.id;
      announce(t('comments.added'));
    });
  }

  function cancelNew() {
    model.pending = null;
    render();
  }

  function saveEdit(id, text, { revertEmpty = true } = {}) {
    const old = model.comments.find((comment) => comment.id === id);
    if (!old) return;
    if (!text.trim()) {
      if (revertEmpty) render(true);
      return;
    }
    if (text.trim() === old.text) return;
    mutate('api/comments/update', { id, text }, ({ comment }) => {
      model.comments = model.comments.map((entry) => (entry.id === id ? comment : entry));
    });
  }

  function remove(id) {
    mutate('api/comments/remove', { id }, () => {
      model.comments = model.comments.filter((comment) => comment.id !== id);
      if (model.selected === id) model.selected = null;
      announce(t('comments.deleted'));
    });
  }

  /** Brings a comment to the user's attention: opens the pane, highlights it and its marker, and puts the cursor in its text. */
  function select(id, { focus = true } = {}) {
    model.selected = id;
    setOpen(true);
    render(true);
    const text = list.querySelector(`.comment[data-id="${id}"] .comment__text`);
    text?.closest('.comment')?.scrollIntoView({ block: 'nearest' });
    if (focus) text?.focus();
  }

  function setOpen(open) {
    if (model.open === open) return;
    model.open = open;
    store.write(open ? '1' : '0');
    render();
  }

  function setMarksVisible(visible) {
    model.visible = visible;
    render(true);
  }

  // ---- drawing the pane and the markers ---------------------------------------------------------------------

  function marker(anchor, label, classes) {
    const area = kindOf(anchor) === 'area';
    const element = document.createElement('div');
    element.className = `cmark ${area ? 'cmark--area' : 'cmark--point'} ${classes}`.trim();
    element.style.left = percent(anchor.x);
    element.style.top = percent(anchor.y);
    if (area) {
      element.style.width = percent(anchor.w);
      element.style.height = percent(anchor.h);
    }
    const badge = document.createElement('button');
    badge.type = 'button';
    badge.className = 'cmark__badge';
    badge.textContent = label;
    element.append(badge);
    return { element, badge };
  }

  function renderMarkers(on) {
    layer.replaceChildren();
    layer.hidden = !on || !model.visible;
    if (layer.hidden) return;
    model.comments.forEach((comment, index) => {
      if (!comment.anchor) return;
      const { element, badge } = marker(comment.anchor, String(index + 1), comment.id === model.selected ? 'is-selected' : '');
      badge.dataset.id = comment.id;
      badge.setAttribute('aria-label', t('comments.openMarker', { n: index + 1, kind: kindLabel(comment.anchor) }));
      badge.title = comment.text.length > 80 ? `${comment.text.slice(0, 80)}…` : comment.text;
      layer.append(element);
    });
    if (model.pending?.anchor) layer.append(marker(model.pending.anchor, t('comments.newMark'), 'cmark--pending').element);
  }

  function commentItem(comment, index) {
    const item = document.createElement('li');
    item.className = `comment${comment.id === model.selected ? ' is-selected' : ''}`;
    item.dataset.id = comment.id;
    const head = document.createElement('div');
    head.className = 'comment__head';
    const number = document.createElement('span');
    number.className = 'comment__num';
    number.textContent = String(index + 1);
    const kind = document.createElement('span');
    kind.className = 'comment__kind';
    kind.textContent = kindLabel(comment.anchor);
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'ibtn comment__del';
    del.dataset.action = 'delete';
    del.setAttribute('aria-label', t('comments.delete', { n: index + 1 }));
    del.title = t('comments.deleteTitle');
    del.innerHTML = icon('trash');
    head.append(number, kind, del);
    const text = document.createElement('textarea');
    text.className = 'comment__text';
    text.value = comment.text;
    text.maxLength = 2000;
    text.setAttribute('aria-label', t('comments.item', { n: index + 1 }));
    item.append(head, text);
    return item;
  }

  function pendingItem() {
    const item = document.createElement('li');
    item.className = 'comment comment--new is-selected';
    const head = document.createElement('div');
    head.className = 'comment__head';
    const number = document.createElement('span');
    number.className = 'comment__num';
    number.textContent = t('comments.newMark');
    const kind = document.createElement('span');
    kind.className = 'comment__kind';
    kind.textContent = t('comments.newKind', { kind: kindLabel(model.pending.anchor) });
    head.append(number, kind);
    const text = document.createElement('textarea');
    text.className = 'comment__text';
    text.value = model.pending.text;
    text.maxLength = 2000;
    text.placeholder = t('comments.placeholder');
    text.setAttribute('aria-label', t('comments.new'));
    const actions = document.createElement('div');
    actions.className = 'comment__actions';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'btn';
    cancel.dataset.action = 'cancel';
    cancel.textContent = t('comments.cancel');
    const save = document.createElement('button');
    save.type = 'button';
    save.className = 'btn btn--primary';
    save.dataset.action = 'save';
    save.textContent = t('comments.save');
    actions.append(cancel, save);
    item.append(head, text, actions);
    return item;
  }

  /**
   * Draws everything that depends on the comments. While the user is typing in the pane, it leaves the pane alone unless `force` says
   * something changed that must show (so a comment changed elsewhere does not wipe what is being typed).
   */
  function render(force = false) {
    const state = current();
    const on = Boolean(state?.deck) && state.phase === 'ready';
    const here = state?.draft
      ? state.annotations?.drafts?.find((entry) => entry.id === state.draft)?.comments ?? 0
      : state?.annotations?.commentCounts?.[state?.current] ?? 0;
    count.textContent = String(here);
    count.hidden = here === 0;
    toggle.setAttribute('aria-pressed', String(model.open));
    toggle.title = model.open ? t('comments.close') : here ? (state.draft ? t('comments.showNew', { count: here }) : t('comments.showSlide', { n: state.current, count: here })) : t('comments.show');
    panel.hidden = !model.open || !state?.deck;
    title.textContent = state?.deck ? pageTitle() : t('comments.title');
    add.disabled = !on;

    const next = JSON.stringify([model.scope, model.comments.map((comment) => [comment.id, comment.text, comment.updatedAt]), model.selected, model.pending && [model.pending.anchor, model.pending.text], model.visible, on]);
    if (next === signature && !force) return;
    signature = next;
    const typing = panel.contains(document.activeElement) && document.activeElement.tagName === 'TEXTAREA';
    const focus = typing ? { id: document.activeElement.closest('.comment')?.dataset.id ?? 'new', value: document.activeElement.value, start: document.activeElement.selectionStart, end: document.activeElement.selectionEnd } : null;
    list.replaceChildren(...model.comments.map(commentItem), ...(model.pending ? [pendingItem()] : []));
    renderMarkers(on);
    if (focus) {
      const target = list.querySelector(focus.id === 'new' ? '.comment--new .comment__text' : `.comment[data-id="${focus.id}"] .comment__text`);
      if (target) {
        target.value = focus.value;
        target.focus();
        target.setSelectionRange(focus.start, focus.end);
      }
    }
  }

  // ---- the page changed, or someone else changed the comments -------------------------------------------------

  /** Brings the comments in line with the page on screen. Called whenever the viewer state changes. */
  function sync(state) {
    if (state?.deck) {
      const scope = `${state.deck.key}:${state.draft ? `d${state.draft}` : `s${state.current}`}`;
      if (scope !== model.scope) {
        model.scope = scope;
        model.loadedScope = '';
        model.comments = [];
        model.pending = null;
        model.selected = null;
        model.epoch += 1;
      }
      if (state.phase === 'ready' && !model.loading && Date.now() >= model.retryAt) {
        const server = parseVersion(state.annotations?.version);
        const stale = Boolean(server.key) && server.key === model.known.key && server.rev > model.known.rev;
        if (model.loadedScope !== scope || stale) load();
      }
    } else {
      model.scope = '';
      model.loadedScope = '';
      model.comments = [];
      model.pending = null;
      model.selected = null;
    }
    render();
  }

  // ---- events -------------------------------------------------------------------------------------------------

  toggle.addEventListener('click', () => setOpen(!model.open));
  close.addEventListener('click', () => {
    setOpen(false);
    toggle.focus();
  });
  add.addEventListener('click', () => startNew(null));

  layer.addEventListener('click', (event) => {
    const badge = event.target instanceof Element ? event.target.closest('.cmark__badge[data-id]') : null;
    if (badge) select(badge.dataset.id);
  });

  list.addEventListener('click', (event) => {
    const target = event.target instanceof Element ? event.target : null;
    const action = target?.closest('[data-action]')?.dataset.action;
    const item = target?.closest('.comment');
    if (action === 'save') submitNew();
    else if (action === 'cancel') cancelNew();
    else if (action === 'delete' && item?.dataset.id) remove(item.dataset.id);
    else if (item?.dataset.id && target?.closest('.comment__head')) select(item.dataset.id);
  });

  list.addEventListener('focusin', (event) => {
    const item = event.target instanceof Element ? event.target.closest('.comment[data-id]') : null;
    if (item && model.selected !== item.dataset.id) {
      model.selected = item.dataset.id;
      renderMarkers(Boolean(current()?.deck) && current().phase === 'ready');
      for (const entry of list.querySelectorAll('.comment')) entry.classList.toggle('is-selected', entry === item || entry.classList.contains('comment--new'));
    }
  });

  // An edit is saved when the cursor leaves the text, and also a moment after typing stops, so nothing is lost if the view is closed.
  const timers = new Map();

  list.addEventListener('input', (event) => {
    const area = event.target;
    const id = area instanceof HTMLTextAreaElement ? area.closest('.comment')?.dataset.id : null;
    if (!id) return;
    clearTimeout(timers.get(id));
    timers.set(id, setTimeout(() => {
      timers.delete(id);
      if (area.isConnected) saveEdit(id, area.value, { revertEmpty: false });
    }, AUTOSAVE_MS));
  });

  list.addEventListener('focusout', (event) => {
    const area = event.target;
    if (!(area instanceof HTMLTextAreaElement)) return;
    const item = area.closest('.comment');
    if (item?.classList.contains('comment--new')) {
      if (model.pending) model.pending.text = area.value;
    } else if (item?.dataset.id) {
      clearTimeout(timers.get(item.dataset.id));
      timers.delete(item.dataset.id);
      saveEdit(item.dataset.id, area.value);
    }
  });

  list.addEventListener('keydown', (event) => {
    const area = event.target;
    if (!(area instanceof HTMLTextAreaElement)) return;
    const isNew = area.closest('.comment')?.classList.contains('comment--new');
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      if (isNew) submitNew();
      else area.blur();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      if (isNew) cancelNew();
      else {
        const old = model.comments.find((comment) => comment.id === area.closest('.comment')?.dataset.id);
        if (old) area.value = old.text;
        area.blur();
      }
    }
  });

  return { sync, startNew, setMarksVisible, setOpen, isOpen: () => model.open, refresh: () => render(true) };
}
