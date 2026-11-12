// Comments about the page on screen, as threads: a comment (the user's, or the agent's), the replies under it, and a resolved state.
// A thread tied to a point or an area is written, read and answered in a bubble beside its marker on the slide. The pane on the right
// lists the threads of the page, is where a thread about the whole page is written and answered, and keeps the resolved threads folded
// away. The server holds the authoritative copy: this module sends every change and reloads the list when the page changes or someone
// else (the agent) changed it. The comments of one page are fetched with the same version number the strokes use.

import { createBubble } from './bubble.js';
import { icon } from './icons.js';
import { t, tn } from './i18n.js';

const AUTOSAVE_MS = 1000;

const kindOf = (anchor) => (!anchor ? 'page' : anchor.w === 0 && anchor.h === 0 ? 'point' : 'area');
const percent = (value) => `${Math.round(value * 1e4) / 100}%`;
const authorOf = (comment) => comment.author ?? 'user';

function parseVersion(version) {
  const cut = String(version ?? '').lastIndexOf(':');
  return cut < 0 ? { key: '', rev: -1 } : { key: version.slice(0, cut), rev: Number(version.slice(cut + 1)) };
}

function element(tag, className = '', text = '') {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text) node.textContent = text;
  return node;
}

function textarea(role, id, { value = '', className, label, placeholder = '', rows }) {
  const area = element('textarea', className);
  area.dataset.role = role;
  area.dataset.id = id;
  area.maxLength = 2000;
  area.value = value;
  area.setAttribute('aria-label', label);
  if (placeholder) area.placeholder = placeholder;
  if (rows) area.rows = rows;
  return area;
}

/** "You" or "Agent", with its icon. */
function who(author) {
  const holder = element('span', `who who--${author}`);
  holder.innerHTML = icon(author === 'agent' ? 'bot' : 'user');
  holder.append(author === 'agent' ? t('comments.agent') : t('comments.you'));
  return holder;
}

function iconButton(action, name, label, tip = label) {
  const button = element('button', 'ibtn thread__btn');
  button.type = 'button';
  button.dataset.action = action;
  button.setAttribute('aria-label', label);
  button.title = tip;
  button.innerHTML = icon(name);
  return button;
}

/**
 * `elements`: `panel`, `list`, `title`, `close`, `add` (starts a comment about the whole page), `layer` (over the slide), `toggle`
 * (header button) and `count` (its badge); `stage`, `box` and `bubble` (`{ root, body }`) for the bubble. `getState()` is the viewer
 * state; `getJson` and `post` talk to the server.
 */
export function createComments({ elements, getState, getJson, post, onError, announce, store = { read: () => null, write: () => {} } }) {
  const { panel, list, title, close, add, compose, layer, toggle, count } = elements;
  const model = {
    open: store.read() === '1',
    scope: '',
    loadedScope: '',
    loading: false,
    retryAt: 0,
    epoch: 0,
    comments: [],
    known: { key: '', rev: -1 },
    // A new comment that is being written: `{ anchor, text }`. Without an anchor it is about the whole page.
    pending: null,
    adding: false,
    // The thread that is open in the bubble, and the one that is highlighted.
    editing: null,
    selected: null,
    // The thread the pointer is over, and the one that has the keyboard focus, in the pane: its place on the slide is shown.
    hover: null,
    focus: null,
    visible: true,
    resolvedOpen: false,
    // The thread that is selected in the pane (by clicking it): it shows its box for a reply. A thread with an unsent reply keeps its box too.
    replying: null,
    // What the user typed in a reply box and has not sent, by comment id, so that it survives a redraw.
    replies: new Map(),
  };
  let signature = '';
  let bubbleSignature = '';
  /** Pending saves of edits, by comment id: an edit is saved a moment after typing stops, so nothing is lost if the view is closed. */
  const timers = new Map();

  const bubble = createBubble({ root: elements.bubble.root, body: elements.bubble.body, stage: elements.stage, box: elements.box });

  const current = () => getState();
  const ready = () => Boolean(current()?.deck) && current().phase === 'ready';
  /** "Comments on slide 3", or "Comments on the new slide" for a new slide. */
  const pageTitle = () => {
    const state = current();
    return state?.draft ? t('comments.pageTitleNew') : t('comments.pageTitleSlide', { n: state?.current ?? '' });
  };
  const query = () => (current().draft ? `draft=${encodeURIComponent(current().draft)}` : `slide=${current().current}`);
  const where = () => (current().draft ? { draft: current().draft } : { slide: current().current });
  const find = (id) => model.comments.find((comment) => comment.id === id);
  /** The open threads are numbered in order, and the agent is told the same numbers; a resolved thread has none. */
  const openThreads = () => model.comments.filter((comment) => !comment.resolved);
  const numberOf = (id) => openThreads().findIndex((comment) => comment.id === id) + 1;
  /** Whether the pane shows the reply box of a thread: it is the selected one, or something is typed in it. */
  const replyShown = (id) => model.replying === id || Boolean(model.replies.get(id));
  /**
   * A thread can be answered when there is someone to answer: the agent started it or has replied in it. The user's own comment, with no
   * word from the agent yet, is only changed, so it has no box for a reply. A resolved thread takes none either.
   */
  const canReply = (comment) => !comment.resolved && (authorOf(comment) === 'agent' || (comment.replies ?? []).some((reply) => authorOf(reply) === 'agent'));

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

  /** Runs a request for the page that is on screen now, and ignores its answer when the page has changed since. */
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

  const replace = (comment) => {
    model.comments = model.comments.map((entry) => (entry.id === comment.id ? comment : entry));
  };

  // ---- what the user does -----------------------------------------------------------------------------------

  /**
   * Starts a new comment: about the whole page (no anchor), or tied to a point or an area. One about the page is written in the composer
   * that opens at the foot of the pane; one tied to a place is written in a bubble beside that place. It is saved when the user presses "Add".
   */
  function startNew(anchor = null) {
    if (!ready()) return;
    model.pending = { anchor, text: model.pending?.text ?? '' };
    model.selected = null;
    model.editing = null;
    if (!anchor) {
      setOpen(true);
      render(true);
      focusRole('new', '');
      return;
    }
    render(true);
    focusRole('new', '');
  }

  function submitNew() {
    const pending = model.pending;
    if (!pending || model.adding) return;
    const text = field('new', '')?.value ?? pending.text;
    if (!text.trim()) return;
    model.adding = true;
    mutate('api/comments/add', { text, anchor: pending.anchor }, ({ comment }) => {
      model.comments = [...model.comments, comment];
      model.pending = null;
      model.selected = comment.id;
      announce(t('comments.added'));
    }).finally(() => {
      model.adding = false;
    });
  }

  function cancelNew() {
    model.pending = null;
    render(true);
  }

  function saveEdit(id, text, { revertEmpty = true } = {}) {
    const old = find(id);
    if (!old || authorOf(old) !== 'user') return;
    if (!text.trim()) {
      if (revertEmpty) render(true);
      return;
    }
    if (text.trim() === old.text) return;
    mutate('api/comments/update', { id, text }, ({ comment }) => replace(comment));
  }

  /** Saves the text of comment `id` a moment after typing stops; `read()` gives the text then, or null when there is nothing to save. */
  function scheduleSave(id, read) {
    cancelSave(id);
    timers.set(id, setTimeout(() => {
      timers.delete(id);
      const text = read();
      if (text !== null) saveEdit(id, text, { revertEmpty: false });
    }, AUTOSAVE_MS));
  }

  function cancelSave(id) {
    clearTimeout(timers.get(id));
    timers.delete(id);
  }

  function submitReply(id) {
    const text = field('reply', id)?.value ?? model.replies.get(id) ?? '';
    if (!text.trim() || !find(id)) return;
    mutate('api/comments/reply', { id, text }, ({ comment }) => {
      replace(comment);
      model.replies.delete(id);
      // The redraw gives the focus back with what was in the box, so the box is emptied first.
      const box = field('reply', id);
      if (box) box.value = '';
      announce(t('comments.replied'));
    }).then(() => focusRole('reply', id));
  }

  /**
   * Selecting a thread in the pane (a click on it): it is highlighted, and when it can be answered (see `canReply`) its box for a reply opens
   * at once and the one of the thread that was selected before closes (unless something is typed in it). `focus` puts the cursor in the box;
   * a click in the text of the user's own comment, to change it, leaves the cursor where it is.
   */
  function selectThread(id, { focus = true } = {}) {
    const comment = find(id);
    if (!comment || comment.resolved) return;
    const replying = canReply(comment) ? id : null;
    const changed = model.replying !== replying || model.selected !== id;
    model.replying = replying;
    model.selected = id;
    if (changed) render(true);
    if (focus && replying) focusRole('reply', id);
  }

  /** Closes the box of the selected thread (the Escape key in an empty box). */
  function stopReply(id) {
    if (model.replying === id) model.replying = null;
    if (model.selected === id) model.selected = null;
    render(true);
  }

  function toggleResolved(id) {
    const comment = find(id);
    if (!comment) return;
    const resolved = !comment.resolved;
    mutate('api/comments/resolve', { id, resolved }, ({ comment: updated }) => {
      replace(updated);
      if (resolved && model.editing === id) model.editing = null;
      announce(resolved ? t('comments.resolvedDone') : t('comments.reopened'));
    });
  }

  function remove(id) {
    if (model.editing === id) model.editing = null;
    mutate('api/comments/remove', { id }, () => {
      model.comments = model.comments.filter((comment) => comment.id !== id);
      model.replies.delete(id);
      if (model.replying === id) model.replying = null;
      if (model.selected === id) model.selected = null;
      announce(t('comments.deleted'));
    });
  }

  /** Opens the bubble of an existing thread tied to a place, to read, answer or change it. An unfinished new comment is not thrown away. */
  function openBubble(id) {
    const comment = find(id);
    if (!comment?.anchor) return;
    if (model.editing === id) {
      closeBubble();
      return;
    }
    if (model.pending?.anchor) {
      if (field('new', '')?.value.trim()) {
        focusRole('new', '');
        return;
      }
      model.pending = null;
    }
    closeBubble({ keepRender: true });
    model.editing = id;
    model.selected = id;
    render(true);
    // The cursor goes to the box for a reply; a thread with none (the user's own comment) puts the focus on the bubble, so that Escape closes it.
    if (canReply(comment)) focusRole('reply', id);
    else bubble.body.focus({ preventScroll: true });
    if (model.open) list.querySelector(`.thread[data-id="${id}"]`)?.scrollIntoView({ block: 'nearest' });
  }

  /** Closes the bubble of a thread, keeping what was typed in it. */
  function closeBubble({ keepRender = false } = {}) {
    const id = model.editing;
    if (!id) return;
    cancelSave(id);
    const root = field('root', id, bubble.body);
    model.editing = null;
    model.selected = null;
    if (root) saveEdit(id, root.value);
    if (!keepRender) render(true);
  }

  /** Brings a thread to the user's attention from the pane: one tied to a place opens in its bubble, one about the page is highlighted. */
  function select(id) {
    const comment = find(id);
    if (!comment) return;
    if (comment.anchor) {
      openBubble(id);
      return;
    }
    model.selected = id;
    model.editing = null;
    model.replying = canReply(comment) ? id : null;
    setOpen(true);
    render(true);
    list.querySelector(`.thread[data-id="${id}"]`)?.scrollIntoView({ block: 'nearest' });
    if (model.replying) focusRole('reply', id);
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

  // ---- finding things in the pane and the bubble ---------------------------------------------------------------

  function field(role, id, root = null) {
    const selector = `textarea[data-role="${role}"][data-id="${id}"]`;
    if (root) return root.querySelector(selector);
    return bubble.body.querySelector(selector) ?? compose.querySelector(selector) ?? list.querySelector(selector);
  }

  function focusRole(role, id) {
    const target = field(role, id);
    if (!target) return;
    target.focus();
    target.setSelectionRange(target.value.length, target.value.length);
  }

  /** What has the keyboard focus, when it is a text box of the pane, the bubble or the composer, so that a redraw can give it back. */
  function captureFocus() {
    const active = document.activeElement;
    if (!(active instanceof HTMLTextAreaElement)) return null;
    const where = bubble.body.contains(active) ? 'bubble' : compose.contains(active) ? 'compose' : list.contains(active) ? 'list' : null;
    if (!where) return null;
    return { where, role: active.dataset.role, id: active.dataset.id ?? '', value: active.value, start: active.selectionStart, end: active.selectionEnd };
  }

  function restoreFocus(saved) {
    if (!saved) return;
    const target = { bubble: bubble.body, compose, list }[saved.where].querySelector(`textarea[data-role="${saved.role}"][data-id="${saved.id}"]`);
    if (!target) return;
    target.value = saved.value;
    target.focus();
    target.setSelectionRange(saved.start, saved.end);
    syncReplyButton(target);
  }

  /** The "Reply" button is enabled when there is something to send. */
  function syncReplyButton(area) {
    if (area.dataset.role !== 'reply') return;
    const button = area.closest('.thread')?.querySelector('[data-action="reply"]');
    if (button) button.disabled = !area.value.trim();
  }

  // ---- drawing the threads ---------------------------------------------------------------------------------------

  function threadHead(comment, { bubbled = false } = {}) {
    const head = element('div', 'thread__head');
    const number = element('span', 'thread__num');
    if (comment.resolved) number.innerHTML = icon('check-circle');
    else number.textContent = String(numberOf(comment.id));
    head.append(number, who(authorOf(comment)), element('span', 'thread__spacer'));
    head.append(comment.resolved
      ? iconButton('resolve', 'undo', t('comments.reopen'), t('comments.reopenTitle'))
      : iconButton('resolve', 'check-circle', t('comments.resolve'), t('comments.resolveTitle')));
    head.append(iconButton('delete', 'trash', t('comments.deleteThis'), t('comments.deleteTitle')));
    if (bubbled) head.append(iconButton('close', 'close', t('comments.closeTitle')));
    return head;
  }

  function threadClasses(comment, extra = '') {
    return `thread${authorOf(comment) === 'agent' ? ' thread--agent' : ''}${comment.resolved ? ' is-resolved' : ''}${comment.id === model.selected || comment.id === model.editing || comment.id === model.replying ? ' is-selected' : ''}${extra}`;
  }

  function repliesList(comment) {
    const replies = element('ol', 'thread__replies');
    for (const reply of comment.replies ?? []) {
      const item = element('li', `reply reply--${authorOf(reply)}`);
      item.append(who(authorOf(reply)), element('p', '', reply.text));
      replies.append(item);
    }
    return replies;
  }

  /** A whole thread, to read, answer and (for the user's own comment) change. */
  function threadView(comment, { bubbled = false } = {}) {
    const article = element('article', threadClasses(comment));
    article.dataset.id = comment.id;
    article.append(threadHead(comment, { bubbled }));
    article.append(authorOf(comment) === 'agent'
      ? element('p', 'thread__text', comment.text)
      : textarea('root', comment.id, { value: comment.text, className: 'thread__text', label: t('comments.item', { n: numberOf(comment.id) || '' }) }));
    if (comment.replies?.length) article.append(repliesList(comment));
    // The bubble has its box for a reply whenever the thread can be answered. In the pane the box is there for the selected thread (see selectThread).
    if (canReply(comment) && (bubbled || replyShown(comment.id))) {
      const row = element('div', 'thread__reply');
      const input = textarea('reply', comment.id, { value: model.replies.get(comment.id) ?? '', className: 'thread__input', label: t('comments.replyLabel'), placeholder: t('comments.replyPlaceholder'), rows: 2 });
      const send = element('button', 'btn btn--primary', t('comments.reply'));
      send.type = 'button';
      send.dataset.action = 'reply';
      send.disabled = !input.value.trim();
      row.append(input, send);
      article.append(row);
    }
    return article;
  }

  /** A short thread, for one that is resolved or open in the bubble: its text and how many replies it has. Clicking it opens (or closes) it in the bubble. */
  function compactView(comment) {
    const article = element('article', threadClasses(comment, ' thread--compact'));
    article.dataset.id = comment.id;
    article.tabIndex = 0;
    article.append(threadHead(comment), element('p', 'thread__preview', comment.text));
    if (comment.replies?.length) article.append(element('p', 'thread__more', tn('comments.replies', comment.replies.length)));
    return article;
  }

  /** The "+" of a comment being written: the same drawing on the slide and in the bubble, whatever the language or the font. */
  function plusMark() {
    const node = element('span', 'plus');
    node.innerHTML = icon('plus');
    return node;
  }

  /** The form of a new comment: in the bubble beside a place, or (`page`) in the composer at the foot of the pane for one about the whole page. */
  function newView({ page = false } = {}) {
    const article = element('article', 'thread thread--new is-selected');
    const head = element('div', 'thread__head');
    const mark = element('span', 'thread__num');
    mark.append(plusMark());
    head.append(mark, page ? element('span', 'thread__title', t('comments.new')) : who('user'), element('span', 'thread__spacer'));
    const area = textarea('new', '', { value: model.pending.text, className: 'thread__text', label: t('comments.new'), placeholder: t('comments.placeholder'), rows: page ? 3 : undefined });
    const actions = element('div', 'thread__actions');
    const cancel = element('button', 'btn', t('comments.cancel'));
    cancel.type = 'button';
    cancel.dataset.action = 'cancel';
    const save = element('button', 'btn btn--primary', t('comments.save'));
    save.type = 'button';
    save.dataset.action = 'save';
    actions.append(cancel, save);
    article.append(head, area, actions);
    return article;
  }

  // ---- drawing the markers, the bubble and the pane -------------------------------------------------------------

  /** `label` is the text in the badge, or `null` for the "+" of a comment being written. */
  function marker(anchor, label, classes) {
    const area = kindOf(anchor) === 'area';
    const node = element('div', `cmark ${area ? 'cmark--area' : 'cmark--point'} ${classes}`.trim());
    node.style.left = percent(anchor.x);
    node.style.top = percent(anchor.y);
    if (area) {
      node.style.width = percent(anchor.w);
      node.style.height = percent(anchor.h);
    }
    const badge = element('button', 'cmark__badge', label ?? '');
    if (label === null) badge.append(plusMark());
    badge.type = 'button';
    node.append(badge);
    return { node, badge };
  }

  /** The thread whose place is shown on the slide: the one under the pointer, else the one with the focus. */
  const highlighted = () => model.hover ?? model.focus;

  /** Shows which thread of the pane the pointer or the focus is on, without redrawing the markers. */
  function paintHighlight() {
    const id = highlighted();
    for (const node of layer.children) node.classList.toggle('is-hover', node.dataset.id === id);
  }

  function renderMarkers(on) {
    layer.replaceChildren();
    layer.hidden = !on || !model.visible;
    if (layer.hidden) return;
    for (const comment of model.comments) {
      if (comment.resolved && comment.id !== model.editing) continue;
      const agent = authorOf(comment) === 'agent' ? 'cmark--agent' : '';
      if (!comment.anchor) {
        // A thread about the whole page has no place: while it is pointed at, the whole slide is shaded.
        const shade = element('div', `cmark cmark--page ${agent}`.trim());
        shade.dataset.id = comment.id;
        layer.append(shade);
        continue;
      }
      const classes = [agent, comment.id === model.selected || comment.id === model.editing ? 'is-selected' : ''].join(' ').trim();
      const { node, badge } = marker(comment.anchor, comment.resolved ? '✓' : String(numberOf(comment.id)), classes);
      node.dataset.id = comment.id;
      badge.dataset.id = comment.id;
      badge.setAttribute('aria-label', t('comments.openMarker', { n: numberOf(comment.id) || '' }));
      badge.title = comment.text.length > 80 ? `${comment.text.slice(0, 80)}…` : comment.text;
      layer.append(node);
    }
    if (model.pending?.anchor) {
      const { node, badge } = marker(model.pending.anchor, null, 'cmark--pending');
      badge.setAttribute('aria-label', t('comments.new'));
      layer.append(node);
    }
    paintHighlight();
  }

  /** What the bubble should show now, or null when it should not be there. */
  function bubbleTarget() {
    if (model.pending?.anchor) return { key: 'new', anchor: model.pending.anchor, author: 'user', signature: t('comments.new'), build: newView };
    const comment = model.editing ? find(model.editing) : null;
    if (!comment?.anchor) return null;
    return {
      key: `thread:${comment.id}`,
      anchor: comment.anchor,
      author: authorOf(comment),
      signature: JSON.stringify([comment.text, comment.updatedAt, comment.resolved, (comment.replies ?? []).map((reply) => reply.id), numberOf(comment.id), t('comments.you')]),
      build: () => threadView(comment, { bubbled: true }),
    };
  }

  function renderBubble(on, saved, force) {
    const target = on && model.visible ? bubbleTarget() : null;
    if (!target) {
      // A thread that is gone (deleted elsewhere) cannot stay open; an unfinished new comment is kept while the marks are hidden.
      if (model.editing && !find(model.editing)) model.editing = null;
      bubble.hide();
      bubbleSignature = '';
      return;
    }
    bubble.show({ key: target.key, anchor: target.anchor, author: target.author });
    const next = `${target.key}|${target.signature}`;
    if (next !== bubbleSignature || force) {
      bubble.body.replaceChildren(target.build());
      bubbleSignature = next;
      if (saved?.where === 'bubble') restoreFocus(saved);
      bubble.place();
    }
  }

  /** The box for a comment about the whole page opens at the foot of the pane, under the list, so that it is never out of sight. */
  function renderCompose(saved, force) {
    const open = Boolean(model.pending && !model.pending.anchor);
    add.hidden = open;
    if (!open) {
      if (compose.hidden) return;
      const hadFocus = compose.contains(document.activeElement);
      compose.replaceChildren();
      compose.hidden = true;
      if (hadFocus) add.focus();
      return;
    }
    // Once it is open it is left alone, so that what is being typed is not wiped by a redraw; `force` builds it again (a new language).
    if (!compose.hidden && !force) return;
    compose.hidden = false;
    compose.replaceChildren(newView({ page: true }));
    if (saved?.where === 'compose') restoreFocus(saved);
  }

  /** The pane shows every open thread whole, with its replies and a box to answer, so that nothing needs the bubble to be answered. */
  function renderList(saved) {
    const items = [];
    for (const comment of openThreads()) {
      const item = element('li');
      // The thread that is open in the bubble is short here, so that there is only one place to type in.
      item.append(comment.id === model.editing ? compactView(comment) : threadView(comment));
      items.push(item);
    }
    const done = model.comments.filter((comment) => comment.resolved);
    if (done.length) {
      const item = element('li');
      const details = element('details', 'resolved');
      details.open = model.resolvedOpen;
      details.addEventListener('toggle', () => {
        model.resolvedOpen = details.open;
      });
      const summary = element('summary', '', t('comments.resolvedCount', { count: done.length }));
      const resolved = element('ul', 'resolved__list');
      for (const comment of done) {
        const row = element('li');
        row.append(compactView(comment));
        resolved.append(row);
      }
      details.append(summary, resolved);
      item.append(details);
      items.push(item);
    }
    list.replaceChildren(...items);
    // A focused element that the redraw removed sends no "focusout", so the pane forgets its focus here (a restored focus sets it again).
    if (!list.contains(document.activeElement)) model.focus = null;
    if (saved?.where === 'list') restoreFocus(saved);
  }

  /**
   * Draws everything that depends on the comments. While the user is typing, it leaves the pane and the bubble alone unless `force` says
   * something changed that must show (so a comment changed elsewhere does not wipe what is being typed).
   */
  function render(force = false) {
    const state = current();
    const on = ready();
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

    const saved = captureFocus();
    renderBubble(on, saved, force);
    renderCompose(saved, force);

    // What is typed in the bubble or the composer is not part of the signature: they keep it themselves, and the pane does not show it.
    const pendingKey = model.pending && (model.pending.anchor ?? 'page');
    const threads = model.comments.map((comment) => [comment.id, comment.text, comment.updatedAt, comment.resolved, authorOf(comment), (comment.replies ?? []).length]);
    const next = JSON.stringify([model.scope, threads, model.selected, pendingKey, model.editing, model.visible, on, model.resolvedOpen, model.replying]);
    if (next === signature && !force) return;
    signature = next;
    renderList(saved);
    renderMarkers(on);
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
        model.editing = null;
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
      model.editing = null;
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
    if (badge) openBubble(badge.dataset.id);
  });

  // Clicking away from the bubble closes the bubble of a thread, and drops a new comment that is still empty. A new comment with text in
  // it stays until it is added or cancelled, so nothing typed is lost by a stray click.
  document.addEventListener('pointerdown', (event) => {
    if (!bubble.isOpen()) return;
    const target = event.target instanceof Element ? event.target : null;
    if (!target || bubble.contains(target) || target.closest('.cmark__badge')) return;
    if (model.editing) closeBubble();
    else if (model.pending?.anchor && !field('new', '')?.value.trim()) cancelNew();
  }, true);

  /** The same events in the pane and in the bubble: both hold threads and the new-comment form. */
  function wire(container) {
    container.addEventListener('click', (event) => {
      const target = event.target instanceof Element ? event.target : null;
      const action = target?.closest('[data-action]')?.dataset.action;
      const thread = target?.closest('.thread');
      const id = thread?.dataset.id;
      if (action === 'save') submitNew();
      else if (action === 'cancel') cancelNew();
      else if (action === 'reply' && id) submitReply(id);
      else if (action === 'resolve' && id) toggleResolved(id);
      else if (action === 'delete' && id) remove(id);
      else if (action === 'close') closeBubble();
      else if (!action && id && thread.classList.contains('thread--compact')) select(id);
      // A click on a thread in the pane selects it, and so opens its box for a reply (not on a button or on the heading of a folded list).
      else if (!action && id && container === list && !thread.classList.contains('thread--new') && !target.closest('button, summary')) {
        selectThread(id, { focus: !target.closest('textarea') });
      }
    });

    container.addEventListener('keydown', (event) => {
      const area = event.target;
      if (!(area instanceof HTMLTextAreaElement)) {
        if (event.key === 'Enter' && event.target instanceof Element && event.target.matches('.thread--compact') && event.target.dataset.id) select(event.target.dataset.id);
        // A bubble with no text box in focus (the user's own comment, read before it is changed) closes with Escape too.
        else if (event.key === 'Escape' && container === bubble.body) {
          event.preventDefault();
          closeBubble();
        }
        return;
      }
      const { role, id } = area.dataset;
      if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        if (role === 'new') submitNew();
        else if (role === 'reply') submitReply(id);
        else area.blur();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        if (role === 'new') cancelNew();
        else if (container === bubble.body) closeBubble();
        else if (role === 'reply' && !area.value) stopReply(id);
        else area.blur();
      }
    });

    container.addEventListener('input', (event) => {
      const area = event.target;
      if (!(area instanceof HTMLTextAreaElement)) return;
      const { role, id } = area.dataset;
      if (role === 'root') scheduleSave(id, () => (area.isConnected ? area.value : null));
      else if (role === 'new') {
        if (model.pending) model.pending.text = area.value;
      } else if (role === 'reply') {
        if (area.value) model.replies.set(id, area.value);
        else model.replies.delete(id);
        syncReplyButton(area);
      }
    });

    // An edit is saved when the cursor leaves the text, and also a moment after typing stops.
    container.addEventListener('focusout', (event) => {
      const area = event.target;
      if (!(area instanceof HTMLTextAreaElement) || area.dataset.role !== 'root') return;
      cancelSave(area.dataset.id);
      saveEdit(area.dataset.id, area.value);
    });
  }

  wire(list);
  wire(bubble.body);
  wire(compose);

  // Pointing at a thread in the pane (or tabbing into it) shows its place on the slide: its marker grows and blinks, and a thread about the
  // whole page shades the whole slide. "pointerleave" is not used: it never comes when the redraw replaced the thread under the pointer.
  const threadIdOf = (target) => (target instanceof Element ? target.closest('.thread[data-id]')?.dataset.id ?? null : null);
  document.addEventListener('pointerover', (event) => {
    const id = event.target instanceof Node && list.contains(event.target) ? threadIdOf(event.target) : null;
    if (id === model.hover) return;
    model.hover = id;
    paintHighlight();
  });
  document.addEventListener('pointerout', (event) => {
    if (event.relatedTarget || !model.hover) return;
    model.hover = null;
    paintHighlight();
  });
  list.addEventListener('focusin', (event) => {
    model.focus = threadIdOf(event.target);
    paintHighlight();
  });
  list.addEventListener('focusout', (event) => {
    if (threadIdOf(event.relatedTarget) === model.focus) return;
    model.focus = null;
    paintHighlight();
  });

  return { sync, startNew, setMarksVisible, setOpen, isOpen: () => model.open, refresh: () => render(true) };
}
