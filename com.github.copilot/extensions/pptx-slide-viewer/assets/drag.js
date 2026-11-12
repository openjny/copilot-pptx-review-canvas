// Drag-and-drop in the strip of slide thumbnails (laid out left to right). A slide that is dragged moves in the file;
// a new (not yet saved) slide moves among the others. This module only turns pointer movements into "the dragged one goes
// to this place"; the page decides what that does (`onMoveSlide`, `onMoveDraft`) and when dragging is allowed (`canDrag`).

import { t } from './i18n.js';

const SLOP = 6;
const EDGE = 64;
const MAX_SCROLL = 22;

export function createDrag({ list, scroller, canDrag, onMoveSlide, onMoveDraft, announce }) {
  let drag = null;
  let line = null;
  let ghost = null;
  let frame = 0;
  let suppressClick = false;

  const pages = () => [...list.querySelectorAll(':scope > li.page')];

  /** Where the pointer would drop the dragged page: a slot among the slides (for a slide) or among all the other pages (for a new slide). */
  function slotAt(x) {
    const mid = (item) => {
      const box = item.getBoundingClientRect();
      return box.left + box.width / 2;
    };
    if (drag.kind === 'slide') {
      const slides = pages().filter((item) => item.dataset.kind === 'slide');
      const index = slides.findIndex((item) => x < mid(item));
      return index < 0 ? slides.length + 1 : index + 1;
    }
    return pages().filter((item) => item !== drag.item && x >= mid(item)).length;
  }

  const isNoop = (slot) => (drag.kind === 'slide' ? slot === drag.from || slot === drag.from + 1 : slot === drag.index);

  /** The page the drop line sits in front of, or null for "after the last page". */
  function beforeItem(slot) {
    if (drag.kind === 'slide') {
      const slides = pages().filter((item) => item.dataset.kind === 'slide');
      return slot <= slides.length ? slides[slot - 1] : null;
    }
    return pages().filter((item) => item !== drag.item)[slot] ?? null;
  }

  function placeLine(slot) {
    line.hidden = isNoop(slot);
    if (line.hidden) return;
    const base = list.getBoundingClientRect();
    const before = beforeItem(slot);
    const last = pages().at(-1);
    const x = before ? before.getBoundingClientRect().left - base.left - 7 : last.getBoundingClientRect().right - base.left + 7;
    line.style.left = `${Math.round(x)}px`;
  }

  function update() {
    drag.slot = slotAt(drag.x);
    placeLine(drag.slot);
  }

  /** A see-through copy of the thumbnail that follows the pointer, so it is clear which slide is being moved. */
  function createGhost() {
    const source = drag.item.querySelector('button.thumb');
    const box = source.getBoundingClientRect();
    ghost = document.createElement('div');
    ghost.className = 'drag-ghost';
    ghost.setAttribute('aria-hidden', 'true');
    ghost.style.width = `${box.width}px`;
    ghost.style.setProperty('--ratio', getComputedStyle(list).getPropertyValue('--ratio'));
    const copy = source.cloneNode(true);
    copy.removeAttribute('aria-label');
    copy.tabIndex = -1;
    ghost.append(copy);
    document.body.append(ghost);
    drag.offsetX = drag.startX - box.left;
    drag.offsetY = drag.startY - box.top;
    moveGhost();
  }

  function moveGhost() {
    if (ghost) ghost.style.transform = `translate3d(${Math.round(drag.x - drag.offsetX)}px, ${Math.round(drag.y - drag.offsetY)}px, 0)`;
  }

  /** Scrolls the strip while the pointer is near its left or right edge, so a long deck can be dragged from end to end. */
  function autoScroll() {
    frame = requestAnimationFrame(() => {
      if (!drag?.active) return;
      const box = scroller.getBoundingClientRect();
      const left = box.left + EDGE - drag.x;
      const right = drag.x - (box.right - EDGE);
      if (left > 0) scroller.scrollLeft -= Math.ceil(MAX_SCROLL * Math.min(1, left / EDGE));
      else if (right > 0) scroller.scrollLeft += Math.ceil(MAX_SCROLL * Math.min(1, right / EDGE));
      update();
      autoScroll();
    });
  }

  function begin() {
    drag.active = true;
    drag.item.classList.add('is-dragging');
    document.body.classList.add('is-reordering');
    line = document.createElement('div');
    line.className = 'drop-line';
    line.setAttribute('aria-hidden', 'true');
    list.append(line);
    createGhost();
    announce(drag.kind === 'slide' ? t('drag.startSlide', { n: drag.from }) : t('drag.startDraft'));
    autoScroll();
  }

  function end() {
    const finished = drag;
    drag = null;
    cancelAnimationFrame(frame);
    line?.remove();
    line = null;
    ghost?.remove();
    ghost = null;
    finished?.item.classList.remove('is-dragging');
    document.body.classList.remove('is-reordering');
    window.removeEventListener('pointermove', onPointerMove);
    window.removeEventListener('pointerup', onPointerUp);
    window.removeEventListener('pointercancel', onPointerCancel);
    window.removeEventListener('keydown', onKey, true);
    return finished;
  }

  function swallowNextClick() {
    suppressClick = true;
    setTimeout(() => {
      suppressClick = false;
    }, 0);
  }

  function onPointerMove(event) {
    if (!drag || event.pointerId !== drag.pointerId) return;
    drag.x = event.clientX;
    drag.y = event.clientY;
    if (!drag.active) {
      if (Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < SLOP) return;
      begin();
    }
    moveGhost();
    update();
  }

  function onPointerUp(event) {
    if (!drag || event.pointerId !== drag.pointerId) return;
    const finished = end();
    if (!finished.active) return;
    swallowNextClick();
    if (isNoopFor(finished)) {
      announce(t('drag.unchanged'));
      return;
    }
    if (finished.kind === 'slide') onMoveSlide(finished.from, finished.slot > finished.from ? finished.slot - 1 : finished.slot);
    else onMoveDraft(finished.id, finished.slot);
  }

  const isNoopFor = (finished) => (finished.kind === 'slide' ? finished.slot === finished.from || finished.slot === finished.from + 1 : finished.slot === finished.index);

  function onPointerCancel(event) {
    if (!drag || event.pointerId !== drag.pointerId) return;
    if (end().active) announce(t('drag.canceled'));
  }

  function onKey(event) {
    if (event.key !== 'Escape' || !drag) return;
    event.preventDefault();
    event.stopPropagation();
    if (end().active) {
      swallowNextClick();
      announce(t('drag.canceled'));
    }
  }

  list.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || drag) return;
    const button = event.target instanceof Element ? event.target.closest('button.thumb') : null;
    const item = button?.closest('li.page');
    if (!item || !canDrag(item.dataset.kind)) return;
    const kind = item.dataset.kind;
    drag = {
      kind,
      item,
      from: Number(item.dataset.n),
      id: item.dataset.draft,
      index: pages().indexOf(item),
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
      offsetX: 0,
      offsetY: 0,
      active: false,
      slot: 0,
    };
    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    window.addEventListener('pointercancel', onPointerCancel);
    window.addEventListener('keydown', onKey, true);
  });

  // The thumbnails contain images, which the browser would otherwise start dragging as a picture.
  list.addEventListener('dragstart', (event) => event.preventDefault());

  // A drag ends with a click on the thumbnail it started from; that click must not also select the slide.
  list.addEventListener('click', (event) => {
    if (!suppressClick) return;
    event.stopPropagation();
    event.preventDefault();
  }, true);
}
