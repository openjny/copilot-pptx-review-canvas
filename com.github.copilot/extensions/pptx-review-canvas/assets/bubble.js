// The bubble beside a spot on the slide, in which a comment is written and a thread is read and answered. It only places itself: what is
// inside (`body`) is built by comments.js, which also decides when the bubble opens and closes.

import { placeBubble } from './place.js';

/** A point is drawn as a marker about this far from its centre, so the bubble keeps clear of it. */
const POINT_RADIUS = 14;

/**
 * `root` is the bubble and `body` the element inside it that holds the content. `stage` is the positioned element the bubble lives in,
 * and `box` the slide inside it.
 */
export function createBubble({ root, body, stage, box }) {
  let shown = null;

  function place() {
    if (!shown) return;
    const stageRect = stage.getBoundingClientRect();
    const boxRect = box.getBoundingClientRect();
    if (boxRect.width === 0 || root.offsetWidth === 0) return;
    const { x, y, w, h } = shown.anchor;
    const pad = w === 0 && h === 0 ? POINT_RADIUS : 0;
    const offsetX = boxRect.left - stageRect.left;
    const offsetY = boxRect.top - stageRect.top;
    const anchor = {
      left: offsetX + x * boxRect.width - pad,
      top: offsetY + y * boxRect.height - pad,
      right: offsetX + (x + w) * boxRect.width + pad,
      bottom: offsetY + (y + h) * boxRect.height + pad,
    };
    const spot = placeBubble({
      anchor,
      size: { width: root.offsetWidth, height: root.offsetHeight },
      bounds: { left: 0, top: 0, right: stageRect.width, bottom: stageRect.height },
    });
    root.style.left = `${Math.round(spot.left)}px`;
    root.style.top = `${Math.round(spot.top)}px`;
    root.style.setProperty('--tail', `${Math.round(spot.tail)}px`);
    root.dataset.side = spot.side;
  }

  /** Shows the bubble for `key` at `anchor` (`{ x, y, w, h }` in fractions of the slide), in the colour of `author`. Returns true when it was not open for that key. */
  function show({ key, anchor, author = 'user' }) {
    const opened = shown?.key !== key;
    shown = { key, anchor };
    root.hidden = false;
    root.dataset.author = author;
    place();
    return opened;
  }

  function hide() {
    if (!shown) return;
    shown = null;
    root.hidden = true;
    body.replaceChildren();
  }

  const observer = new ResizeObserver(place);
  observer.observe(stage);
  observer.observe(box);
  observer.observe(root);

  return {
    body,
    show,
    hide,
    place,
    isOpen: () => shown !== null,
    key: () => shown?.key ?? null,
    contains: (node) => root.contains(node),
  };
}
