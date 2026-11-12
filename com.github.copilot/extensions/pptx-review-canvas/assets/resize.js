// The drag handle on an edge of a panel: the left edge of the comments pane (its width) and the top edge of the strip of thumbnails (its
// height). It only works out a size: `onCommit` is called with the final size (or null for the default) when the user lets go, so the page
// can save it. The size is a CSS variable (`variable`) on `layout`; nothing here sets an inline style attribute, which the page's content
// security policy would block.

export const COMMENTS_WIDTH = { min: 240, max: 720, fallback: 320, maxShare: 0.6 };
/** The height of the row of thumbnails, without the scroll bar. The default is about the size of a 16:9 thumbnail of 136 pixels. */
export const RAIL_HEIGHT = { min: 64, max: 360, fallback: 98, maxShare: 0.45 };

const STEP = 16;
const STEP_BIG = 64;
const COMMIT_DELAY_MS = 300;

/** `value` kept between the smallest size and the largest the window allows (a share of `available`, never more than `limits.max`). */
export function clampSize(value, available, limits = COMMENTS_WIDTH) {
  const ceiling = Math.max(limits.min, Math.min(limits.max, Math.floor(available * limits.maxShare)));
  return Math.min(ceiling, Math.max(limits.min, Math.round(value)));
}

/** `axis` is 'x' for a pane on the right (dragging its edge to the left makes it wider) and 'y' for one at the bottom (dragging up makes it taller). */
export function createResize({ handle, layout, onCommit, limits = COMMENTS_WIDTH, variable = '--comments-width', axis = 'x' }) {
  const vertical = axis === 'y';
  let size = null;
  let dragging = false;
  let timer = null;
  let origin = null;

  const available = () => {
    const rect = layout.getBoundingClientRect();
    return vertical ? rect.height : rect.width;
  };

  function show() {
    if (size === null) layout.style.removeProperty(variable);
    else layout.style.setProperty(variable, `${size}px`);
    handle.setAttribute('aria-valuemin', String(limits.min));
    handle.setAttribute('aria-valuemax', String(limits.max));
    handle.setAttribute('aria-valuenow', String(size ?? limits.fallback));
  }

  /** Applies a size (null for the default). With `commit`, it is also reported for saving. */
  function set(value, { commit = false } = {}) {
    size = value === null ? null : clampSize(value, available(), limits);
    show();
    if (commit) onCommit(size);
  }

  const current = () => size ?? limits.fallback;

  function commitSoon() {
    clearTimeout(timer);
    timer = setTimeout(() => onCommit(size), COMMIT_DELAY_MS);
  }

  function finishDrag() {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('is-dragging');
    document.body.classList.remove('is-resizing', 'is-resizing-y');
    onCommit(size);
  }

  handle.addEventListener('pointerdown', (event) => {
    if (event.button !== 0) return;
    dragging = true;
    origin = { pointer: vertical ? event.clientY : event.clientX, size: current() };
    handle.setPointerCapture(event.pointerId);
    handle.classList.add('is-dragging');
    document.body.classList.add(vertical ? 'is-resizing-y' : 'is-resizing');
    event.preventDefault();
  });
  handle.addEventListener('pointermove', (event) => {
    if (!dragging) return;
    // The pane is on the right or at the bottom, so moving its edge to the left or up makes it bigger.
    if (vertical) set(origin.size + origin.pointer - event.clientY);
    else set(layout.getBoundingClientRect().right - event.clientX);
  });
  handle.addEventListener('pointerup', finishDrag);
  handle.addEventListener('pointercancel', finishDrag);
  handle.addEventListener('lostpointercapture', finishDrag);
  handle.addEventListener('dblclick', () => set(null, { commit: true }));
  handle.addEventListener('keydown', (event) => {
    const step = event.shiftKey ? STEP_BIG : STEP;
    const [bigger, smaller] = vertical ? ['ArrowUp', 'ArrowDown'] : ['ArrowLeft', 'ArrowRight'];
    if (event.key === bigger) set(current() + step);
    else if (event.key === smaller) set(current() - step);
    else if (event.key === 'Home') set(limits.min);
    else if (event.key === 'End') set(limits.max);
    else return;
    // The page's own shortcuts (arrows change the slide) must not see these keys.
    event.preventDefault();
    event.stopPropagation();
    commitSoon();
  });

  show();
  return { set: (value) => set(value), get size() { return size; } };
}
