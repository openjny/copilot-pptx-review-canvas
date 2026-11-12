// The drag handle on the left edge of the comments pane. It only works out a width: `onCommit` is called with the final width (or null for
// the default) when the user lets go, so the page can save it. The width is the CSS variable `--comments-width` on `layout`; nothing
// here sets an inline style attribute, which the page's content security policy would block.

export const COMMENTS_WIDTH = { min: 240, max: 720, fallback: 320, maxShare: 0.6 };

const STEP = 16;
const STEP_BIG = 64;
const COMMIT_DELAY_MS = 300;

/** `value` kept between the smallest width and the largest the window allows (a share of `available`, never more than `limits.max`). */
export function clampWidth(value, available, limits = COMMENTS_WIDTH) {
  const ceiling = Math.max(limits.min, Math.min(limits.max, Math.floor(available * limits.maxShare)));
  return Math.min(ceiling, Math.max(limits.min, Math.round(value)));
}

export function createResize({ handle, layout, onCommit, limits = COMMENTS_WIDTH }) {
  let width = null;
  let dragging = false;
  let timer = null;

  const available = () => layout.getBoundingClientRect().width;

  function show() {
    if (width === null) layout.style.removeProperty('--comments-width');
    else layout.style.setProperty('--comments-width', `${width}px`);
    handle.setAttribute('aria-valuemin', String(limits.min));
    handle.setAttribute('aria-valuemax', String(limits.max));
    handle.setAttribute('aria-valuenow', String(width ?? limits.fallback));
  }

  /** Applies a width (null for the default). With `commit`, it is also reported for saving. */
  function set(value, { commit = false } = {}) {
    width = value === null ? null : clampWidth(value, available(), limits);
    show();
    if (commit) onCommit(width);
  }

  const current = () => width ?? limits.fallback;

  function commitSoon() {
    clearTimeout(timer);
    timer = setTimeout(() => onCommit(width), COMMIT_DELAY_MS);
  }

  function finishDrag() {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('is-dragging');
    document.body.classList.remove('is-resizing');
    onCommit(width);
  }

  handle.addEventListener('pointerdown', (event) => {
    if (event.button !== 0) return;
    dragging = true;
    handle.setPointerCapture(event.pointerId);
    handle.classList.add('is-dragging');
    document.body.classList.add('is-resizing');
    event.preventDefault();
  });
  handle.addEventListener('pointermove', (event) => {
    if (dragging) set(layout.getBoundingClientRect().right - event.clientX);
  });
  handle.addEventListener('pointerup', finishDrag);
  handle.addEventListener('pointercancel', finishDrag);
  handle.addEventListener('lostpointercapture', finishDrag);
  handle.addEventListener('dblclick', () => set(null, { commit: true }));
  handle.addEventListener('keydown', (event) => {
    const step = event.shiftKey ? STEP_BIG : STEP;
    // The pane is on the right, so moving its edge to the left makes it wider.
    if (event.key === 'ArrowLeft') set(current() + step);
    else if (event.key === 'ArrowRight') set(current() - step);
    else if (event.key === 'Home') set(limits.min);
    else if (event.key === 'End') set(limits.max);
    else return;
    // The page's own shortcuts (arrows change the slide) must not see these keys.
    event.preventDefault();
    event.stopPropagation();
    commitSoon();
  });

  show();
  return { set: (value) => set(value), get width() { return width; } };
}
