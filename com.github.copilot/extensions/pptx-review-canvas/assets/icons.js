// Line icons for the toolbar, drawn on a 24 x 24 grid in the current text colour. Markup elsewhere asks for one with
// `<span data-icon="pen"></span>`; `mountIcons()` fills those in.

const ICONS = {
  pen: '<path d="M4 20l1-4.2L16.3 4.5a2.1 2.1 0 0 1 3 0l.2.2a2.1 2.1 0 0 1 0 3L8.2 19 4 20z"/><path d="M14.5 6.3l3.2 3.2"/>',
  select: '<rect x="3" y="4" width="12.5" height="10.5" rx="1" stroke-dasharray="2.6 2.4"/><path d="M12.5 11.5l8 3.2-3.5 1.4-1.4 3.6z"/>',
  comment: '<path d="M5 4h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1h-8l-5 4v-4H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1z"/><path d="M8 8.5h8M8 11.5h5"/>',
  highlighter: '<path d="M15.5 3.5l5 5-8.7 8.7-5-5z"/><path d="M6.8 12.2l-2.6 4.6 3 3 4.6-2.6"/><path d="M4 21h5"/>',
  eraser: '<path d="M4.2 15.3l9.5-9.5a2 2 0 0 1 2.8 0l3.7 3.7a2 2 0 0 1 0 2.8L12.3 20H8l-3.8-3.8a1.4 1.4 0 0 1 0-1z"/><path d="M9.5 10.5l5 5"/><path d="M12.5 20H21"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
  trash: '<path d="M4 7h16"/><path d="M9 7V4.5h6V7"/><path d="M6 7l1 13h10l1-13"/><path d="M10 11v5M14 11v5"/>',
  eye: '<path d="M2 12s3.6-6.5 10-6.5S22 12 22 12s-3.6 6.5-10 6.5S2 12 2 12z"/><circle cx="12" cy="12" r="2.8"/>',
  'eye-off': '<path d="M2 12s3.6-6.5 10-6.5c1.9 0 3.6.5 5 1.3M22 12s-3.6 6.5-10 6.5c-1.9 0-3.6-.5-5-1.3"/><path d="M9.9 9.9a2.8 2.8 0 0 0 4.2 4.2"/><path d="M4 4l16 16"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  help: '<circle cx="12" cy="12" r="9"/><path d="M9.6 9.4a2.5 2.5 0 1 1 3.6 2.2c-.8.4-1.2 1-1.2 1.8"/><path d="M12 17h.01"/>',
  fullscreen: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
  notes: '<path d="M6 3h8.5L19 7.5V21H6z"/><path d="M14 3v5h5"/><path d="M9 13h7M9 17h7"/>',
  refresh: '<path d="M20 11a8 8 0 1 0-2.1 5.9"/><path d="M20 4v7h-7"/>',
  external: '<path d="M14 4h6v6"/><path d="M20 4l-9 9"/><path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>',
  folder: '<path d="M3 18V6.5A1.5 1.5 0 0 1 4.5 5h4.2L11 7.5h8.5A1.5 1.5 0 0 1 21 9v1"/><path d="M3 18l2.6-7.5h16.3L19.3 18z"/>',
  'file-plus': '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M12 11.5v6M9 14.5h6"/>',
  'line-thin': '<path d="M4 12h16" stroke-width="1.4"/>',
  'line-medium': '<path d="M4 12h16" stroke-width="3"/>',
  'line-thick': '<path d="M4 12h16" stroke-width="5.5"/>',
  close: '<path d="M6 6l12 12M18 6L6 18"/>',
};

export function icon(name) {
  const body = ICONS[name];
  if (!body) return '';
  return `<svg class="icon" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">${body}</svg>`;
}

/** Fills every `[data-icon]` under `root` with its icon. */
export function mountIcons(root = document) {
  for (const holder of root.querySelectorAll('[data-icon]')) {
    if (!holder.firstElementChild) holder.innerHTML = icon(holder.dataset.icon);
  }
}
