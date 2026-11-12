// A small popup menu at a point of the page: the right-click menu of a slide in the strip. It knows nothing about slides. It is given the
// entries, shows them where the pointer is, and calls the `onSelect` of the one that is chosen. Entries: `{ label, title, danger, disabled,
// divider, onSelect }`; `divider` draws a line above the entry.

const MARGIN = 6;

export function createMenu({ root }) {
  let opener = null;

  const items = () => [...root.querySelectorAll('[role="menuitem"]:not([aria-disabled="true"])')];
  const isOpen = () => !root.hidden;

  /** Closes the menu. The focus goes back to the element that opened it, unless the user went somewhere else (`restore: false`). */
  function close({ restore = true } = {}) {
    if (root.hidden) return;
    root.hidden = true;
    root.replaceChildren();
    const target = opener;
    opener = null;
    if (restore && target?.isConnected) target.focus();
  }

  function build(entry) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `ctx__item${entry.danger ? ' ctx__item--danger' : ''}${entry.divider ? ' ctx__item--divided' : ''}`;
    button.setAttribute('role', 'menuitem');
    button.tabIndex = -1;
    button.textContent = entry.label;
    if (entry.title) button.title = entry.title;
    if (entry.disabled) button.setAttribute('aria-disabled', 'true');
    button.addEventListener('click', () => {
      if (entry.disabled) return;
      close({ restore: false });
      entry.onSelect();
    });
    return button;
  }

  /** Opens the menu with its top-left corner at (`x`, `y`) in window coordinates, moved so that it stays inside the window. */
  function open(entries, { x, y, from = null }) {
    close({ restore: false });
    opener = from;
    root.replaceChildren(...entries.map(build));
    root.hidden = false;
    const { width, height } = root.getBoundingClientRect();
    const left = Math.max(MARGIN, Math.min(x, window.innerWidth - width - MARGIN));
    // The strip is at the bottom of the window, so a menu that does not fit below the pointer opens above it.
    const top = y + height + MARGIN > window.innerHeight ? Math.max(MARGIN, y - height) : y;
    root.style.left = `${left}px`;
    root.style.top = `${top}px`;
    items()[0]?.focus();
  }

  root.addEventListener('keydown', (event) => {
    // The page's own shortcuts (arrows change the slide, P picks the pen) must not see keys typed in the menu.
    event.stopPropagation();
    const list = items();
    const index = list.indexOf(document.activeElement);
    if (event.key === 'ArrowDown') list[(index + 1) % list.length]?.focus();
    else if (event.key === 'ArrowUp') list[(index - 1 + list.length) % list.length]?.focus();
    else if (event.key === 'Home') list[0]?.focus();
    else if (event.key === 'End') list.at(-1)?.focus();
    else if (event.key === 'Escape' || event.key === 'Tab') close();
    else return;
    event.preventDefault();
  });

  // Anything that happens outside the menu closes it: a click, another window, a resize, a scroll of the strip.
  document.addEventListener('pointerdown', (event) => {
    if (isOpen() && !(event.target instanceof Node && root.contains(event.target))) close({ restore: false });
  }, true);
  window.addEventListener('blur', () => close({ restore: false }));
  window.addEventListener('resize', () => close({ restore: false }));
  document.addEventListener('scroll', () => close({ restore: false }), true);

  return { open, close, isOpen };
}
