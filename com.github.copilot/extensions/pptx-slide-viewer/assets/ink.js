// Drawing layer over the slide image (pen and highlighter). Strokes are kept as fractions of the slide (0..1) so they stay put at any size,
// and the server holds the authoritative copy: every change is sent in order, and the page reloads when someone else changed it.

import { t } from './i18n.js';

const COLORS = { pen: '#e5231b', highlighter: '#ffd400' };
const WIDTHS = {
  thin: { pen: 0.0025, highlighter: 0.012 },
  medium: { pen: 0.004, highlighter: 0.02 },
  thick: { pen: 0.008, highlighter: 0.035 },
};
const HIGHLIGHTER_ALPHA = 0.35;
const MIN_STEP = 0.002;
const MAX_POINTS = 4000;
const ERASE_RADIUS = 0.012;
const SELECT_COLOR = '#1f6feb';
// A drag shorter than this many pixels (both ways) is a click: a point, not an area.
const CLICK_PX = 6;

const byId = (id) => document.getElementById(id);
const clamp01 = (value) => Math.min(1, Math.max(0, value));
const round = (value) => Math.round(value * 1e4) / 1e4;
const newId = () => globalThis.crypto?.randomUUID?.() ?? `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;

function parseVersion(version) {
  const cut = String(version ?? '').lastIndexOf(':');
  return cut < 0 ? { key: '', rev: -1 } : { key: version.slice(0, cut), rev: Number(version.slice(cut + 1)) };
}

function paint(context, strokes, width, height) {
  context.clearRect(0, 0, width, height);
  context.lineCap = 'round';
  context.lineJoin = 'round';
  for (const stroke of strokes) {
    context.globalAlpha = stroke.tool === 'highlighter' ? HIGHLIGHTER_ALPHA : 1;
    context.strokeStyle = stroke.color;
    context.fillStyle = stroke.color;
    context.lineWidth = Math.max(1, stroke.width * width);
    const [first, ...rest] = stroke.points;
    context.beginPath();
    if (rest.length === 0) {
      context.arc(first[0] * width, first[1] * height, context.lineWidth / 2, 0, 2 * Math.PI);
      context.fill();
      continue;
    }
    context.moveTo(first[0] * width, first[1] * height);
    for (const point of rest) context.lineTo(point[0] * width, point[1] * height);
    context.stroke();
  }
  context.globalAlpha = 1;
}

/** The dashed rectangle shown while the user drags out an area to comment on. */
function paintSelection(context, gesture, width, height) {
  const x = Math.min(gesture.start[0], gesture.end[0]) * width;
  const y = Math.min(gesture.start[1], gesture.end[1]) * height;
  const w = Math.abs(gesture.end[0] - gesture.start[0]) * width;
  const h = Math.abs(gesture.end[1] - gesture.start[1]) * height;
  context.save();
  context.fillStyle = 'rgba(31, 111, 235, 0.12)';
  context.strokeStyle = SELECT_COLOR;
  context.lineWidth = 2;
  context.setLineDash([7, 5]);
  context.fillRect(x, y, w, h);
  context.strokeRect(x, y, w, h);
  context.restore();
}

/** Distance from p to the segment a-b, in slide-width units (y is scaled by the slide's aspect ratio). */
function segmentDistance(p, a, b, aspect) {
  const px = p[0];
  const py = p[1] / aspect;
  const ax = a[0];
  const ay = a[1] / aspect;
  const dx = b[0] - ax;
  const dy = b[1] / aspect - ay;
  const length = dx * dx + dy * dy;
  const t = length === 0 ? 0 : Math.min(1, Math.max(0, ((px - ax) * dx + (py - ay) * dy) / length));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

function touches(stroke, point, aspect) {
  const reach = ERASE_RADIUS + stroke.width / 2;
  const { points } = stroke;
  if (points.length === 1) return segmentDistance(point, points[0], points[0], aspect) <= reach;
  for (let i = 1; i < points.length; i++) if (segmentDistance(point, points[i - 1], points[i], aspect) <= reach) return true;
  return false;
}

export function createInk({ box, canvas, getState, getJson, post, onError, announce, history = null, onSelect = () => {}, onVisibleChange = () => {} }) {
  const ui = {
    tools: [...document.querySelectorAll('button[data-tool]')],
    swatches: [...document.querySelectorAll('.swatch')],
    sizes: [...document.querySelectorAll('[data-size]')],
    undo: byId('ink-undo'),
    redo: byId('ink-redo'),
    clear: byId('ink-clear'),
    toggle: byId('ink-toggle'),
    count: byId('ink-count'),
    bar: byId('ink'),
  };
  const context = canvas.getContext('2d');
  const model = {
    tool: 'pen',
    colors: { ...COLORS },
    size: 'medium',
    visible: true,
    deckKey: '',
    slide: 0,
    draft: null,
    scope: '',
    loadedScope: '',
    loading: false,
    retryAt: 0,
    strokes: [],
    known: { key: '', rev: -1 },
    epoch: 0,
    pending: 0,
    stacks: new Map(),
    gesture: null,
  };
  let queue = Promise.resolve();
  let frame = 0;

  const state = () => getState();
  const enabled = () => Boolean(state()?.deck) && state().phase === 'ready' && model.visible && model.loadedScope === model.scope;
  const stack = () => {
    const key = model.scope;
    if (!model.stacks.has(key)) model.stacks.set(key, { undo: [], redo: [] });
    return model.stacks.get(key);
  };
  /** The page strokes belong to: a slide, or a draft page (a new-slide idea). */
  const where = () => (model.draft ? { draft: model.draft } : { slide: model.slide });
  const pageQuery = () => (model.draft ? `draft=${encodeURIComponent(model.draft)}` : `slide=${model.slide}`);
  const aspect = () => {
    const rect = canvas.getBoundingClientRect();
    return rect.height > 0 ? rect.width / rect.height : 16 / 9;
  };

  // ---- drawing ----------------------------------------------------------------------------------------------

  function redraw() {
    frame = 0;
    const gesture = model.gesture;
    const strokes = gesture?.stroke ? [...model.strokes, gesture.stroke] : model.strokes;
    paint(context, strokes, canvas.width, canvas.height);
    if (gesture?.type === 'select') paintSelection(context, gesture, canvas.width, canvas.height);
  }

  function schedule() {
    if (!frame) frame = requestAnimationFrame(redraw);
  }

  function fit() {
    const rect = box.getBoundingClientRect();
    const scale = window.devicePixelRatio || 1;
    const width = Math.max(1, Math.round(rect.width * scale));
    const height = Math.max(1, Math.round(rect.height * scale));
    if (canvas.width !== width || canvas.height !== height) {
      canvas.width = width;
      canvas.height = height;
    }
    schedule();
  }

  new ResizeObserver(fit).observe(box);

  // ---- talking to the server --------------------------------------------------------------------------------

  function adopt(version) {
    const next = parseVersion(version);
    if (next.key !== model.known.key || next.rev > model.known.rev) model.known = next;
  }

  function resync() {
    model.loadedScope = '';
    sync(state());
  }

  /** Requests go out one at a time, in the order the user made them. */
  function remote(path, body) {
    model.pending += 1;
    const run = queue.then(() => post(path, body));
    queue = run.catch(() => {});
    run.then((result) => adopt(result.version)).catch((error) => {
      onError(error);
      resync();
    }).finally(() => {
      model.pending -= 1;
      if (model.pending === 0) sync(state());
    });
  }

  function load() {
    const { scope, epoch } = model;
    model.loading = true;
    getJson(`api/annotations?${pageQuery()}`).then((data) => {
      if (scope !== model.scope || epoch !== model.epoch) return;
      model.strokes = data.strokes;
      model.loadedScope = scope;
      adopt(data.version);
      redraw();
    }).catch(() => {
      model.retryAt = Date.now() + 1000;
      setTimeout(() => sync(state()), 1000);
    }).finally(() => {
      model.loading = false;
      sync(state());
    });
  }

  // ---- changes the user makes -------------------------------------------------------------------------------

  function addLocal(strokes) {
    model.epoch += 1;
    model.strokes = [...model.strokes, ...strokes];
    for (const stroke of strokes) remote('api/annotations/add', { ...where(), stroke });
  }

  function removeLocal(strokes) {
    const ids = new Set(strokes.map((stroke) => stroke.id));
    model.epoch += 1;
    model.strokes = model.strokes.filter((stroke) => !ids.has(stroke.id));
    remote('api/annotations/remove', { ...where(), ids: [...ids] });
  }

  function record(entry) {
    const { undo, redo } = stack();
    undo.push({ ...entry, at: Date.now() });
    redo.length = 0;
  }

  /**
   * The deck's own history (a deleted slide) when it comes before the strokes of this page for `kind` ('undo' or 'redo'), else null.
   * Whichever was done last is undone first, and whichever was undone last is redone first.
   */
  function deckFirst(kind) {
    const deck = history?.peek(kind);
    if (!deck) return null;
    const local = stack()[kind].at(-1);
    return !local || deck.at >= local.at ? deck : null;
  }

  function undo() {
    if (deckFirst('undo')) {
      history.run('undo');
      return;
    }
    const { undo: undone, redo } = stack();
    const entry = undone.pop();
    if (!entry) return;
    if (entry.type === 'add') removeLocal(entry.strokes);
    else addLocal(entry.strokes);
    entry.at = Date.now();
    redo.push(entry);
    announce(t('ink.undone'));
    schedule();
    renderUi();
  }

  function redo() {
    if (deckFirst('redo')) {
      history.run('redo');
      return;
    }
    const { undo: undone, redo: redone } = stack();
    const entry = redone.pop();
    if (!entry) return;
    if (entry.type === 'add') addLocal(entry.strokes);
    else removeLocal(entry.strokes);
    entry.at = Date.now();
    undone.push(entry);
    announce(t('ink.redone'));
    schedule();
    renderUi();
  }

  function clearSlide() {
    if (!enabled() || model.strokes.length === 0) return;
    const strokes = model.strokes;
    record({ type: 'remove', strokes });
    model.epoch += 1;
    model.strokes = [];
    remote('api/annotations/clear', where());
    announce(model.draft ? t('ink.clearedDraft') : t('ink.clearedSlide', { n: model.slide }));
    schedule();
    renderUi();
  }

  // ---- pointer ----------------------------------------------------------------------------------------------

  function pointOf(event) {
    const rect = canvas.getBoundingClientRect();
    return [clamp01((event.clientX - rect.left) / rect.width), clamp01((event.clientY - rect.top) / rect.height)];
  }

  function extend(event) {
    const { gesture } = model;
    const point = pointOf(event);
    if (gesture.type === 'select') {
      gesture.end = point;
      return;
    }
    if (gesture.type === 'erase') {
      const keep = [];
      for (const stroke of model.strokes) (touches(stroke, point, aspect()) ? gesture.erased : keep).push(stroke);
      if (keep.length !== model.strokes.length) model.strokes = keep;
      return;
    }
    const { points } = gesture.stroke;
    const last = points[points.length - 1];
    if (points.length >= MAX_POINTS) return;
    if (Math.hypot(point[0] - last[0], (point[1] - last[1]) / aspect()) < MIN_STEP) return;
    points.push([round(point[0]), round(point[1])]);
  }

  function begin(event) {
    if (!enabled() || !model.tool || model.gesture || (event.pointerType === 'mouse' && event.button !== 0)) return;
    event.preventDefault();
    try {
      canvas.setPointerCapture(event.pointerId);
    } catch {
      // The pointer may already be gone; drawing still works while it stays over the canvas.
    }
    const point = pointOf(event);
    if (model.tool === 'select') {
      model.gesture = { type: 'select', pointerId: event.pointerId, start: point, end: point };
    } else if (model.tool === 'eraser') {
      model.gesture = { type: 'erase', pointerId: event.pointerId, erased: [] };
      extend(event);
    } else {
      model.gesture = {
        type: 'draw',
        pointerId: event.pointerId,
        stroke: { id: newId(), tool: model.tool, color: model.colors[model.tool], width: WIDTHS[model.size][model.tool], points: [[round(point[0]), round(point[1])]] },
      };
    }
    schedule();
  }

  function move(event) {
    const { gesture } = model;
    if (!gesture || gesture.pointerId !== event.pointerId) return;
    const coalesced = event.getCoalescedEvents?.();
    for (const sample of coalesced?.length ? coalesced : [event]) extend(sample);
    schedule();
  }

  /** A click is a point; a drag is the area it spans. Both in fractions of the slide. */
  function selectionAnchor(gesture) {
    const rect = canvas.getBoundingClientRect();
    const [x0, y0] = gesture.start;
    const [x1, y1] = gesture.end;
    if (Math.abs(x1 - x0) * rect.width < CLICK_PX && Math.abs(y1 - y0) * rect.height < CLICK_PX) return { x: round(x0), y: round(y0), w: 0, h: 0 };
    return { x: round(Math.min(x0, x1)), y: round(Math.min(y0, y1)), w: round(Math.abs(x1 - x0)), h: round(Math.abs(y1 - y0)) };
  }

  function finish(event, { keep }) {
    const { gesture } = model;
    if (!gesture || gesture.pointerId !== event.pointerId) return;
    model.gesture = null;
    if (gesture.type === 'select') {
      if (keep) onSelect(selectionAnchor(gesture));
      schedule();
      renderUi();
      return;
    }
    if (gesture.type === 'draw' && keep) {
      model.epoch += 1;
      model.strokes = [...model.strokes, gesture.stroke];
      record({ type: 'add', strokes: [gesture.stroke] });
      remote('api/annotations/add', { ...where(), stroke: gesture.stroke });
    } else if (gesture.type === 'erase' && gesture.erased.length) {
      if (keep) {
        model.epoch += 1;
        record({ type: 'remove', strokes: gesture.erased });
        remote('api/annotations/remove', { ...where(), ids: gesture.erased.map((stroke) => stroke.id) });
      } else {
        model.strokes = [...model.strokes, ...gesture.erased];
      }
    }
    schedule();
    renderUi();
  }

  canvas.addEventListener('pointerdown', begin);
  canvas.addEventListener('pointermove', move);
  canvas.addEventListener('pointerup', (event) => finish(event, { keep: true }));
  canvas.addEventListener('pointercancel', (event) => finish(event, { keep: false }));
  canvas.addEventListener('contextmenu', (event) => event.preventDefault());

  function cancelGesture() {
    if (!model.gesture) return;
    model.gesture = null;
    schedule();
  }

  // ---- toolbar ----------------------------------------------------------------------------------------------

  function pick(tool) {
    model.tool = model.tool === tool ? null : tool;
    const names = { select: t('ink.tool.select'), pen: t('ink.tool.pen'), highlighter: t('ink.tool.highlighter'), eraser: t('ink.tool.eraser') };
    announce(model.tool ? t('ink.pickTool', { tool: names[model.tool] }) : t('ink.pickNone'));
    renderUi();
  }

  for (const button of ui.tools) button.addEventListener('click', () => pick(button.dataset.tool));
  for (const swatch of ui.swatches) {
    swatch.addEventListener('click', () => {
      if (model.tool !== 'pen' && model.tool !== 'highlighter') model.tool = 'pen';
      model.colors[model.tool] = swatch.dataset.color;
      renderUi();
    });
  }
  for (const button of ui.sizes) {
    button.addEventListener('click', () => {
      model.size = button.dataset.size;
      renderUi();
    });
  }
  ui.undo.addEventListener('click', undo);
  ui.redo.addEventListener('click', redo);
  ui.clear.addEventListener('click', clearSlide);
  ui.toggle.addEventListener('click', toggleVisible);

  function toggleVisible() {
    model.visible = !model.visible;
    cancelGesture();
    announce(model.visible ? t('ink.shown') : t('ink.hidden'));
    onVisibleChange(model.visible);
    renderUi();
  }

  function renderUi() {
    const current = state();
    const on = enabled();
    const colour = model.colors[model.tool] ?? null;
    for (const button of ui.tools) button.setAttribute('aria-pressed', String(model.tool === button.dataset.tool));
    for (const swatch of ui.swatches) swatch.setAttribute('aria-pressed', String(swatch.dataset.color === colour));
    for (const button of ui.sizes) button.setAttribute('aria-pressed', String(model.size === button.dataset.size));
    const entry = current?.deck ? stack() : { undo: [], redo: [] };
    const deckUndo = deckFirst('undo');
    const deckRedo = deckFirst('redo');
    ui.undo.disabled = !((on && entry.undo.length > 0) || deckUndo);
    ui.redo.disabled = !((on && entry.redo.length > 0) || deckRedo);
    ui.undo.title = deckUndo ? t('ink.undoLabel', { label: deckUndo.label }) : t('ink.undoTitle');
    ui.redo.title = deckRedo ? t('ink.redoLabel', { label: deckRedo.label }) : t('ink.redoTitle');
    ui.clear.disabled = !on || model.strokes.length === 0;
    ui.toggle.disabled = !current?.deck;
    ui.toggle.setAttribute('aria-pressed', String(model.visible));
    for (const control of [...ui.tools, ...ui.swatches, ...ui.sizes]) control.disabled = !current?.deck;
    ui.bar.hidden = !current?.deck;
    canvas.hidden = !current?.deck || !model.visible;
    if (on && model.tool) canvas.dataset.mode = model.tool;
    else delete canvas.dataset.mode;
    const counts = current?.annotations?.counts ?? {};
    const drafts = current?.annotations?.drafts ?? [];
    const total = Object.values(counts).reduce((sum, count) => sum + count, 0) + drafts.reduce((sum, draft) => sum + draft.count, 0);
    const here = model.draft ? drafts.find((draft) => draft.id === model.draft)?.count ?? 0 : counts[model.slide] ?? 0;
    ui.count.textContent = current?.deck
      ? `${!model.visible ? t('ink.countHidden') : current.phase !== 'ready' ? t('ink.countRendering') : ''}${t('ink.count', { here, total })}`
      : '';
  }

  // ---- called by the page -----------------------------------------------------------------------------------

  /** Brings the layer in line with the page state: which page is shown (a slide or a draft), whether someone else changed the marks, and the toolbar. */
  function sync(current) {
    if (current?.deck) {
      if (current.deck.key !== model.deckKey) {
        model.stacks.clear();
        model.deckKey = current.deck.key;
      }
      model.slide = current.current;
      model.draft = current.draft ?? null;
      const scope = `${model.deckKey}:${model.draft ? `d${model.draft}` : `s${model.slide}`}`;
      if (scope !== model.scope) {
        cancelGesture();
        model.scope = scope;
        model.loadedScope = '';
        model.strokes = [];
        model.epoch += 1;
      }
      if (current.phase === 'ready' && model.pending === 0 && !model.gesture && !model.loading && Date.now() >= model.retryAt) {
        const server = parseVersion(current.annotations?.version);
        const stale = Boolean(server.key) && server.key === model.known.key && server.rev > model.known.rev;
        if (model.loadedScope !== scope || stale) load();
      }
    } else {
      cancelGesture();
      model.deckKey = '';
      model.slide = 0;
      model.draft = null;
      model.scope = '';
      model.loadedScope = '';
      model.strokes = [];
    }
    renderUi();
    schedule();
  }

  /** Returns true when the key was a drawing shortcut. */
  function handleKey(event) {
    const modifier = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();
    if (modifier && !event.altKey && key === 'z') {
      if (event.shiftKey) redo();
      else undo();
      return true;
    }
    if (modifier && !event.altKey && key === 'y') {
      redo();
      return true;
    }
    if (modifier || event.altKey || !state()?.deck) return false;
    const tool = { c: 'select', p: 'pen', h: 'highlighter', e: 'eraser' }[key];
    if (tool) {
      pick(tool);
      return true;
    }
    if (key === 'm') {
      toggleVisible();
      return true;
    }
    return false;
  }

  return { sync, handleKey };
}
