import { commentKind } from './comments.mjs';
import { tEn } from './i18n.mjs';

const CLUSTER_GAP = 0.02;
const BADGE_RADIUS = 0.0125;
const MAX_CANDIDATES = 6;
const REGION_COLUMNS = ['region.left', 'region.center', 'region.right'];
const REGION_ROWS = ['region.top', 'region.middle', 'region.bottom'];
// How far around a point comment the shapes are looked for, as a fraction of the slide width.
const POINT_REACH = 0.012;

/** The box a stroke covers, as fractions of the slide. Width is a fraction of the image width, so the vertical half-extent scales by the aspect ratio. */
export function strokeBox(stroke, aspect) {
  let x0 = 1;
  let y0 = 1;
  let x1 = 0;
  let y1 = 0;
  for (const [x, y] of stroke.points) {
    x0 = Math.min(x0, x);
    y0 = Math.min(y0, y);
    x1 = Math.max(x1, x);
    y1 = Math.max(y1, y);
  }
  const halfX = stroke.width / 2;
  const halfY = (stroke.width / 2) * aspect;
  return { x0: Math.max(0, x0 - halfX), y0: Math.max(0, y0 - halfY), x1: Math.min(1, x1 + halfX), y1: Math.min(1, y1 + halfY) };
}

function overlaps(a, b, gapX, gapY) {
  return a.x0 - gapX <= b.x1 && b.x0 - gapX <= a.x1 && a.y0 - gapY <= b.y1 && b.y0 - gapY <= a.y1;
}

/** Groups strokes that touch or nearly touch (a circle and its arrow, two strokes of one underline) into numbered marks, in reading order. */
export function clusterMarks(strokes, aspect) {
  const boxes = strokes.map((stroke) => strokeBox(stroke, aspect));
  const parent = boxes.map((_, i) => i);
  const find = (i) => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      if (overlaps(boxes[i], boxes[j], CLUSTER_GAP, CLUSTER_GAP * aspect)) parent[find(j)] = find(i);
    }
  }
  const groups = new Map();
  boxes.forEach((box, i) => {
    const root = find(i);
    const group = groups.get(root) ?? { strokes: [], box: { x0: 1, y0: 1, x1: 0, y1: 0 } };
    group.strokes.push(strokes[i]);
    group.box = { x0: Math.min(group.box.x0, box.x0), y0: Math.min(group.box.y0, box.y0), x1: Math.max(group.box.x1, box.x1), y1: Math.max(group.box.y1, box.y1) };
    groups.set(root, group);
  });
  return [...groups.values()].sort((a, b) => Math.round(a.box.y0 / 0.12) - Math.round(b.box.y0 / 0.12) || a.box.x0 - b.box.x0);
}

function describeStroke(stroke, aspect) {
  const pts = stroke.points.map(([x, y]) => [x * aspect, y]);
  let length = 0;
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  pts.forEach(([x, y], i) => {
    x0 = Math.min(x0, x);
    y0 = Math.min(y0, y);
    x1 = Math.max(x1, x);
    y1 = Math.max(y1, y);
    if (i > 0) length += Math.hypot(x - pts[i - 1][0], y - pts[i - 1][1]);
  });
  const bw = x1 - x0;
  const bh = y1 - y0;
  const longest = Math.max(bw, bh);
  if (longest < 0.01) return 'dot';
  const ratio = length / Math.max(bw + bh, 1e-6);
  const gap = Math.hypot(pts[0][0] - pts.at(-1)[0], pts[0][1] - pts.at(-1)[1]);
  if (gap < 0.3 * longest && ratio > 1.35) return 'enclosure';
  if (ratio > 3) return 'scribble';
  if (Math.min(bw, bh) < 0.2 * longest && ratio < 1.3) return bw >= bh ? 'underline' : 'vertical-line';
  return ratio <= 1.3 ? 'line-or-arrow' : 'curve';
}

function region(box, t) {
  const cx = (box.x0 + box.x1) / 2;
  const cy = (box.y0 + box.y1) / 2;
  const col = Math.min(2, Math.floor(cx * 3));
  const row = Math.min(2, Math.floor(cy * 3));
  return col === 1 && row === 1 ? t('region.center') : t('region.cell', { column: t(REGION_COLUMNS[col]), row: t(REGION_ROWS[row]) });
}

function parentGroup(shapes, index) {
  const depth = shapes[index].depth;
  for (let i = index - 1; i >= 0; i--) if (shapes[i].depth === depth - 1) return shapes[i].name;
  return null;
}

/** The shapes a mark sits on or around, most useful first: text and pictures before empty containers. */
export function candidatesFor(box, kind, shapes, slideW, slideH) {
  if (!shapes?.length || !slideW || !slideH) return [];
  const pad = 0.012;
  const mark = { x: (box.x0 - pad) * slideW, y: (box.y0 - pad) * slideH, w: (box.x1 - box.x0 + 2 * pad) * slideW, h: (box.y1 - box.y0 + 2 * pad) * slideH };
  const markArea = mark.w * mark.h;
  const slideArea = slideW * slideH;
  const found = [];
  shapes.forEach((shape, index) => {
    if (shape.kind === 'group') return;
    const w = Math.max(shape.w, 1);
    const h = Math.max(shape.h, 1);
    const ix = Math.min(shape.x + w, mark.x + mark.w) - Math.max(shape.x, mark.x);
    const iy = Math.min(shape.y + h, mark.y + mark.h) - Math.max(shape.y, mark.y);
    if (ix <= 0 || iy <= 0) return;
    const inter = ix * iy;
    const area = w * h;
    if (inter / area < 0.6 && inter / markArea < 0.5) return;
    found.push({ shape, index, area, content: Boolean(shape.text || shape.alt), background: area > 0.7 * slideArea });
  });
  const enclosing = kind === 'enclosure' || markArea > 0.04 * slideArea;
  const specific = found.filter((candidate) => !candidate.background);
  let pool = specific.length ? specific : found;
  let near = false;
  if (!pool.length) {
    // A thin mark on the line between two shapes barely overlaps either: take the closest shapes instead.
    pool = nearest(mark, shapes, slideH, slideArea);
    near = true;
  }
  pool.sort((a, b) => Number(b.content) - Number(a.content) || (near ? a.gap - b.gap : enclosing ? b.area - a.area : a.area - b.area));
  return pool.slice(0, MAX_CANDIDATES).map(({ shape, index }) => {
    const group = parentGroup(shapes, index);
    const candidate = { name: shape.name, kind: shape.kind };
    const text = shape.text || shape.alt;
    if (text) candidate.text = text.length > 60 ? `${text.slice(0, 60)}...` : text;
    if (group) candidate.group = group;
    if (near) candidate.near = true;
    return candidate;
  });
}

const NEAR_REACH = 0.05;

/** Shapes within a small distance of the mark box (in points), closest first. */
function nearest(mark, shapes, slideH, slideArea) {
  const reach = NEAR_REACH * slideH;
  const result = [];
  shapes.forEach((shape, index) => {
    if (shape.kind === 'group') return;
    const w = Math.max(shape.w, 1);
    const h = Math.max(shape.h, 1);
    if (w * h > 0.7 * slideArea) return;
    const dx = Math.max(shape.x - (mark.x + mark.w), mark.x - (shape.x + w), 0);
    const dy = Math.max(shape.y - (mark.y + mark.h), mark.y - (shape.y + h), 0);
    const gap = Math.hypot(dx, dy);
    if (gap <= reach) result.push({ shape, index, gap, area: w * h, content: Boolean(shape.text || shape.alt) });
  });
  return result;
}

/** Turns the strokes of one slide into numbered marks (where, what kind, which shapes) and the badges to draw next to them. `t` names the regions. */
export function analyzeSlide({ strokes, aspect, shapes, slideW, slideH, t = tEn }) {
  const marks = clusterMarks(strokes, aspect).map((group, index) => {
    const kinds = [...new Set(group.strokes.map((stroke) => describeStroke(stroke, aspect)))];
    const kind = group.strokes.length === 1 ? kinds[0] : 'multiple';
    const { x0, y0, x1, y1 } = group.box;
    const mark = {
      number: index + 1,
      region: region(group.box, t),
      kind,
      ...(kind === 'multiple' ? { parts: kinds } : {}),
      strokeCount: group.strokes.length,
      colors: [...new Set(group.strokes.map((stroke) => stroke.color))],
      boxPercent: { x: Math.round(x0 * 100), y: Math.round(y0 * 100), w: Math.round((x1 - x0) * 100), h: Math.round((y1 - y0) * 100) },
      onShapes: candidatesFor(group.box, kind, shapes, slideW, slideH),
    };
    return { mark, box: group.box };
  });
  return {
    marks: marks.map(({ mark }) => mark),
    badges: marks.map(({ mark, box }) => ({ number: mark.number, x: box.x0 - BADGE_RADIUS, y: box.y0 - BADGE_RADIUS * aspect })),
  };
}

/**
 * Turns the comments of one page (the open threads) into what the agent reads (number, who wrote it, text, the replies, kind, where it
 * is, which shapes lie under it) and what the image shows: a rectangle around each area comment and a numbered badge at each point or
 * area. Comments about the whole page have a number and a text but no position.
 */
export function analyzeComments({ comments, aspect, shapes, slideW, slideH, t = tEn }) {
  const rects = [];
  const badges = [];
  const list = comments.map((comment, index) => {
    const number = index + 1;
    const kind = commentKind(comment);
    const replies = comment.replies ?? [];
    const entry = {
      number,
      id: comment.id,
      kind,
      author: comment.author ?? 'user',
      text: comment.text,
      ...(replies.length ? { replies: replies.map((reply) => ({ author: reply.author ?? 'user', text: reply.text })) } : {}),
    };
    if (!comment.anchor) return entry;
    const { x, y, w, h } = comment.anchor;
    const box = kind === 'point'
      ? { x0: Math.max(0, x - POINT_REACH), y0: Math.max(0, y - POINT_REACH * aspect), x1: Math.min(1, x + POINT_REACH), y1: Math.min(1, y + POINT_REACH * aspect) }
      : { x0: x, y0: y, x1: x + w, y1: y + h };
    entry.region = region(box, t);
    entry.boxPercent = kind === 'point'
      ? { x: Math.round(x * 100), y: Math.round(y * 100) }
      : { x: Math.round(x * 100), y: Math.round(y * 100), w: Math.round(w * 100), h: Math.round(h * 100) };
    entry.onShapes = candidatesFor(box, kind === 'area' ? 'enclosure' : 'point', shapes, slideW, slideH);
    if (kind === 'area') rects.push({ x, y, w, h });
    badges.push({ number, x: kind === 'point' ? x : x - BADGE_RADIUS, y: kind === 'point' ? y : y - BADGE_RADIUS * aspect });
    return entry;
  });
  return { comments: list, rects, badges };
}
