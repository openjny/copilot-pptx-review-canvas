import { decodePng, encodePng } from './png.mjs';

export const HIGHLIGHTER_ALPHA = 0.35;

export function parseColor(color) {
  const match = /^#([0-9a-f]{6})$/i.exec(String(color));
  if (!match) throw new Error(`Invalid color format: ${color}`);
  const value = parseInt(match[1], 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

/** Draws one anti-aliased, round-capped polyline. Coordinates and width are in pixels; the stroke is blended once, so overlaps do not darken it. */
export function drawPolyline(image, points, widthPx, color, alpha = 1) {
  if (points.length === 0) return;
  const { width, height, data } = image;
  const radius = Math.max(0.75, widthPx / 2);
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const [x, y] of points) {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  }
  const x0 = Math.max(0, Math.floor(minX - radius - 1));
  const y0 = Math.max(0, Math.floor(minY - radius - 1));
  const x1 = Math.min(width - 1, Math.ceil(maxX + radius + 1));
  const y1 = Math.min(height - 1, Math.ceil(maxY + radius + 1));
  if (x1 < x0 || y1 < y0) return;
  const boxW = x1 - x0 + 1;
  const mask = new Float32Array(boxW * (y1 - y0 + 1));

  const segments = points.length === 1 ? [[points[0], points[0]]] : points.slice(1).map((point, i) => [points[i], point]);
  for (const [[ax, ay], [bx, by]] of segments) {
    const sx0 = Math.max(x0, Math.floor(Math.min(ax, bx) - radius - 1));
    const sy0 = Math.max(y0, Math.floor(Math.min(ay, by) - radius - 1));
    const sx1 = Math.min(x1, Math.ceil(Math.max(ax, bx) + radius + 1));
    const sy1 = Math.min(y1, Math.ceil(Math.max(ay, by) + radius + 1));
    const dx = bx - ax;
    const dy = by - ay;
    const lengthSquared = dx * dx + dy * dy;
    for (let y = sy0; y <= sy1; y++) {
      for (let x = sx0; x <= sx1; x++) {
        const px = x + 0.5;
        const py = y + 0.5;
        const t = lengthSquared === 0 ? 0 : Math.min(1, Math.max(0, ((px - ax) * dx + (py - ay) * dy) / lengthSquared));
        const distance = Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
        const coverage = Math.min(1, Math.max(0, radius + 0.5 - distance));
        if (coverage > 0) {
          const index = (y - y0) * boxW + (x - x0);
          if (coverage > mask[index]) mask[index] = coverage;
        }
      }
    }
  }

  const [r, g, b] = parseColor(color);
  for (let y = y0; y <= y1; y++) {
    for (let x = x0; x <= x1; x++) {
      const coverage = mask[(y - y0) * boxW + (x - x0)];
      if (coverage === 0) continue;
      const a = coverage * alpha;
      const i = (y * width + x) * 4;
      data[i] = Math.round(data[i] * (1 - a) + r * a);
      data[i + 1] = Math.round(data[i + 1] * (1 - a) + g * a);
      data[i + 2] = Math.round(data[i + 2] * (1 - a) + b * a);
    }
  }
}

function toHex([r, g, b]) {
  return `#${[r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('')}`;
}

/** Draws strokes whose points and width are fractions of the image (the form the viewer stores). */
export function drawStrokes(image, strokes) {
  for (const stroke of strokes) {
    const points = stroke.points.map(([x, y]) => [x * image.width, y * image.height]);
    drawPolyline(image, points, stroke.width * image.width, stroke.color, stroke.tool === 'highlighter' ? HIGHLIGHTER_ALPHA : 1);
  }
}

const SEGMENTS = { a: [[0, 0], [1, 0]], b: [[1, 0], [1, 0.5]], c: [[1, 0.5], [1, 1]], d: [[0, 1], [1, 1]], e: [[0, 0.5], [0, 1]], f: [[0, 0], [0, 0.5]], g: [[0, 0.5], [1, 0.5]] };
const DIGITS = ['abcdef', 'bc', 'abdeg', 'abcdg', 'bcfg', 'acdfg', 'acdefg', 'abc', 'abcdefg', 'abcdfg'];

/** A numbered dot: a dark disc with the number in seven-segment digits, so it needs no font. */
export function drawBadge(image, cx, cy, radius, number, { fill = '#1f1f1f', ink = '#ffffff' } = {}) {
  drawPolyline(image, [[cx, cy]], radius * 2, fill, 1);
  const digits = String(number).slice(0, 2).split('').map(Number);
  const digitHeight = radius * (digits.length > 1 ? 0.9 : 1.1);
  const digitWidth = digitHeight * 0.55;
  const gap = digitWidth * 0.45;
  const totalWidth = digits.length * digitWidth + (digits.length - 1) * gap;
  const left = cx - totalWidth / 2;
  const top = cy - digitHeight / 2;
  const strokeWidth = Math.max(1.5, radius * 0.17);
  digits.forEach((digit, index) => {
    const originX = left + index * (digitWidth + gap);
    for (const name of DIGITS[digit]) {
      const [[sx, sy], [ex, ey]] = SEGMENTS[name];
      drawPolyline(image, [[originX + sx * digitWidth, top + sy * digitHeight], [originX + ex * digitWidth, top + ey * digitHeight]], strokeWidth, ink, 1);
    }
  });
}

export const COMMENT_COLOR = '#1f6feb';

/** Draws the outline of a rectangle given in fractions of the image. */
export function drawRectOutline(image, rect, widthPx, color) {
  const x0 = rect.x * image.width;
  const y0 = rect.y * image.height;
  const x1 = (rect.x + rect.w) * image.width;
  const y1 = (rect.y + rect.h) * image.height;
  drawPolyline(image, [[x0, y0], [x1, y0], [x1, y1], [x0, y1], [x0, y0]], widthPx, color, 1);
}

/**
 * The slide image with its strokes on top. `badges` are `{ x, y, number }` in fractions of the image,
 * so that a numbered dot sits next to each mark the agent is told about. `comments` (`{ rects, badges }`, from `analyzeComments`)
 * adds a blue outline around each commented area and a blue numbered dot at each comment, so they can be told from the marks.
 */
export function compositeAnnotations(pngBuffer, strokes, badges = [], comments = { rects: [], badges: [] }) {
  const image = decodePng(pngBuffer);
  drawStrokes(image, strokes);
  for (const rect of comments.rects) drawRectOutline(image, rect, Math.max(2, image.width * 0.003), COMMENT_COLOR);
  const radius = Math.max(10, image.width * 0.0125);
  const place = (badge, style) => {
    const cx = Math.min(image.width - radius - 2, Math.max(radius + 2, badge.x * image.width));
    const cy = Math.min(image.height - radius - 2, Math.max(radius + 2, badge.y * image.height));
    drawBadge(image, cx, cy, radius, badge.number, style);
  };
  for (const badge of badges) place(badge);
  for (const badge of comments.badges) place(badge, { fill: COMMENT_COLOR });
  return encodePng(image);
}

export { toHex };
