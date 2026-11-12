// Shapes for the pen and the highlighter: with Shift held a stroke is a straight line, with Ctrl a rectangle. The shape goes from where the
// stroke began to where the pointer is now, so the key can be pressed or let go at any time during the stroke.

/** 'rectangle' for Ctrl (or Cmd), 'line' for Shift, else null. Ctrl wins when both are held. */
export function shapeKey(event) {
  if (event.ctrlKey || event.metaKey) return 'rectangle';
  return event.shiftKey ? 'line' : null;
}

/**
 * The points of a stroke from `start` to `end` for `kind` ('line', 'rectangle', or null for the freehand `free` points). Points are fractions
 * of the slide. A shape with no size is the single point `start`, which is drawn as a dot.
 */
export function shapePoints(kind, free, start, end, { minStep = 0, aspect = 1 } = {}) {
  if (!kind) return free;
  const dx = end[0] - start[0];
  const dy = (end[1] - start[1]) / aspect;
  if (Math.hypot(dx, dy) < minStep) return [start];
  if (kind === 'line') return [start, end];
  return [start, [end[0], start[1]], end, [start[0], end[1]], start];
}
