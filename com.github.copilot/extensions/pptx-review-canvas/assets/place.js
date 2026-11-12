// Where the comment bubble goes: beside the spot it is about, inside the area it may use, with its tail pointing at the spot.

const SIDES = ['right', 'left', 'below', 'above'];

const clamp = (value, low, high) => Math.min(Math.max(value, low), Math.max(low, high));

/**
 * `anchor` and `bounds` are rectangles `{ left, top, right, bottom }` in one coordinate system, and `size` is `{ width, height }`.
 * The bubble goes on the first side of the anchor, in the order right, left, below, above, that has room for it; when none has, on the
 * side with the most room. It returns the bubble's `left` and `top`, that `side`, and `tail`: how far along the edge facing the anchor
 * the tail sits (from the top for left and right, from the left for below and above).
 */
export function placeBubble({ anchor, size, bounds, gap = 14, margin = 8, tailInset = 18, tailReach = 28 }) {
  const room = {
    right: bounds.right - margin - (anchor.right + gap) - size.width,
    left: anchor.left - gap - (bounds.left + margin) - size.width,
    below: bounds.bottom - margin - (anchor.bottom + gap) - size.height,
    above: anchor.top - gap - (bounds.top + margin) - size.height,
  };
  const side = SIDES.find((name) => room[name] >= 0) ?? SIDES.reduce((best, name) => (room[name] > room[best] ? name : best));
  const centerX = (anchor.left + anchor.right) / 2;
  const centerY = (anchor.top + anchor.bottom) / 2;
  const innerLeft = bounds.left + margin;
  const innerTop = bounds.top + margin;
  const innerRight = bounds.right - margin;
  const innerBottom = bounds.bottom - margin;

  if (side === 'right' || side === 'left') {
    const left = side === 'right' ? anchor.right + gap : anchor.left - gap - size.width;
    const top = clamp(centerY - tailReach, innerTop, innerBottom - size.height);
    return { left: clamp(left, innerLeft, innerRight - size.width), top, side, tail: clamp(centerY - top, tailInset, size.height - tailInset) };
  }
  const top = side === 'below' ? anchor.bottom + gap : anchor.top - gap - size.height;
  const left = clamp(centerX - size.width / 2, innerLeft, innerRight - size.width);
  return { left, top: clamp(top, innerTop, innerBottom - size.height), side, tail: clamp(centerX - left, tailInset, size.width - tailInset) };
}
