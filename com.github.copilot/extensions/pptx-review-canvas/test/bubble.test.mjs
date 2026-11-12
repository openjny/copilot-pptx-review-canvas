import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { placeBubble } from '../assets/place.js';
import { RAIL_HEIGHT, COMMENTS_WIDTH, clampSize } from '../assets/resize.js';
import { shapeKey, shapePoints } from '../assets/shape.js';

const bounds = { left: 0, top: 0, right: 1000, bottom: 600 };
const size = { width: 300, height: 140 };
const point = (x, y, pad = 14) => ({ left: x - pad, top: y - pad, right: x + pad, bottom: y + pad });

test('the bubble goes to the right of the spot when there is room', () => {
  const spot = placeBubble({ anchor: point(200, 300), size, bounds });
  assert.equal(spot.side, 'right');
  assert.equal(spot.left, 214 + 14);
  assert.ok(spot.tail >= 18 && spot.tail <= size.height - 18);
  // the tail points at the spot: its distance from the top of the bubble is the spot's distance from it
  assert.equal(spot.top + spot.tail, 300);
});

test('it flips to the left, then below, then above when the side before it has no room', () => {
  assert.equal(placeBubble({ anchor: point(900, 300), size, bounds }).side, 'left');
  const narrow = { left: 0, top: 0, right: 320, bottom: 600 };
  assert.equal(placeBubble({ anchor: point(160, 100), size, bounds: narrow }).side, 'below');
  assert.equal(placeBubble({ anchor: point(160, 520), size, bounds: narrow }).side, 'above');
});

test('the bubble stays inside the bounds, and the tail keeps off the corners', () => {
  const nearTop = placeBubble({ anchor: point(200, 10), size, bounds });
  assert.equal(nearTop.top, 8);
  assert.equal(nearTop.tail, 18);
  const nearBottom = placeBubble({ anchor: point(200, 595), size, bounds });
  assert.equal(nearBottom.top + size.height, bounds.bottom - 8);
  assert.equal(nearBottom.tail, size.height - 18);
  const nearEdge = placeBubble({ anchor: point(160, 100), size, bounds: { left: 0, top: 0, right: 320, bottom: 600 } });
  assert.ok(nearEdge.left >= 8 && nearEdge.left + size.width <= 312);
});

test('for an area the bubble sits beside the area, not on it', () => {
  const area = { left: 300, top: 200, right: 500, bottom: 400 };
  const spot = placeBubble({ anchor: area, size, bounds });
  assert.equal(spot.side, 'right');
  assert.ok(spot.left >= area.right + 14);
});

test('when no side has room it still returns a place inside the bounds', () => {
  const tiny = { left: 0, top: 0, right: 200, bottom: 120 };
  const spot = placeBubble({ anchor: point(100, 60), size, bounds: tiny });
  assert.ok(Number.isFinite(spot.left) && Number.isFinite(spot.top) && Number.isFinite(spot.tail));
});

test('Shift draws a straight line and Ctrl a rectangle, Ctrl winning when both are held', () => {
  assert.equal(shapeKey({}), null);
  assert.equal(shapeKey({ shiftKey: true }), 'line');
  assert.equal(shapeKey({ ctrlKey: true }), 'rectangle');
  assert.equal(shapeKey({ metaKey: true }), 'rectangle');
  assert.equal(shapeKey({ shiftKey: true, ctrlKey: true }), 'rectangle');
});

test('a shape runs from where the stroke began to where the pointer is, however it wandered', () => {
  const free = [[0.1, 0.2], [0.3, 0.5], [0.2, 0.4]];
  const start = [0.1, 0.2];
  const end = [0.6, 0.5];
  assert.equal(shapePoints(null, free, start, end), free);
  assert.deepEqual(shapePoints('line', free, start, end), [start, end]);
  assert.deepEqual(shapePoints('rectangle', free, start, end), [[0.1, 0.2], [0.6, 0.2], [0.6, 0.5], [0.1, 0.5], [0.1, 0.2]]);
  // dragging up and to the left gives the same closed box
  assert.deepEqual(shapePoints('rectangle', free, [0.6, 0.5], [0.1, 0.2]).slice(0, 3), [[0.6, 0.5], [0.1, 0.5], [0.1, 0.2]]);
});

test('a shape with no size is a single point, so a click with a key held still leaves a dot', () => {
  const start = [0.4, 0.4];
  assert.deepEqual(shapePoints('line', [start], start, [0.4005, 0.4], { minStep: 0.002 }), [start]);
  assert.deepEqual(shapePoints('rectangle', [start], start, start, { minStep: 0.002 }), [start]);
});

test('a size dragged by the user stays between the smallest and what the window allows', () => {
  // the comments pane: at most 60% of the width of the layout and 720px
  assert.equal(clampSize(100, 1280), COMMENTS_WIDTH.min);
  assert.equal(clampSize(500, 1280), 500);
  assert.equal(clampSize(900, 1280), 720);
  assert.equal(clampSize(900, 1000), 600, 'a narrow window keeps room for the slide');
  assert.equal(clampSize(900, 300), COMMENTS_WIDTH.min, 'a window too small for the share still gets the smallest size');
  // the strip of thumbnails: at most 45% of the height of the layout and 360px
  assert.equal(clampSize(20, 800, RAIL_HEIGHT), RAIL_HEIGHT.min);
  assert.equal(clampSize(150.4, 800, RAIL_HEIGHT), 150);
  assert.equal(clampSize(500, 800, RAIL_HEIGHT), 360);
  assert.equal(clampSize(500, 600, RAIL_HEIGHT), 270, 'a short window keeps room for the slide');
});

test('every file in assets/ is served by the viewer', () => {
  const assets = fileURLToPath(new URL('../assets/', import.meta.url));
  const source = fs.readFileSync(fileURLToPath(new URL('../lib/viewer-instance.mjs', import.meta.url)), 'utf8');
  const listed = new Set([...source.matchAll(/^\s+'([\w/.-]+\.(?:js|css|html))':/gm)].map((match) => match[1]));
  const files = [];
  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isDirectory()) walk(path.join(dir, entry.name));
      else if (/\.(js|css|html)$/.test(entry.name)) files.push(path.relative(assets, path.join(dir, entry.name)).replaceAll('\\', '/'));
    }
  };
  walk(assets);
  assert.deepEqual(files.filter((file) => !listed.has(file)), []);
});
