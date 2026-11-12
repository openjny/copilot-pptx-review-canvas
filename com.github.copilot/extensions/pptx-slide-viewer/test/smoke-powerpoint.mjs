// Opt-in check against the real PowerPoint: node test/smoke-powerpoint.mjs <deck.pptx>
// It is not named *.test.mjs so that `node --test` does not start PowerPoint.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { DeckCache, resolveDeckSource } from '../lib/deck-cache.mjs';
import { renderWithPowerPoint } from '../lib/powerpoint-renderer.mjs';

const target = process.argv[2];
if (!target) {
  console.error('usage: node test/smoke-powerpoint.mjs <deck.pptx>');
  process.exit(2);
}

const root = await fs.mkdtemp(path.join(os.tmpdir(), 'pptx-viewer-smoke-'));
const cache = new DeckCache({ root, renderer: renderWithPowerPoint, log: (message) => console.log(`log: ${message}`) });
const source = await resolveDeckSource(target);
const started = Date.now();
let last = '';
const deck = await cache.resolve(source, {
  onProgress: (done, total) => {
    const text = `${done}/${total}`;
    if (text !== last && (done === 0 || done === total || done % 10 === 0)) console.log(`progress ${text}`);
    last = text;
  },
});
console.log(`rendered ${deck.slideCount} slides in ${((Date.now() - started) / 1000).toFixed(1)} s (${deck.slideWidth} x ${deck.slideHeight})`);
console.log(`first title: ${deck.slides[0].title}`);
console.log(`slide 4 title: ${deck.slides[3]?.title}`);
console.log(`slide 4 notes: ${deck.slides[3]?.notes.split('\n')[0]}`);
const image = await fs.stat(path.join(cache.keyDir(deck.key), 's0004.png'));
console.log(`s0004.png: ${image.size} bytes`);
const again = Date.now();
await cache.resolve(source);
console.log(`cache hit in ${Date.now() - again} ms`);
await fs.rm(root, { recursive: true, force: true });
