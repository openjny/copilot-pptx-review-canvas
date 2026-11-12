import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { DeckCache, deckKey, resolveDeckSource } from '../lib/deck-cache.mjs';
import { ViewerError } from '../lib/errors.mjs';
import { fakeRenderer, makeDeckFile, makeTempDir } from './helpers.mjs';

async function rejectsWith(promise, code) {
  await assert.rejects(promise, (error) => error instanceof ViewerError && error.code === code);
}

test('resolveDeckSource validates the path, the extension and the file', async () => {
  const dir = await makeTempDir();
  const file = await makeDeckFile(dir);
  await rejectsWith(resolveDeckSource('  '), 'invalid_path');
  await rejectsWith(resolveDeckSource('relative.pptx'), 'invalid_path');
  await rejectsWith(resolveDeckSource(path.join(dir, 'notes.txt')), 'unsupported_file');
  await rejectsWith(resolveDeckSource(path.join(dir, 'missing.pptx')), 'file_not_found');
  await fs.mkdir(path.join(dir, 'folder.pptx'));
  await rejectsWith(resolveDeckSource(path.join(dir, 'folder.pptx')), 'file_not_found');

  const source = await resolveDeckSource(file);
  assert.equal(source.name, 'deck.pptx');
  assert.equal(source.extension, '.pptx');
  assert.equal((await resolveDeckSource(`"${file}"`)).path, source.path);
  assert.equal((await resolveDeckSource('deck.pptx', dir)).path, source.path);
});

test('deckKey follows the identity of the file', async () => {
  const base = { path: path.join(path.resolve('/x'), 'a.pptx'), size: 10, mtimeMs: 1000 };
  assert.match(deckKey(base), /^[0-9a-f]{24}$/);
  assert.equal(deckKey(base), deckKey({ ...base }));
  assert.notEqual(deckKey(base), deckKey({ ...base, size: 11 }));
  assert.notEqual(deckKey(base), deckKey({ ...base, mtimeMs: 1001 }));
});

test('DeckCache renders once, then serves the cache, and renders again when forced or damaged', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const source = await resolveDeckSource(await makeDeckFile(decks));
  const calls = [];
  const cache = new DeckCache({ root, renderer: fakeRenderer({ slides: 3, calls }) });
  const progress = [];

  const first = await cache.resolve(source, { onProgress: (done, total) => progress.push([done, total]) });
  assert.equal(first.slideCount, 3);
  assert.equal(first.slides[0].title, 'Title 1');
  assert.equal(first.slides[0].notes, 'Notes 1\nsecond line');
  assert.deepEqual(progress.at(-1), [3, 3]);
  assert.equal(calls.length, 1);
  assert.equal(path.basename(calls[0].pptx), 'source.pptx');

  const again = await cache.resolve(source);
  assert.equal(again.key, first.key);
  assert.equal(calls.length, 1);

  await cache.resolve(source, { force: true });
  assert.equal(calls.length, 2);

  await fs.rm(path.join(cache.keyDir(first.key), 't0002.png'));
  await cache.resolve(source);
  assert.equal(calls.length, 3);

  await fs.writeFile(path.join(cache.keyDir(first.key), 'manifest.json'), '{ broken');
  await cache.resolve(source);
  assert.equal(calls.length, 4);

  assert.deepEqual((await fs.readdir(cache.workDir)).filter(Boolean), []);
});

test('DeckCache shares one render between concurrent callers and leaves nothing behind on failure', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const source = await resolveDeckSource(await makeDeckFile(decks));
  const calls = [];
  const cache = new DeckCache({ root, renderer: fakeRenderer({ calls, delayMs: 5 }) });
  const [one, two] = await Promise.all([cache.resolve(source), cache.resolve(source)]);
  assert.equal(one.key, two.key);
  assert.equal(calls.length, 1);

  const failing = new DeckCache({ root: await makeTempDir(), renderer: fakeRenderer({ failWith: new ViewerError('render_failed', 'boom') }) });
  await rejectsWith(failing.resolve(source), 'render_failed');
  assert.deepEqual(await fs.readdir(failing.workDir), []);
  assert.deepEqual(await fs.readdir(failing.decksDir).catch(() => []), []);
});

test('DeckCache keeps a shared render running while any waiter is left, and cancels it when the last one gives up', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const source = await resolveDeckSource(await makeDeckFile(decks));
  const calls = [];
  const cache = new DeckCache({ root, renderer: fakeRenderer({ calls, slides: 4, delayMs: 30 }) });

  const first = new AbortController();
  const second = new AbortController();
  const progress = [];
  const leaving = cache.resolve(source, { signal: first.signal });
  const staying = cache.resolve(source, { signal: second.signal, onProgress: (done, total) => progress.push([done, total]) });
  await new Promise((resolve) => setTimeout(resolve, 20));
  first.abort();
  await rejectsWith(leaving, 'cancelled');
  assert.equal((await staying).slideCount, 4, 'the render is not cancelled for the waiter that stayed');
  assert.equal(calls.length, 1);
  assert.ok(progress.some(([done]) => done > 0), 'progress reaches the waiter that joined later');

  const other = await resolveDeckSource(await makeDeckFile(decks, 'other.pptx'));
  const lone = new AbortController();
  const abandoned = cache.resolve(other, { signal: lone.signal });
  await new Promise((resolve) => setTimeout(resolve, 20));
  lone.abort();
  await rejectsWith(abandoned, 'cancelled');
  assert.equal((await cache.resolve(other)).slideCount, 4, 'a later caller starts a fresh render instead of joining the cancelled one');
  assert.deepEqual(await fs.readdir(cache.workDir), []);
});

test('DeckCache drops the superseded render of a rebuilt file once it is old enough', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const file = await makeDeckFile(decks);
  const cache = new DeckCache({ root, renderer: fakeRenderer({ slides: 2 }) });
  const oldSource = await resolveDeckSource(file);
  const old = await cache.resolve(oldSource);
  const aged = new Date(Date.now() - 60 * 60 * 1000);
  await fs.utimes(path.join(cache.keyDir(old.key), 'manifest.json'), aged, aged);

  await fs.writeFile(file, 'rebuilt deck with different size');
  const fresh = await cache.resolve(await resolveDeckSource(file));
  assert.notEqual(fresh.key, old.key);
  const end = Date.now() + 3000;
  let remaining = await fs.readdir(cache.decksDir);
  while (remaining.length > 1 && Date.now() < end) {
    await new Promise((resolve) => setTimeout(resolve, 20));
    remaining = await fs.readdir(cache.decksDir);
  }
  assert.deepEqual(remaining, [fresh.key]);
});
