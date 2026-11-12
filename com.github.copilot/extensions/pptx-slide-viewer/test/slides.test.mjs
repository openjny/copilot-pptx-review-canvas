import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { AnnotationStore } from '../lib/annotations.mjs';
import { DeckCache, resolveDeckSource } from '../lib/deck-cache.mjs';
import { ViewerError } from '../lib/errors.mjs';
import { reserveNewDeckFile, resolveNewDeckTarget, suggestNewDeckPath } from '../lib/new-deck.mjs';
import { findEntry, readEntry, readZipIndex } from '../lib/pptx-zip.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makeDeckFile, makeDeckZip, makePng, makePptxBuffer, makeTempDir, postJson, request } from './helpers.mjs';

const SHAPES = Object.fromEntries([1, 2, 3, 4].map((n) => [n, [{ name: `shape-${n}`, kind: 'text', depth: 0, x: 10, y: 10, w: 100, h: 40, text: `T${n}` }]]));
const IDS = [256, 257, 258, 259];
const pen = (id, x = 0.5) => ({ id, tool: 'pen', color: '#e5231b', width: 0.004, points: [[x, 0.5], [x + 0.2, 0.6]] });
const titles = (instance) => instance.snapshot().deck.slides.map((slide) => slide.title);
const idsInFile = async (file) => {
  const buffer = await fs.readFile(file);
  const entry = findEntry(readZipIndex(buffer), 'ppt/presentation.xml');
  return [...readEntry(buffer, entry).toString('utf8').matchAll(/<p:sldId id="(\d+)"/g)].map((match) => Number(match[1]));
};

async function setup({ ids = IDS, name = 'order.pptx', watchIntervalMs = 0, creator } = {}) {
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  const calls = [];
  const registry = new ViewerRegistry({ root: dataRoot, watchIntervalMs, renderer: fakeRenderer({ slides: ids.length, calls, shapes: SHAPES }), ...(creator ? { creator } : {}) });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makeDeckZip(decks, name, ids);
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  return { registry, instance, file, decks, dataRoot, calls };
}

async function until(condition, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  assert.fail('condition was not met in time');
}

test('DeckCache.derive reuses the pictures under the new order, so no PowerPoint run is needed', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const calls = [];
  const cache = new DeckCache({ root, renderer: fakeRenderer({ slides: 4, calls, shapes: SHAPES }) });
  const file = await makeDeckZip(decks, 'd.pptx', IDS);
  const before = await cache.resolve(await resolveDeckSource(file));
  assert.equal(calls.length, 1);

  await fs.appendFile(file, Buffer.alloc(3));
  const source = await resolveDeckSource(file);
  const derived = cache.derive(before.key, source, [1, 3, 4, 2]);
  const asked = cache.resolve(source);
  assert.equal((await asked).key, (await derived).key, 'someone opening the file meanwhile waits for it instead of rendering');
  assert.equal(calls.length, 1);
  const result = await derived;
  assert.notEqual(result.key, before.key);
  assert.deepEqual(result.slides.map((slide) => [slide.n, slide.title]), [[1, 'Title 1'], [2, 'Title 3'], [3, 'Title 4'], [4, 'Title 2']]);
  assert.deepEqual(result.slides.map((slide) => slide.hash), [1, 3, 4, 2].map((old) => before.slides[old - 1].hash));
  assert.deepEqual(await fs.readFile(path.join(cache.keyDir(result.key), 's0004.png')), makePng(2));
  assert.deepEqual(await fs.readFile(path.join(cache.keyDir(result.key), 't0002.png')), makePng(3));
  const shapes = await cache.readShapes(result.key);
  assert.deepEqual(Object.keys(shapes).map((n) => shapes[n][0].name), ['shape-1', 'shape-3', 'shape-4', 'shape-2']);

  assert.equal((await cache.resolve(source)).key, result.key, 'opening the new file afterwards is a cache hit');
  assert.equal(calls.length, 1, 'nothing was rendered');
  assert.deepEqual(await fs.readdir(cache.workDir), []);

  for (const bad of [[1, 2, 3], [1, 1, 2, 3], [0, 1, 2, 3], 'abcd', null]) {
    await assert.rejects(() => cache.derive(before.key, source, bad), (error) => error instanceof ViewerError && error.code === 'invalid_input');
  }
  await assert.rejects(() => cache.derive('f'.repeat(24), source, [1, 2, 3, 4]), (error) => error.code === 'render_failed');
});

test('a shared book is rearranged in place: strokes, unread marks and draft pages follow their slides', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const deck = (key, hashes) => ({ key, source: { path: 'C:\\x\\deck.pptx' }, slideCount: hashes.length, slides: hashes.map((hash) => ({ hash })) });
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2', 'h3', 'h4']));
  book.add(2, pen('two'));
  book.add(4, pen('four'));
  book.markRead([4]);
  const afterTwo = book.addDraft({ after: 2 });
  const atTop = book.addDraft({ after: 0 });

  const moved = await store.attach(deck('b'.repeat(24), ['h1', 'h3', 'h4', 'h2']), { order: [1, 3, 4, 2] });
  assert.equal(moved.book, book, 'every viewer of the deck keeps the same book');
  assert.deepEqual(book.counts(), { 3: 1, 4: 1 });
  assert.deepEqual(book.strokes(4).map((stroke) => stroke.id), ['two']);
  assert.deepEqual(book.unreadSlides(), [4], 'unread follows the slide: the moved slide 2 was unread');
  assert.deepEqual(book.drafts().map((draft) => [draft.id, draft.after]), [[atTop.id, 0], [afterTwo.id, 4]]);
  assert.equal(book.deckKey, 'b'.repeat(24));

  const again = await store.attach(deck('b'.repeat(24), ['h1', 'h3', 'h4', 'h2']));
  assert.equal(again.book, book);
  assert.equal(again.changed, false);
  await store.flush();
  const reloaded = await new AnnotationStore(root).attach(deck('b'.repeat(24), ['h1', 'h3', 'h4', 'h2']));
  assert.deepEqual(reloaded.book.counts(), { 3: 1, 4: 1 }, 'the order survives a restart');
  assert.equal(reloaded.book.drafts().find((draft) => draft.id === afterTwo.id).after, 4);
});

test('moving a slide rewrites the file, shows the result at once and takes marks and drafts along', async () => {
  const { registry, instance, file, decks, calls, dataRoot } = await setup({ watchIntervalMs: 30 });
  try {
    const other = await registry.ensure('panel-2', { baseDir: decks });
    await other.openDeck(file);
    assert.equal((await other.whenSettled(5000)).phase, 'ready');
    instance.setCurrent(1);
    instance.addStroke(2, pen('two'));
    instance.addStroke(4, pen('four'));
    const draft = instance.addDraft({ after: 2 }).draft;
    const key = instance.snapshot().deck.key;
    const picture = await instance.slideImage(2);

    const started = Date.now();
    const response = await postJson(`${instance.url}api/slides/move`, { from: 2, to: 4 });
    assert.equal(response.status, 200);
    assert.deepEqual(json(response), { slide: 4, changed: true });
    assert.ok(Date.now() - started < 3000);

    assert.deepEqual(await idsInFile(file), [256, 258, 259, 257], 'the file itself has the new order');
    const state = instance.snapshot();
    assert.notEqual(state.deck.key, key);
    assert.deepEqual(titles(instance), ['Title 1', 'Title 3', 'Title 4', 'Title 2']);
    assert.equal(state.current, 4, 'the moved slide is shown');
    assert.deepEqual(await instance.slideImage(4), picture, 'the picture went with the slide');
    assert.deepEqual(state.annotations.counts, { 3: 1, 4: 1 });
    assert.deepEqual(instance.annotationsFor(4).strokes.map((stroke) => stroke.id), ['two']);
    assert.deepEqual(state.annotations.drafts.map((entry) => [entry.id, entry.after]), [[draft.id, 4]]);
    assert.equal(calls.length, 1, 'the deck was not drawn again');
    assert.equal((await fs.readdir(path.join(dataRoot, 'backups'))).length, 1);

    await until(() => other.snapshot().deck.key === state.deck.key && other.summary().phase === 'ready');
    assert.deepEqual(titles(other), titles(instance), 'another viewer of the same file follows');
    assert.deepEqual(other.snapshot().annotations.counts, { 3: 1, 4: 1 });
    assert.equal(calls.length, 1, 'and does not draw it again either');
    assert.equal(other._book, instance._book);

    await instance._flushPosition();
    assert.equal((await instance.positions.load(file)).slide, 4);

    assert.deepEqual(json(await postJson(`${instance.url}api/slides/move`, { from: 4, to: 2 })), { slide: 2, changed: true });
    assert.deepEqual(await idsInFile(file), IDS, 'moving it back restores the file');
    assert.deepEqual(titles(instance), ['Title 1', 'Title 2', 'Title 3', 'Title 4']);
    assert.deepEqual(instance.snapshot().annotations.counts, { 2: 1, 4: 1 });
    assert.equal(calls.length, 1);
  } finally {
    await registry.closeAll();
  }
});

test('a move that cannot be done leaves the file and the viewer as they were', async () => {
  const { registry, instance, file, decks } = await setup();
  try {
    const post = (body) => postJson(`${instance.url}api/slides/move`, body);

    assert.equal((await post({ from: 1 })).status, 400);
    assert.equal((await post({ from: 1.5, to: 2 })).status, 400);
    assert.equal((await post({ from: 1, to: 9 })).status, 400);
    assert.deepEqual(json(await post({ from: 3, to: 3 })), { slide: 3, changed: false });

    const results = await Promise.allSettled([instance.moveSlide(1, 2), instance.moveSlide(2, 3)]);
    assert.equal(results[0].status, 'fulfilled');
    assert.equal(results[1].reason.code, 'busy', 'a second move while one is running is refused');
    await instance.moveSlide(2, 1);
    assert.deepEqual(await idsInFile(file), IDS, 'moving it back restores the order');

    const locked = await fs.readFile(file);
    await fs.chmod(file, 0o444);
    const refused = await post({ from: 1, to: 2 });
    await fs.chmod(file, 0o666);
    assert.equal(refused.status, 409);
    assert.equal(json(refused).error.code, 'file_locked');
    assert.deepEqual(await fs.readFile(file), locked, 'a file that could not be written is unchanged');
    assert.deepEqual(titles(instance), ['Title 1', 'Title 2', 'Title 3', 'Title 4']);
    assert.deepEqual((await fs.readdir(decks)).filter((name) => name.endsWith('.tmp')), []);

    await fs.writeFile(file, makePptxBuffer([300, 301, 302, 303]));
    const stale = await post({ from: 1, to: 2 });
    assert.equal(stale.status, 409);
    assert.equal(json(stale).error.code, 'file_changed', 'the file was changed behind the viewer, so nothing is rewritten blindly');
    assert.deepEqual(await idsInFile(file), [300, 301, 302, 303]);
  } finally {
    await registry.closeAll();
  }
});

test('only OOXML decks can be rearranged; a sketchbook and a legacy .ppt cannot', async () => {
  const { registry, instance, decks } = await setup();
  try {
    assert.equal(instance.snapshot().deck.reorderable, true);
    const legacy = await makeDeckFile(decks, 'old.ppt', 'legacy binary');
    await instance.openDeck(legacy);
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.equal(instance.snapshot().deck.reorderable, false);
    assert.equal((await postJson(`${instance.url}api/slides/move`, { from: 1, to: 2 })).status, 400);
    await assert.rejects(() => instance.moveSlide(1, 2), (error) => error.code === 'not_reorderable');

    await instance.newSketchbook('Reorder check');
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.equal(instance.snapshot().deck.reorderable, false);
    await assert.rejects(() => instance.moveSlide(1, 1), (error) => error.code === 'not_reorderable');
  } finally {
    await registry.closeAll();
  }
});

test('a presentation can be created from the page and is opened right away', async () => {
  const made = [];
  const creator = async ({ path: input, template, baseDir }) => {
    made.push({ input, template, baseDir });
    if (/exists/.test(input)) throw new ViewerError('file_exists', 'error.fileExistsShort');
    const file = await resolveNewDeckTarget(input, baseDir);
    await fs.writeFile(file, makePptxBuffer([256]));
    return file;
  };
  const { registry, instance, decks } = await setup({ creator });
  try {
    const suggested = json(await request(`${instance.url}api/new-deck/suggest`)).path;
    assert.equal(suggested, path.join(decks, 'New deck.pptx'), 'the suggestion is in the folder the session works in');

    const response = await postJson(`${instance.url}api/new-deck`, { path: suggested, template: 'C:\\tpl\\brand.potx' });
    assert.equal(response.status, 202);
    assert.deepEqual(made[0], { input: suggested, template: 'C:\\tpl\\brand.potx', baseDir: decks });
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.equal(instance.snapshot().deck.path, suggested);
    assert.equal(instance.snapshot().deck.slideCount, 4, 'the fake renderer decides how many slides it draws');
    assert.equal(json(await request(`${instance.url}api/new-deck/suggest`)).path, path.join(decks, 'New deck (2).pptx'), 'the next suggestion is a free name');

    const clash = await postJson(`${instance.url}api/new-deck`, { path: path.join(decks, 'exists.pptx') });
    assert.equal(clash.status, 409);
    assert.equal(json(clash).error.code, 'file_exists');
    assert.equal((await postJson(`${instance.url}api/new-deck`, { path: 5 })).status, 400);
    assert.equal((await postJson(`${instance.url}api/new-deck`, { path: 'x.pptx', template: 5 })).status, 400);
    assert.equal(instance.snapshot().deck.path, suggested, 'a failed creation leaves the open deck alone');
  } finally {
    await registry.closeAll();
  }
});

test('a new presentation\'s file name is taken for good, so two creations of the same name cannot overwrite each other', async () => {
  const dir = await makeTempDir();
  const file = path.join(dir, 'race.pptx');
  const outcomes = await Promise.allSettled([reserveNewDeckFile(file), reserveNewDeckFile(file), reserveNewDeckFile(file)]);
  assert.equal(outcomes.filter((outcome) => outcome.status === 'fulfilled').length, 1, 'exactly one caller gets the name');
  for (const outcome of outcomes.filter((outcome) => outcome.status === 'rejected')) assert.equal(outcome.reason.code, 'file_exists');
  assert.equal((await fs.stat(file)).size, 0);

  const release = outcomes.find((outcome) => outcome.status === 'fulfilled').value;
  await release();
  assert.equal(await fs.access(file).then(() => true, () => false), false, 'a failed creation gives the name back');

  const again = await reserveNewDeckFile(file);
  await fs.writeFile(file, makePptxBuffer([256]));
  await again();
  assert.ok((await fs.stat(file)).size > 0, 'a file that was filled in is never removed');
  await assert.rejects(() => reserveNewDeckFile(file), (error) => error.code === 'file_exists');
});

test('the place for a new presentation is checked before PowerPoint is started', async () => {
  const dir = await makeTempDir();
  const code = async (input, base = dir) => resolveNewDeckTarget(input, base).then(() => 'ok', (error) => error.code);
  assert.equal(await resolveNewDeckTarget(path.join(dir, 'a'), dir), path.join(dir, 'a.pptx'), 'the extension is added');
  assert.equal(await resolveNewDeckTarget(`"${path.join(dir, 'b.PPTX')}"`, dir), path.join(dir, 'b.PPTX'), 'quotes from a pasted path are removed');
  assert.equal(await resolveNewDeckTarget('sub.pptx', dir), path.join(dir, 'sub.pptx'), 'relative to the session folder');
  assert.equal(await code('', dir), 'invalid_path');
  assert.equal(await code('rel.pptx', null), 'invalid_path');
  assert.equal(await code(path.join(dir, 'x.docx')), 'unsupported_file');
  assert.equal(await code(path.join(dir, 'x.ppt')), 'unsupported_file');
  assert.equal(await code(path.join(dir, 'missing', 'x.pptx')), 'folder_not_found');
  await fs.writeFile(path.join(dir, 'taken.pptx'), 'x');
  assert.equal(await code(path.join(dir, 'taken.pptx')), 'file_exists');
  assert.equal(await code(path.join(dir, 'taken')), 'file_exists', 'the same goes for a name given without the extension');

  assert.equal(await suggestNewDeckPath(dir), path.join(dir, 'New deck.pptx'));
  await fs.writeFile(path.join(dir, 'New deck.pptx'), 'x');
  await fs.writeFile(path.join(dir, 'New deck (2).pptx'), 'x');
  assert.equal(await suggestNewDeckPath(dir), path.join(dir, 'New deck (3).pptx'));
});
