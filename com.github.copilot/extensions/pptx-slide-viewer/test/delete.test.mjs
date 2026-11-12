import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { AnnotationStore } from '../lib/annotations.mjs';
import { DeckCache, resolveDeckSource } from '../lib/deck-cache.mjs';
import { ViewerError } from '../lib/errors.mjs';
import { deleteSlideInFile, removeSlideFromPresentationXml } from '../lib/pptx-delete.mjs';
import { findEntry, readEntry, readZipIndex, removeEntries, validateZip } from '../lib/pptx-zip.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makePackageBuffer, makePackageFile, makePng, makeTempDir, makeZip, postJson, presentationXml } from './helpers.mjs';

const IDS = [256, 257, 258, 259];
const pen = (id, x = 0.5) => ({ id, tool: 'pen', color: '#e5231b', width: 0.004, points: [[x, 0.5], [x + 0.2, 0.6]] });
const names = (buffer) => readZipIndex(buffer).entries.map((entry) => entry.name);
const textOf = (buffer, name) => readEntry(buffer, findEntry(readZipIndex(buffer), name)).toString('utf8');
const rejects = (promise, code) => assert.rejects(promise, (error) => error instanceof ViewerError && error.code === code);

test('removeEntries drops the named entries and copies the others as they are', () => {
  const entries = ['a', 'b', 'c', 'd'].map((name, i) => ({ name: `${name}.txt`, data: `${name}`.repeat(40 + i), method: i % 2 ? 0 : 8, descriptor: i % 2 === 0 }));
  const zip = makeZip(entries, { comment: 'kept' });
  const output = removeEntries(zip, readZipIndex(zip), ['b.txt', 'd.txt']);
  validateZip(output);
  assert.deepEqual(names(output), ['a.txt', 'c.txt']);
  assert.equal(textOf(output, 'a.txt'), 'a'.repeat(40));
  assert.equal(textOf(output, 'c.txt'), 'c'.repeat(42));
  assert.equal(readZipIndex(output).eocdLength, 22 + 4, 'the ZIP comment stays');
  assert.deepEqual(names(removeEntries(zip, readZipIndex(zip), [])), names(zip));
});

test('removeSlideFromPresentationXml takes the slide out of the list, the sections and the custom shows', () => {
  const sections = [{ name: 'One', ids: [256, 257] }, { name: 'Two', ids: [258] }, { name: 'Three', ids: [259] }];
  const result = removeSlideFromPresentationXml(presentationXml(IDS, { sections, customShows: [[1, 3], [2]] }), 3);
  assert.equal(result.rId, 'rId12');
  assert.deepEqual(result.order, [1, 2, 4]);
  assert.deepEqual([...result.xml.matchAll(/<p:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256, 257, 259]);
  assert.deepEqual([...result.xml.matchAll(/<p14:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256, 257, 259]);
  assert.match(result.xml, /name="Two"[^>]*><p14:sldIdLst\/>/, 'a section left without slides stays, empty');
  assert.deepEqual([...result.xml.matchAll(/<p:sld r:id="(rId\d+)"/g)].map((match) => match[1]), ['rId10', 'rId11'], 'custom shows no longer name it');
  assert.ok(!result.xml.includes('rId12'));

  const first = removeSlideFromPresentationXml(presentationXml(IDS), 1);
  assert.deepEqual(first.order, [2, 3, 4]);
  assert.ok(!/id="256"/.test(first.xml));
  const prefixed = removeSlideFromPresentationXml(presentationXml(IDS, { prefix: 'x' }), 2);
  assert.deepEqual([...prefixed.xml.matchAll(/<x:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256, 258, 259]);
});

test('removeSlideFromPresentationXml refuses what it cannot do safely', () => {
  assert.throws(() => removeSlideFromPresentationXml(presentationXml([256]), 1), (error) => error.code === 'last_slide');
  for (const bad of [0, 5, 1.5, '2']) assert.throws(() => removeSlideFromPresentationXml(presentationXml(IDS), bad), (error) => error.code === 'slide_out_of_range');
  assert.throws(() => removeSlideFromPresentationXml('<p:presentation/>', 1), (error) => error.code === 'invalid_pptx');
  const broken = presentationXml(IDS, { sections: [{ name: 'One', ids: [256, 257, 258] }] });
  assert.throws(() => removeSlideFromPresentationXml(broken, 4), (error) => error.code === 'unsupported_pptx', 'a slide that no section lists');
});

test('deleteSlideInFile removes the slide and what only it used, and keeps the rest and a backup', async () => {
  const dir = await makeTempDir();
  const backupDir = path.join(dir, 'backups');
  const sections = [{ name: 'One', ids: [256, 257] }, { name: 'Two', ids: [258, 259] }];
  const file = await makePackageFile(dir, 'deck.pptx', IDS, { sections, customShows: [[2, 4]] });
  const original = await fs.readFile(file);

  const result = await deleteSlideInFile(file, 2, { backupDir });
  assert.deepEqual(result.order, [1, 3, 4]);
  assert.equal(result.removedParts, 5);
  const output = await fs.readFile(file);
  validateZip(output);
  const gone = ['ppt/slides/slide257.xml', 'ppt/slides/_rels/slide257.xml.rels', 'ppt/notesSlides/notesSlide257.xml', 'ppt/notesSlides/_rels/notesSlide257.xml.rels', 'ppt/media/own257.png'];
  const left = names(output);
  for (const name of gone) assert.ok(!left.includes(name), `${name} is gone`);
  assert.deepEqual(names(original).filter((name) => !gone.includes(name)), left, 'every other part is still there, in the same order');
  for (const name of left) {
    if (['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'].includes(name)) continue;
    assert.deepEqual(readEntry(output, findEntry(readZipIndex(output), name)), readEntry(original, findEntry(readZipIndex(original), name)), `${name} is unchanged`);
  }
  assert.ok(left.includes('ppt/media/shared.png') && left.includes('ppt/slideLayouts/slideLayout1.xml'), 'what other slides share stays');

  const presentation = textOf(output, 'ppt/presentation.xml');
  assert.deepEqual([...presentation.matchAll(/<p:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256, 258, 259]);
  assert.deepEqual([...presentation.matchAll(/<p:sld r:id="(rId\d+)"/g)].map((match) => match[1]), ['rId13'], 'the custom show keeps its other slide');
  assert.ok(!textOf(output, 'ppt/_rels/presentation.xml.rels').includes('slide257.xml'));
  const types = textOf(output, '[Content_Types].xml');
  assert.ok(!types.includes('slide257.xml') && !types.includes('notesSlide257.xml'));
  assert.ok(types.includes('/ppt/slides/slide256.xml') && types.includes('/ppt/slides/slide258.xml') && types.includes('Default Extension="png"'));

  assert.deepEqual(await fs.readFile(path.join(backupDir, (await fs.readdir(backupDir))[0])), original, 'the file as it was is kept');
});

test('deleteSlideInFile refuses a slide that another slide links to, and changes nothing', async () => {
  const dir = await makeTempDir();
  const file = await makePackageFile(dir, 'linked.pptx', IDS, { links: [[256, 258]] });
  const original = await fs.readFile(file);
  await rejects(deleteSlideInFile(file, 3, { backupDir: path.join(dir, 'backups') }), 'slide_linked');
  assert.deepEqual(await fs.readFile(file), original);
  assert.equal((await deleteSlideInFile(file, 2)).changed, true, 'a slide nobody links to can go');
});

test('deleteSlideInFile refuses the last slide, legacy files and files that are not a presentation', async () => {
  const dir = await makeTempDir();
  const single = await makePackageFile(dir, 'single.pptx', [256]);
  await rejects(deleteSlideInFile(single, 1), 'last_slide');
  const legacy = path.join(dir, 'old.ppt');
  await fs.writeFile(legacy, 'x');
  await rejects(deleteSlideInFile(legacy, 1), 'not_reorderable');
  const bare = path.join(dir, 'bare.pptx');
  await fs.writeFile(bare, makeZip([{ name: 'readme.txt', data: 'hello' }]));
  await rejects(deleteSlideInFile(bare, 1), 'invalid_pptx');
  const noRels = path.join(dir, 'norels.pptx');
  await fs.writeFile(noRels, makeZip([{ name: 'ppt/presentation.xml', data: presentationXml(IDS) }]));
  await rejects(deleteSlideInFile(noRels, 1), 'invalid_pptx');
});

test('DeckCache.derive keeps only the pictures of the slides that are left', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const cache = new DeckCache({ root, renderer: fakeRenderer({ slides: 4, calls: [] }) });
  const file = await makePackageFile(decks, 'd.pptx', IDS);
  const before = await cache.resolve(await resolveDeckSource(file));
  await fs.appendFile(file, Buffer.alloc(3));
  const source = await resolveDeckSource(file);

  const derived = await cache.derive(before.key, source, [1, 3, 4], { dropping: true });
  assert.equal(derived.slideCount, 3);
  assert.deepEqual(derived.slides.map((slide) => [slide.n, slide.title]), [[1, 'Title 1'], [2, 'Title 3'], [3, 'Title 4']]);
  assert.deepEqual(await fs.readFile(path.join(cache.keyDir(derived.key), 's0002.png')), makePng(3));
  await assert.rejects(() => fs.access(path.join(cache.keyDir(derived.key), 's0004.png')));

  for (const bad of [[1, 2, 3, 4], [1, 1, 2], [0, 1], [], [5], 'abc', null]) {
    await rejects(cache.derive(before.key, source, bad, { dropping: true }), 'invalid_input');
  }
  await rejects(cache.derive(before.key, source, [1, 3, 4]), 'invalid_input');
});

test('a deleted slide takes its marks to the archive; the others, and the draft pages, move up', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const deck = (key, hashes) => ({ key, source: { path: 'C:\\x\\deck.pptx' }, slideCount: hashes.length, slides: hashes.map((hash) => ({ hash })) });
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2', 'h3', 'h4']));
  book.add(2, pen('two'));
  book.add(3, pen('three'));
  book.add(4, pen('four'));
  book.markRead([4]);
  const afterKept = book.addDraft({ after: 3 });
  const afterDeleted = book.addDraft({ after: 2 });
  const top = book.addDraft({ after: 0 });
  const end = book.addDraft({ after: 4 });

  const result = await store.attach(deck('b'.repeat(24), ['h1', 'h3', 'h4']), { order: [1, 3, 4] });
  assert.equal(result.book, book);
  assert.equal(result.archived, 1);
  assert.deepEqual(book.counts(), { 2: 1, 3: 1 });
  assert.deepEqual(book.strokes(2).map((stroke) => stroke.id), ['three']);
  assert.deepEqual(book.strokes(3).map((stroke) => stroke.id), ['four']);
  assert.deepEqual(book.unreadSlides(), [2], 'unread follows the slide: the old slide 3 is slide 2 now, and slide 4 was read');
  assert.deepEqual(
    book.drafts().map((draft) => [draft.id, draft.after]),
    [[top.id, 0], [afterDeleted.id, 1], [afterKept.id, 2], [end.id, 3]],
    'a draft that followed the deleted slide follows the slide before it, and stays ahead of the draft that followed the next slide',
  );
  assert.equal(book.slideCount, 3);

  const archive = path.join(root, 'annotations', 'archive');
  const [file] = await fs.readdir(archive);
  const archived = JSON.parse(await fs.readFile(path.join(archive, file), 'utf8'));
  assert.deepEqual(Object.keys(archived.data.slides), ['2']);
  assert.deepEqual(archived.data.slides['2'].strokes.map((stroke) => stroke.id), ['two']);
});

async function setup({ ids = IDS, name = 'deck.pptx', options, watchIntervalMs = 30 } = {}) {
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  const calls = [];
  const registry = new ViewerRegistry({ root: dataRoot, watchIntervalMs, renderer: fakeRenderer({ slides: ids.length, calls }) });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makePackageFile(decks, name, ids, options);
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  return { registry, instance, file, dataRoot, calls };
}

async function until(condition, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  assert.fail('condition was not met in time');
}

test('deleting a slide rewrites the file, shows the result at once and moves marks and drafts', async () => {
  const { registry, instance, file, dataRoot, calls } = await setup();
  try {
    const other = await registry.ensure('panel-2', { baseDir: path.dirname(file) });
    await other.openDeck(file);
    assert.equal((await other.whenSettled(5000)).phase, 'ready');
    instance.setCurrent(2);
    instance.addStroke(2, pen('two'));
    instance.addStroke(4, pen('four'));
    const draft = instance.addDraft({ after: 3 }).draft.id;
    const key = instance.snapshot().deck.key;
    const third = await instance.slideImage(3);

    const response = await postJson(`${instance.url}api/slides/delete`, { slide: 2 });
    assert.equal(response.status, 200);
    const body = json(response);
    assert.deepEqual([body.slide, body.slideCount, body.changed], [2, 3, true]);

    const output = await fs.readFile(file);
    assert.deepEqual([...textOf(output, 'ppt/presentation.xml').matchAll(/<p:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256, 258, 259]);
    const state = instance.snapshot();
    assert.notEqual(state.deck.key, key);
    assert.equal(state.deck.slideCount, 3);
    assert.deepEqual(state.deck.slides.map((slide) => slide.title), ['Title 1', 'Title 3', 'Title 4']);
    assert.equal(state.current, 2, 'the slide that followed takes its place on screen');
    assert.deepEqual(await instance.slideImage(2), third);
    assert.deepEqual(state.annotations.counts, { 3: 1 });
    assert.deepEqual(state.annotations.drafts.map((entry) => [entry.id, entry.after]), [[draft, 2]]);

    await until(() => other.snapshot().deck.key === state.deck.key && other.summary().phase === 'ready');
    const sibling = other.snapshot();
    assert.equal(sibling.deck.slideCount, 3, 'the other viewer of the same file shows it too');
    assert.deepEqual(sibling.annotations.counts, { 3: 1 });
    assert.equal(calls.length, 1, 'and does not draw it again');
    assert.equal((await fs.readdir(path.join(dataRoot, 'annotations', 'archive'))).length, 1, 'the marks of the deleted slide were set aside');
    assert.equal((await fs.readdir(path.join(dataRoot, 'backups'))).length, 1);
    await new Promise((resolve) => setTimeout(resolve, 150));
    assert.equal(instance.snapshot().phase, 'ready');
    assert.equal(instance.snapshot().deck.key, state.deck.key, 'the file watcher does not render it again');
  } finally {
    await registry.closeAll();
  }
});

test('deleting the last slide of the list, or the only slide, is refused and nothing changes', async () => {
  const { registry, instance, file } = await setup({ ids: [256, 257] });
  try {
    instance.setCurrent(2);
    const response = await postJson(`${instance.url}api/slides/delete`, { slide: 2 });
    assert.equal(response.status, 200);
    assert.equal(instance.snapshot().current, 1, 'the slide before it is shown when the last one goes');
    const again = await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    assert.equal(again.status, 400);
    assert.equal(json(again).error.code, 'last_slide');
    assert.equal(instance.snapshot().deck.slideCount, 1);
    assert.deepEqual([...textOf(await fs.readFile(file), 'ppt/presentation.xml').matchAll(/<p:sldId id="(\d+)"/g)].map((match) => Number(match[1])), [256]);
    const bad = await postJson(`${instance.url}api/slides/delete`, { slide: 'one' });
    assert.equal(bad.status, 400);
  } finally {
    await registry.closeAll();
  }
});

test('a file that changed behind the viewer, or a sketchbook, cannot have slides deleted', async () => {
  // No file watcher: it would draw the changed file again and could win against opening the sketchbook.
  const { registry, instance, file } = await setup({ watchIntervalMs: 0 });
  try {
    const before = await fs.readFile(file);
    await fs.appendFile(file, Buffer.alloc(3));
    const stale = await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    assert.equal(stale.status, 409);
    assert.equal(json(stale).error.code, 'file_changed');
    assert.equal((await fs.readFile(file)).length, before.length + 3);

    await instance.newSketchbook('delete-test');
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    const sketch = await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    assert.equal(sketch.status, 400);
    assert.equal(json(sketch).error.code, 'not_reorderable');
  } finally {
    await registry.closeAll();
  }
});
