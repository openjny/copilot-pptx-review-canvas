import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { AnnotationStore } from '../lib/annotations.mjs';
import { DeckCache, resolveDeckSource } from '../lib/deck-cache.mjs';
import { UndoJournal } from '../lib/history.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makePackageFile, makeTempDir, postJson } from './helpers.mjs';

const IDS = [256, 257, 258, 259];
const pen = (id, x = 0.5) => ({ id, tool: 'pen', color: '#e5231b', width: 0.004, points: [[x, 0.5], [x + 0.2, 0.6]] });
const titles = (instance) => instance.snapshot().deck.slides.map((slide) => slide.title);
const draftsOf = (instance) => instance.snapshot().annotations.drafts.map((draft) => [draft.id, draft.after]);
const undo = (instance, kind = 'undo') => postJson(`${instance.url}api/history/${kind}`, { id: instance.snapshot().history[kind]?.id ?? 'none' });

async function until(condition, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  assert.fail('condition was not met in time');
}

async function setup({ ids = IDS } = {}) {
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  const calls = [];
  const registry = new ViewerRegistry({ root: dataRoot, watchIntervalMs: 30, renderer: fakeRenderer({ slides: ids.length, calls }) });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makePackageFile(decks, 'deck.pptx', ids);
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  return { registry, instance, file, dataRoot, calls };
}

test('UndoJournal is a linear history: recording drops what was waiting to be redone', () => {
  const released = [];
  const journal = new UndoJournal({ limit: 3, release: (entry) => released.push(entry.label) });
  assert.deepEqual(journal.view(), { undo: null, redo: null });
  const a = journal.record({ type: 'slide', label: 'a' });
  const b = journal.record({ type: 'draft', label: 'b' });
  assert.equal(journal.peek('undo').id, b.id);
  assert.equal(journal.peek('redo'), null);

  journal.commit('undo');
  assert.equal(journal.peek('undo').id, a.id);
  assert.equal(journal.peek('redo').id, b.id);
  assert.deepEqual(journal.view().redo, { id: b.id, kind: 'draft', at: b.at, label: 'b' });
  journal.commit('redo');
  assert.equal(journal.peek('redo'), null);

  journal.commit('undo');
  journal.record({ type: 'slide', label: 'c' });
  assert.deepEqual(released, ['b'], 'the undone action that was waiting is dropped');
  assert.equal(journal.peek('redo'), null);

  journal.record({ type: 'slide', label: 'd' });
  journal.record({ type: 'slide', label: 'e' });
  assert.deepEqual(released, ['b', 'a'], 'the oldest goes when the limit is passed');
  assert.equal(journal.size, 3);
  journal.clear();
  assert.deepEqual(released, ['b', 'a', 'c', 'd', 'e']);
  assert.equal(journal.size, 0);
});

test('DeckCache keeps a pinned render from being cleaned up, and cleans it when it is let go', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const cache = new DeckCache({ root, renderer: fakeRenderer({ slides: 2 }) });
  const file = await makePackageFile(decks, 'p.pptx', [256, 257]);
  const first = await cache.resolve(await resolveDeckSource(file));
  await fs.appendFile(file, Buffer.alloc(3));
  const source = await resolveDeckSource(file);
  const second = await cache.resolve(source);
  assert.notEqual(first.key, second.key);
  const old = new Date(Date.now() - 60 * 60 * 1000);
  await fs.utimes(path.join(cache.keyDir(first.key), 'manifest.json'), old, old);

  cache.pin(first.key);
  cache.pin(first.key);
  await cache._prune(source, second.key);
  await fs.access(cache.keyDir(first.key));
  cache.unpin(first.key);
  await cache._prune(source, second.key);
  await fs.access(cache.keyDir(first.key));
  cache.unpin(first.key);
  await cache._prune(source, second.key);
  await assert.rejects(() => fs.access(cache.keyDir(first.key)), 'the render goes once nobody holds it');
});

test('a deleted slide comes back with its marks, and the draft pages that followed it follow it again', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const deck = (key, hashes) => ({ key, source: { path: 'C:\\x\\deck.pptx' }, slideCount: hashes.length, slides: hashes.map((hash) => ({ hash })) });
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2', 'h3', 'h4']));
  book.add(2, pen('two'));
  book.add(3, pen('three'));
  book.add(4, pen('four'));
  book.markRead([4]);
  const stays = book.addDraft({ after: 1 });
  const follows = book.addDraft({ after: 2 });
  const later = book.addDraft({ after: 3 });
  const saved = { strokes: [...book.strokes(2)], unread: book.isUnread(2), draftIds: [follows.id] };

  await store.attach(deck('b'.repeat(24), ['h1', 'h3', 'h4']), { order: [1, 3, 4] });
  assert.deepEqual(book.drafts().map((draft) => [draft.id, draft.after]), [[stays.id, 1], [follows.id, 1], [later.id, 2]]);

  const back = await store.attach(deck('c'.repeat(24), ['h1', 'h2', 'h3', 'h4']), { insert: { slide: 2, ...saved } });
  assert.equal(back.book, book);
  assert.deepEqual(book.counts(), { 2: 1, 3: 1, 4: 1 });
  assert.deepEqual(book.strokes(2).map((stroke) => stroke.id), ['two']);
  assert.deepEqual(book.strokes(3).map((stroke) => stroke.id), ['three']);
  assert.deepEqual(book.unreadSlides(), [2, 3], 'slide 4 was read, and still is');
  assert.deepEqual(book.drafts().map((draft) => [draft.id, draft.after]), [[stays.id, 1], [follows.id, 2], [later.id, 3]]);
  assert.equal(book.slideCount, 4);
  assert.equal(book.deckKey, 'c'.repeat(24));
});

test('a deleted new slide is put back where it was, with its strokes', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach({ key: 'a'.repeat(24), source: { path: 'C:\\x\\deck.pptx' }, slideCount: 3, slides: [{ hash: 'h1' }, { hash: 'h2' }, { hash: 'h3' }] });
  const first = book.addDraft({ after: 1 });
  const second = book.addDraft({ after: 1 });
  const third = book.addDraft({ after: 2 });
  book.add(second.id, pen('s'));
  const before = book.drafts().map((draft) => draft.id);
  const index = book.draftIndex(second.id);
  const draft = book.removeDraft(second.id);
  assert.equal(book.draft(second.id), null);

  book.reinsertDraft(draft, index, true);
  assert.deepEqual(book.drafts().map((entry) => entry.id), before);
  assert.deepEqual(book.strokes(second.id).map((stroke) => stroke.id), ['s']);
  assert.deepEqual(book.unreadDrafts().map((entry) => entry.id).includes(second.id), true);
  assert.throws(() => book.reinsertDraft(draft, index, false), (error) => error.code === 'invalid_input', 'it cannot be put back twice');
  assert.ok(first && third);
});

test('deleting a slide can be undone and redone: the file, the pictures, the marks and the draft pages', async () => {
  const { registry, instance, file, dataRoot, calls } = await setup();
  try {
    const original = await fs.readFile(file);
    instance.setCurrent(2);
    instance.addStroke(2, pen('two'));
    instance.addStroke(4, pen('four'));
    const follows = instance.addDraft({ after: 2 }).draft.id;
    const later = instance.addDraft({ after: 3 }).draft.id;
    const picture = await instance.slideImage(2);
    assert.deepEqual(instance.snapshot().history, { undo: null, redo: null });

    assert.equal((await postJson(`${instance.url}api/slides/delete`, { slide: 2 })).status, 200);
    const deleted = instance.snapshot();
    assert.deepEqual(deleted.deck.slides.map((slide) => slide.title), ['Title 1', 'Title 3', 'Title 4']);
    assert.equal(deleted.history.undo.label, 'Delete slide 2');
    assert.equal(deleted.history.undo.kind, 'slide');
    assert.equal(deleted.history.redo, null);
    assert.deepEqual(draftsOf(instance), [[follows, 1], [later, 2]]);
    assert.equal((await fs.readdir(path.join(dataRoot, 'undo'))).length, 1, 'a copy of the file as it was is kept for the undo');

    const stale = await postJson(`${instance.url}api/history/undo`, { id: 'not-the-one' });
    assert.equal(stale.status, 409);
    assert.equal(json(stale).error.code, 'history_changed');

    const response = await undo(instance);
    assert.equal(response.status, 200);
    assert.deepEqual(json(response), { ok: true, kind: 'slide', label: 'Delete slide 2' });
    assert.deepEqual(await fs.readFile(file), original, 'the file is exactly as it was');
    const back = instance.snapshot();
    assert.notEqual(back.deck.key, deleted.deck.key);
    assert.deepEqual(titles(instance), ['Title 1', 'Title 2', 'Title 3', 'Title 4']);
    assert.equal(back.current, 2, 'the slide that came back is shown');
    assert.deepEqual(await instance.slideImage(2), picture);
    assert.deepEqual(back.annotations.counts, { 2: 1, 4: 1 });
    assert.deepEqual(draftsOf(instance), [[follows, 2], [later, 3]]);
    assert.equal(back.history.undo, null);
    assert.equal(back.history.redo.label, 'Delete slide 2');
    assert.deepEqual(await fs.readdir(path.join(dataRoot, 'undo')), [], 'the copy is not needed any more');
    assert.equal(calls.length, 1, 'nothing was drawn again');

    const again = await undo(instance, 'redo');
    assert.equal(again.status, 200);
    assert.deepEqual(titles(instance), ['Title 1', 'Title 3', 'Title 4']);
    assert.deepEqual(instance.snapshot().annotations.counts, { 3: 1 });
    assert.deepEqual(draftsOf(instance), [[follows, 1], [later, 2]]);
    assert.equal(instance.snapshot().history.redo, null);
    assert.equal(instance.snapshot().history.undo.label, 'Delete slide 2');

    assert.equal((await undo(instance)).status, 200, 'and it can be undone again');
    assert.deepEqual(await fs.readFile(file), original);
    await new Promise((resolve) => setTimeout(resolve, 150));
    assert.equal(instance.snapshot().phase, 'ready');
    assert.equal(calls.length, 1, 'the file watcher did not draw it again either');
  } finally {
    await registry.closeAll();
  }
});

test('an undo is refused when the file is no longer the one the deletion wrote, and the history is dropped', async () => {
  const { registry, instance, file } = await setup();
  try {
    await postJson(`${instance.url}api/slides/delete`, { slide: 3 });
    const written = await fs.readFile(file);
    await fs.appendFile(file, Buffer.alloc(3));
    const refused = await undo(instance);
    assert.equal(refused.status, 409);
    assert.equal(json(refused).error.code, 'history_conflict');
    assert.equal((await fs.readFile(file)).length, written.length + 3, 'the file was not touched');
    assert.deepEqual(instance.snapshot().history, { undo: null, redo: null });
  } finally {
    await registry.closeAll();
  }
});

test('a move ends the history, and so does a deletion that follows an undone one', async () => {
  const { registry, instance, dataRoot } = await setup({ ids: [256, 257, 258, 259, 260, 261] });
  try {
    await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    assert.ok(instance.snapshot().history.undo);
    await postJson(`${instance.url}api/slides/move`, { from: 1, to: 2 });
    assert.deepEqual(instance.snapshot().history, { undo: null, redo: null }, 'after a move');
    await until(async () => (await fs.readdir(path.join(dataRoot, 'undo'))).length === 0);

    await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    await undo(instance);
    assert.ok(instance.snapshot().history.redo);
    await postJson(`${instance.url}api/slides/delete`, { slide: 2 });
    assert.equal(instance.snapshot().history.redo, null, 'a new deletion drops what was waiting to be redone');
    await until(async () => (await fs.readdir(path.join(dataRoot, 'undo'))).length === 1);
  } finally {
    await registry.closeAll();
  }
});

test('drawing the file again ends the history', async () => {
  const { registry, instance, dataRoot } = await setup();
  try {
    await postJson(`${instance.url}api/slides/delete`, { slide: 1 });
    assert.ok(instance.snapshot().history.undo);
    await instance.refresh();
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.deepEqual(instance.snapshot().history, { undo: null, redo: null });
    await until(async () => (await fs.readdir(path.join(dataRoot, 'undo'))).length === 0);
  } finally {
    await registry.closeAll();
  }
});

test('deleting a new slide can be undone and redone too', async () => {
  const { registry, instance } = await setup();
  try {
    const first = instance.addDraft({ after: 1 }).draft.id;
    const second = instance.addDraft({ after: 1 }).draft.id;
    const third = instance.addDraft({ after: 2 }).draft.id;
    instance.addStroke(second, pen('s'));
    const before = draftsOf(instance);

    assert.equal((await postJson(`${instance.url}api/drafts/remove`, { id: second })).status, 200);
    assert.ok(!draftsOf(instance).some(([id]) => id === second));
    assert.equal(instance.snapshot().history.undo.label, 'Delete new slide');

    assert.equal((await undo(instance)).status, 200);
    assert.deepEqual(draftsOf(instance), before, 'it is back in the same place');
    assert.deepEqual(instance.annotationsFor(second).strokes.map((stroke) => stroke.id), ['s'], 'with its strokes');
    assert.equal(instance.snapshot().draft, second, 'and shown');

    assert.equal((await undo(instance, 'redo')).status, 200);
    assert.ok(!draftsOf(instance).some(([id]) => id === second));
    assert.equal((await undo(instance)).status, 200);
    assert.deepEqual(draftsOf(instance), before);
    assert.ok(first && third);
  } finally {
    await registry.closeAll();
  }
});

test('an undo after other work: undoing a deletion keeps the strokes drawn since on the slides they belong to', async () => {
  const { registry, instance, file } = await setup();
  try {
    const original = await fs.readFile(file);
    instance.addStroke(1, pen('one'));
    await postJson(`${instance.url}api/slides/delete`, { slide: 2 });
    instance.addStroke(3, pen('new'));
    assert.deepEqual(instance.snapshot().annotations.counts, { 1: 1, 3: 1 });
    await undo(instance);
    assert.deepEqual(await fs.readFile(file), original);
    assert.deepEqual(instance.snapshot().annotations.counts, { 1: 1, 4: 1 }, 'the stroke on the old slide 3 (now 4) moved down with it');
    await until(() => instance.snapshot().phase === 'ready');
  } finally {
    await registry.closeAll();
  }
});
