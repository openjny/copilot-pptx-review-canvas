import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { ViewerError } from '../lib/errors.mjs';
import { decodePng } from '../lib/png.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makeDeckFile, makeTempDir, postJson, request } from './helpers.mjs';

const SHAPES = {
  2: [
    { name: 'slide/title', kind: 'text', depth: 0, x: 100, y: 30, w: 260, h: 100, text: 'Harness' },
    { name: 'slide/body', kind: 'text', depth: 0, x: 100, y: 300, w: 400, h: 120, text: 'Local and Copilot' },
  ],
};

const pen = (id, points, extra = {}) => ({ id, tool: 'pen', color: '#e5231b', width: 0.004, points, ...extra });
const ring = (id, cx, cy, r = 0.05) => pen(id, Array.from({ length: 25 }, (_, i) => [cx + r * Math.cos((i / 24) * 2 * Math.PI), cy + r * Math.sin((i / 24) * 2 * Math.PI) * 1.6]));

/** A registry whose renderer can be told to render the deck differently, as if it had been edited. */
async function setup({ watchIntervalMs = 0, root, render = {} } = {}) {
  const dataRoot = root ?? (await makeTempDir());
  const decks = await makeTempDir();
  const edit = { slides: 3, changed: [], variant: 0, shapes: SHAPES, ...render };
  const renderer = (args) => fakeRenderer(edit)(args);
  const registry = new ViewerRegistry({ root: dataRoot, renderer, watchIntervalMs });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makeDeckFile(decks, 'review.pptx', 'v1');
  return { registry, instance, edit, file, decks, dataRoot };
}

async function openAndWait(instance, file) {
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
}

async function until(condition, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  assert.fail('condition was not met in time');
}

test('opening the same deck twice in a row ends ready, not stuck rendering', async () => {
  const { registry, instance, file } = await setup({ render: { slides: 4, delayMs: 30 } });
  try {
    await instance.openDeck(file);
    await new Promise((resolve) => setTimeout(resolve, 20));
    await instance.openDeck(file);
    const summary = await instance.whenSettled(5000);
    assert.equal(summary.phase, 'ready');
    assert.equal(summary.slideCount, 4);
  } finally {
    await registry.closeAll();
  }
});

test('red pen is stored through the page API and rejected when invalid or too early', async () => {
  const { registry, instance, file } = await setup();
  try {
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { slide: 1, stroke: ring('a', 0.5, 0.5) })).status, 409);
    await openAndWait(instance, file);

    const added = await postJson(`${instance.url}api/annotations/add`, { slide: 2, stroke: ring('a', 0.3, 0.3) });
    assert.equal(added.status, 200);
    assert.equal(json(added).count, 1);
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { slide: 2, stroke: ring('a', 0.3, 0.3) })).status, 200, 'the same id twice is not an error');
    assert.equal(json(await request(`${instance.url}api/annotations?slide=2`)).strokes.length, 1);

    const state = json(await request(`${instance.url}api/state`));
    assert.deepEqual(state.annotations.counts, { 2: 1 });
    assert.deepEqual(state.annotations.unread, [2]);
    assert.equal(instance.outline().slides[1].strokes, 1);

    assert.equal((await postJson(`${instance.url}api/annotations/add`, { slide: 9, stroke: ring('b', 0.5, 0.5) })).status, 400);
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { slide: 1, stroke: { ...ring('b', 0.5, 0.5), color: 'red' } })).status, 400);
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { slide: 1, stroke: { ...ring('b', 0.5, 0.5), points: [] } })).status, 400);
    assert.equal((await request(`${instance.url}api/annotations?slide=x`)).status, 400);

    assert.equal(json(await postJson(`${instance.url}api/annotations/remove`, { slide: 2, ids: ['a'] })).count, 0);
    assert.equal((await postJson(`${instance.url}api/annotations/remove`, { slide: 2, ids: 'a' })).status, 400);
    await postJson(`${instance.url}api/annotations/add`, { slide: 3, stroke: ring('c', 0.5, 0.5) });
    assert.equal(json(await postJson(`${instance.url}api/annotations/clear`, { slide: 3 })).removed, 1);
  } finally {
    await registry.closeAll();
  }
});

test('the report tells the agent what was marked, with the shapes under it and a numbered picture, and clears the unread flag', async () => {
  const { registry, instance, file } = await setup();
  try {
    await openAndWait(instance, file);
    assert.equal(instance.unreadAnnotations(), null);
    instance.addStroke(2, ring('a', 0.22, 0.14, 0.06));

    assert.deepEqual(instance.unreadAnnotations(), { deck: 'review.pptx', path: file, slides: [2], drafts: [] });
    const { info, images } = await instance.annotationReport();
    assert.equal(info.annotatedSlides, 1);
    assert.equal(info.shapeNamesAvailable, true);
    assert.equal(info.slides[0].slide, 2);
    assert.equal(info.slides[0].marks[0].kind, 'enclosure');
    assert.equal(info.slides[0].marks[0].onShapes[0].name, 'slide/title');
    assert.equal(images.length, 1);
    const picture = decodePng(images[0].data);
    assert.equal(picture.width, 16);

    assert.equal(instance.unreadAnnotations(), null, 'reading marks them as read');
    instance.addStroke(2, ring('b', 0.6, 0.7));
    assert.deepEqual(instance.unreadAnnotations().slides, [2], 'a new stroke makes the slide unread again');

    const without = await instance.annotationReport({ slide: 2, includeImages: false });
    assert.equal(without.images.length, 0);
    assert.equal(without.info.slides[0].marks.length, 2);
    await assert.rejects(() => instance.annotationReport({ slide: 7 }), (error) => error instanceof ViewerError && error.code === 'slide_out_of_range');
  } finally {
    await registry.closeAll();
  }
});

test('two viewers of the same deck share one red pen, and reading it in one marks it read in the other', async () => {
  const { registry, instance, edit, file, decks, dataRoot } = await setup({ watchIntervalMs: 30 });
  try {
    const other = await registry.ensure('panel-2', { baseDir: decks });
    await openAndWait(instance, file);
    await openAndWait(other, file);

    instance.addStroke(2, ring('from-one', 0.3, 0.3));
    other.addStroke(2, ring('from-two', 0.6, 0.6));
    for (const viewer of [instance, other]) {
      assert.deepEqual(viewer.snapshot().annotations.counts, { 2: 2 });
      assert.deepEqual(viewer.annotationsFor(2).strokes.map((stroke) => stroke.id).sort(), ['from-one', 'from-two']);
    }

    const events = [];
    const unsubscribe = other._book.subscribe(() => events.push(other.snapshot().annotations.unread));
    await instance.annotationReport({ includeImages: false });
    unsubscribe();
    assert.deepEqual(other.snapshot().annotations.unread, [], 'the other viewer sees the marks as read');
    assert.equal(events.length, 1);

    edit.changed = [3];
    await fs.writeFile(file, 'rebuilt, and longer than before');
    await until(() => [instance, other].every((viewer) => viewer.summary().phase === 'ready' && viewer.snapshot().notice));
    assert.equal(instance._book, other._book, 'both follow the same book after a rebuild');
    assert.deepEqual(other.snapshot().annotations.counts, { 2: 2 });
  } finally {
    await registry.closeAll();
  }

  const again = await setup({ root: dataRoot });
  try {
    await openAndWait(again.instance, again.file);
    assert.equal(again.instance.snapshot().annotations.counts[2], undefined, 'a different deck path has its own red pen');
  } finally {
    await again.registry.closeAll();
  }
});

test('draft pages: added after the page on screen, drawn on, removed by the user, and set aside rather than lost', async () => {
  const { registry, instance, file, dataRoot } = await setup();
  try {
    await openAndWait(instance, file);
    instance.setCurrent(2);

    const added = await postJson(`${instance.url}api/drafts/add`, {});
    assert.equal(added.status, 200);
    const { id, title, after } = json(added).draft;
    assert.deepEqual([title, after], ['New slide 1', 2]);
    const state = json(await request(`${instance.url}api/state`));
    assert.equal(state.draft, id, 'the new page is shown');
    assert.equal(state.current, 2);
    assert.deepEqual(state.annotations.drafts, [{ id, title, after: 2, count: 0, comments: 0, unread: true }]);

    const second = json(await postJson(`${instance.url}api/drafts/add`, {})).draft;
    assert.deepEqual([second.after, instance.snapshot().annotations.drafts.map((draft) => draft.id)], [2, [id, second.id]], 'added from a draft, it goes right after that draft');

    assert.equal((await postJson(`${instance.url}api/annotations/add`, { draft: id, stroke: ring('d1', 0.5, 0.5) })).status, 200);
    assert.equal(json(await request(`${instance.url}api/annotations?draft=${id}`)).strokes.length, 1);
    assert.equal(json(await request(`${instance.url}api/annotations?slide=2`)).strokes.length, 0, 'the slide itself is untouched');
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { draft: 'dnope00000000', stroke: ring('d2', 0.5, 0.5) })).status, 404);
    assert.equal((await postJson(`${instance.url}api/annotations/add`, { draft: 7, stroke: ring('d2', 0.5, 0.5) })).status, 400);
    assert.equal(json(await postJson(`${instance.url}api/annotations/remove`, { draft: id, ids: ['d1'] })).count, 0);
    await postJson(`${instance.url}api/annotations/add`, { draft: id, stroke: ring('d3', 0.4, 0.4) });

    assert.equal((await postJson(`${instance.url}api/current`, { slide: 3 })).status, 200);
    assert.equal(instance.snapshot().draft, null, 'going to a slide leaves the draft');
    assert.equal((await postJson(`${instance.url}api/current`, { draft: id })).status, 200);
    assert.equal(instance.snapshot().draft, id);
    assert.equal((await postJson(`${instance.url}api/current`, { draft: 'dnope00000000' })).status, 404);

    assert.equal((await postJson(`${instance.url}api/drafts/remove`, { id })).status, 200);
    assert.equal(instance.snapshot().draft, null, 'removing the page on screen goes back to a slide');
    assert.equal(instance.snapshot().current, 2);
    assert.deepEqual(instance.snapshot().annotations.drafts.map((draft) => draft.id), [second.id]);
    assert.equal((await postJson(`${instance.url}api/drafts/remove`, { id })).status, 404);
    await instance.annotationStore.flush();
    await until(async () => (await fs.readdir(`${dataRoot}\\annotations\\archive`).catch(() => [])).length === 1);
  } finally {
    await registry.closeAll();
  }
});

test('draft pages can be dragged to another place, and the viewer keeps showing the one that moved', async () => {
  const { registry, instance, file } = await setup();
  try {
    await openAndWait(instance, file);
    const one = instance.addDraft({ after: 1 }).draft;
    const two = instance.addDraft({ afterDraft: one.id }).draft;
    instance.setDraft(one.id);
    const order = () => instance.snapshot().annotations.drafts.map((draft) => `${draft.id === one.id ? 'one' : 'two'}@${draft.after}`).join(' ');
    assert.equal(order(), 'one@1 two@1');

    const response = await postJson(`${instance.url}api/drafts/move`, { id: one.id, afterDraft: two.id });
    assert.equal(response.status, 200);
    assert.equal(json(response).changed, true);
    assert.equal(order(), 'two@1 one@1');
    assert.equal(instance.snapshot().draft, one.id, 'the page that was dragged stays on screen');
    assert.equal(json(await postJson(`${instance.url}api/drafts/move`, { id: one.id, after: 3 })).after, 3);
    assert.equal(order(), 'two@1 one@3');
    assert.equal(json(await postJson(`${instance.url}api/drafts/move`, { id: one.id, after: 3 })).changed, false);

    assert.equal((await postJson(`${instance.url}api/drafts/move`, { id: 5, after: 1 })).status, 400);
    assert.equal((await postJson(`${instance.url}api/drafts/move`, { id: one.id, afterDraft: 5 })).status, 400);
    assert.equal((await postJson(`${instance.url}api/drafts/move`, { id: 'dnope00000000', after: 1 })).status, 404);
    assert.equal((await postJson(`${instance.url}api/drafts/move`, { id: one.id, after: 99 })).status, 400);
    assert.equal(order(), 'two@1 one@3');
  } finally {
    await registry.closeAll();
  }
});

test('the report lists draft pages, even empty ones, and resolving them takes them off the page', async () => {
  const { registry, instance, file, decks } = await setup();
  try {
    const other = await registry.ensure('panel-2', { baseDir: decks });
    await openAndWait(instance, file);
    await openAndWait(other, file);
    instance.setCurrent(3);
    const sketched = instance.addDraft().draft;
    instance.addStroke(sketched.id, ring('s', 0.3, 0.3, 0.1));
    const empty = instance.addDraft({ after: 0 }).draft;
    assert.deepEqual(instance.unreadAnnotations().drafts.sort(), ['New slide 1', 'New slide 2']);

    const { info, images } = await instance.annotationReport();
    assert.equal(info.annotatedSlides, 0);
    assert.deepEqual(info.draftPages.map((draft) => [draft.id, draft.insertAfterSlide, draft.strokeCount]), [[empty.id, 0, 0], [sketched.id, 3, 1]]);
    assert.equal(info.draftPages[1].afterSlideTitle, 'Title 3');
    assert.equal(info.draftPages[1].marks[0].kind, 'enclosure');
    assert.equal(images.length, 1);
    assert.equal(images[0].draft, sketched.id);
    const picture = decodePng(images[0].data);
    assert.equal(picture.width, 1600, 'a draft page is a white page of the deck\'s size');
    assert.equal(picture.height, 900);
    assert.deepEqual([...picture.data.slice(0, 4)], [255, 255, 255, 255]);
    assert.equal(instance.unreadAnnotations(), null);
    assert.deepEqual(other.snapshot().annotations.unread, []);

    const only = await instance.annotationReport({ draft: sketched.id, includeImages: false });
    assert.equal(only.info.draftPages.length, 1);
    await assert.rejects(() => instance.annotationReport({ draft: 'dnope00000000' }), (error) => error.code === 'draft_not_found');
    assert.deepEqual((await instance.annotationReport({ slide: 2 })).info.draftPages, [], 'asking for one slide does not bring the drafts');

    assert.equal(instance.outline().draftPages.length, 2);
    other.setDraft(sketched.id);
    const resolved = await instance.resolveDrafts([sketched.id, 'dnope00000000']);
    assert.deepEqual(resolved.resolved.map((draft) => draft.id), [sketched.id]);
    assert.deepEqual(resolved.missing, ['dnope00000000']);
    assert.equal(other.snapshot().draft, null, 'a viewer showing a resolved draft falls back to a slide');
    assert.deepEqual(instance.snapshot().annotations.drafts.map((draft) => draft.id), [empty.id]);
    await assert.rejects(() => instance.resolveDrafts('d1'), (error) => error.code === 'invalid_input');
  } finally {
    await registry.closeAll();
  }
});

test('a sketchbook is a deck with one white page: no PowerPoint file, no rendering, and it is found again by name', async () => {
  const calls = [];
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  const make = () => new ViewerRegistry({ root: dataRoot, watchIntervalMs: 30, renderer: fakeRenderer({ calls }) });
  let registry = make();
  let sketchPath;
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    // A name with a slash and non-ASCII characters: it is made safe for a file name and otherwise kept.
    await instance.newSketchbook('アイデア / 資料案');
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    const state = instance.snapshot();
    assert.equal(state.deck.kind, 'sketchbook');
    assert.equal(state.deck.name, 'アイデア - 資料案', 'the name is made safe for a file name');
    assert.equal(state.deck.slideCount, 1);
    assert.equal(state.deck.slides[0].title, 'Blank page');
    assert.equal(instance.title, 'アイデア - 資料案');
    assert.equal(calls.length, 0, 'PowerPoint is not involved');
    assert.ok(state.deck.path.startsWith(dataRoot));
    sketchPath = state.deck.path;

    const page = decodePng(await instance.slideImage(1));
    assert.deepEqual([page.width, page.height], [1600, 900]);
    assert.deepEqual([...page.data.slice(0, 4)], [255, 255, 255, 255]);

    instance.addStroke(1, ring('base', 0.5, 0.5));
    const draft = instance.addDraft().draft;
    instance.addStroke(draft.id, ring('page2', 0.5, 0.5));
    const report = await instance.annotationReport();
    assert.equal(report.info.kind, 'sketchbook');
    assert.equal(report.info.annotatedSlides, 1);
    assert.equal(report.info.draftPages.length, 1);
    assert.equal(report.images.length, 2);
    await assert.rejects(() => instance.refresh(), (error) => error.code === 'not_supported');
    await assert.rejects(() => instance._openExternal(), (error) => error.code === 'not_supported');

    assert.equal((await postJson(`${instance.url}api/sketchbook`, { name: 'アイデア / 資料案' })).status, 202);
    assert.equal(instance.snapshot().deck.path, sketchPath, 'the same name continues the same sketchbook');
    await new Promise((resolve) => setTimeout(resolve, 200));
    assert.equal(calls.length, 0, 'the file watcher leaves a sketchbook alone');
  } finally {
    await registry.closeAll();
  }

  registry = make();
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    await instance.openDeck(sketchPath);
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.deepEqual(instance.snapshot().annotations.counts, { 1: 1 });
    assert.equal(instance.snapshot().annotations.drafts.length, 1);
    assert.equal(instance.snapshot().annotations.drafts[0].count, 1);

    const recent = json(await request(`${instance.url}api/state`)).recent;
    assert.ok(recent.some((entry) => entry.path === sketchPath), 'it is listed among the recent files');

    await assert.rejects(() => instance.openDeck(`${sketchPath}.missing.sketch`), () => true);
    await assert.rejects(() => instance.openDeck(path.join(decks, 'outside.sketch')), (error) => error.code === 'unsupported_file');
    await instance.newSketchbook('');
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    assert.match(instance.snapshot().deck.name, /^Sketch \d{8}-\d{6}$/);
    for (const hostile of ['CON', '..', '...', '   ']) {
      await instance.newSketchbook(hostile);
      assert.equal((await instance.whenSettled(5000)).phase, 'ready');
      const { path: made } = instance.snapshot().deck;
      assert.equal(path.dirname(made), path.join(dataRoot, 'sketchbooks'), `"${hostile}" stays inside the sketchbooks folder`);
    }
  } finally {
    await registry.closeAll();
  }
});

test('red pen survives a restart of the viewer', async () => {
  const first = await setup();
  try {
    await openAndWait(first.instance, first.file);
    first.instance.addStroke(1, ring('a', 0.5, 0.5));
    first.instance.addStroke(3, ring('b', 0.5, 0.5));
  } finally {
    await first.registry.closeAll();
  }
  const second = await setup({ root: first.dataRoot });
  try {
    await openAndWait(second.instance, first.file);
    assert.deepEqual(second.instance.snapshot().annotations.counts, { 1: 1, 3: 1 });
    assert.deepEqual(second.instance.snapshot().annotations.unread, [1, 3]);
    assert.equal(second.instance.snapshot().notice, null, 'nothing was rebuilt, so there is nothing to tell');
  } finally {
    await second.registry.closeAll();
  }
});

test('an edited deck is picked up by itself: unchanged slides keep their red pen, changed slides give theirs back', async () => {
  const { registry, instance, edit, file, dataRoot } = await setup({ watchIntervalMs: 30 });
  try {
    await openAndWait(instance, file);
    instance.setCurrent(3);
    instance.addStroke(1, ring('one', 0.5, 0.5));
    instance.addStroke(2, ring('two', 0.5, 0.5));
    instance.addStroke(2, ring('two-b', 0.2, 0.2));
    instance.addStroke(3, ring('three', 0.5, 0.5));
    const before = instance.snapshot().deck.key;

    edit.changed = [2];
    await fs.writeFile(file, 'v2 with a different size');

    await until(() => instance.snapshot().deck.key !== before && instance.summary().phase === 'ready');
    const state = instance.snapshot();
    assert.equal(state.current, 3, 'the slide being looked at does not change');
    assert.deepEqual(state.annotations.counts, { 1: 1, 3: 1 });
    assert.match(state.notice.text, /changed on disk/);
    assert.match(state.notice.text, /were kept/);
    assert.match(state.notice.text, /set aside/);

    const archive = await fs.readdir(`${dataRoot}\\annotations\\archive`);
    assert.equal(archive.length, 1);
    const archived = JSON.parse(await fs.readFile(`${dataRoot}\\annotations\\archive\\${archive[0]}`, 'utf8'));
    assert.deepEqual(Object.keys(archived.data.slides), ['2']);
  } finally {
    await registry.closeAll();
  }
});

test('a deck that keeps changing is reloaded once it settles, and a failed reload is not retried in a loop', async () => {
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  let renders = 0;
  let broken = false;
  const working = fakeRenderer({ slides: 2 });
  const registry = new ViewerRegistry({
    root: dataRoot,
    watchIntervalMs: 30,
    renderer: (args) => {
      renders += 1;
      return broken ? Promise.reject(new ViewerError('render_failed', 'boom')) : working(args);
    },
  });
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    const file = await makeDeckFile(decks, 'busy.pptx', 'v1');
    await openAndWait(instance, file);
    assert.equal(renders, 1);

    broken = true;
    await fs.writeFile(file, 'v2 changed');
    await until(() => instance.summary().phase === 'error');
    const failed = renders;
    await new Promise((resolve) => setTimeout(resolve, 300));
    assert.equal(renders, failed, 'the same broken file is not rendered again and again');

    broken = false;
    await fs.writeFile(file, 'v3 fixed and longer');
    await until(() => renders > failed && instance.summary().phase === 'ready');
    assert.equal(instance.summary().phase, 'ready');
  } finally {
    await registry.closeAll();
  }
});
