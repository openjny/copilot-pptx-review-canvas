import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { test } from 'node:test';
import { AnnotationBook, AnnotationStore, validateStroke } from '../lib/annotations.mjs';
import { ViewerError } from '../lib/errors.mjs';
import { analyzeSlide, clusterMarks } from '../lib/marks.mjs';
import { makeTempDir } from './helpers.mjs';

const stroke = (id, points, extra = {}) => ({ id, tool: 'pen', color: '#E5231B', width: 0.004, points, ...extra });
const line = (id, x0, y0, x1, y1) => stroke(id, [[x0, y0], [(x0 + x1) / 2, (y0 + y1) / 2], [x1, y1]]);
const circle = (id, cx, cy, r) => stroke(id, Array.from({ length: 41 }, (_, i) => [cx + r * Math.cos((i / 40) * 2 * Math.PI), cy + r * Math.sin((i / 40) * 2 * Math.PI)]));

function deckOf(hashes, key = 'a'.repeat(24), file = 'C:\\x\\deck.pptx') {
  return { key, source: { path: file }, slideCount: hashes.length, slides: hashes.map((hash) => ({ hash })) };
}

test('validateStroke clamps and normalizes, and rejects anything malformed', () => {
  const ok = validateStroke(stroke('s_1-a', [[-0.2, 0.5], [1.4, 0.123456789]]));
  assert.deepEqual(ok.points, [[0, 0.5], [1, 0.12346]]);
  assert.equal(ok.color, '#e5231b');
  for (const bad of [null, [], stroke('', [[0, 0]]), stroke('a b', [[0, 0]]), stroke('a', []), stroke('a', [[0]]), stroke('a', [[0, NaN]]),
    stroke('a', [[0, 0]], { tool: 'laser' }), stroke('a', [[0, 0]], { color: 'red' }), stroke('a', [[0, 0]], { width: 0 }), stroke('a', [[0, 0]], { width: 0.5 })]) {
    assert.throws(() => validateStroke(bad), (error) => error instanceof ViewerError && error.code === 'invalid_input');
  }
});

test('AnnotationBook counts strokes, flags changes as unread and enforces its limits', () => {
  const book = new AnnotationBook({ path: 'p', deckKey: 'k', slideCount: 3, hashes: ['h1', 'h2', 'h3'] });
  assert.equal(book.add(2, stroke('a', [[0, 0]])), true);
  assert.equal(book.add(2, stroke('a', [[0, 0]])), false);
  book.add(2, stroke('b', [[0, 0]]));
  book.add(3, stroke('c', [[0, 0]]));
  assert.deepEqual(book.counts(), { 2: 2, 3: 1 });
  assert.deepEqual(book.unreadSlides(), [2, 3]);
  assert.equal(book.rev, 3);
  book.markRead([2]);
  assert.deepEqual(book.unreadSlides(), [3]);
  assert.equal(book.remove(2, ['a', 'zzz']), 1);
  assert.deepEqual(book.unreadSlides(), [2, 3]);
  assert.equal(book.clear(3), 1);
  assert.deepEqual(book.counts(), { 2: 1 });
  for (let i = 0; i < 399; i++) book.add(1, stroke(`x${i}`, [[0, 0]]));
  book.add(1, stroke('last', [[0, 0]]));
  assert.throws(() => book.add(1, stroke('over', [[0, 0]])), (error) => error.code === 'annotation_limit');
});

test('AnnotationStore keeps strokes where the picture did not change and archives the rest', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const first = await store.attach(deckOf(['h1', 'h2', 'h3', 'h4']));
  assert.equal(first.changed, false);
  first.book.add(1, stroke('a', [[0.1, 0.1]]));
  first.book.add(2, stroke('b', [[0.2, 0.2]]));
  first.book.add(4, stroke('c', [[0.3, 0.3]]));
  first.book.markRead([1]);
  await store.save(first.book);

  const same = await store.attach(deckOf(['h1', 'h2', 'h3', 'h4']));
  assert.deepEqual(same.book.counts(), { 1: 1, 2: 1, 4: 1 });
  assert.deepEqual(same.book.unreadSlides(), [2, 4]);

  // Slide 2 was fixed (new picture), slide 4 moved to position 3 after a slide was removed, slide 1 is untouched.
  const rebuilt = await store.attach(deckOf(['h1', 'h2-fixed', 'h4'], 'b'.repeat(24)));
  assert.deepEqual(rebuilt.book.counts(), { 1: 1, 3: 1 });
  assert.deepEqual({ kept: rebuilt.kept, moved: rebuilt.moved, archived: rebuilt.archived, changed: rebuilt.changed }, { kept: 2, moved: 1, archived: 1, changed: true });
  const archives = await fs.readdir(store.archiveDir);
  assert.equal(archives.length, 1);
  const archived = JSON.parse(await fs.readFile(`${store.archiveDir}/${archives[0]}`, 'utf8'));
  assert.deepEqual(Object.keys(archived.data.slides), ['2']);

  const again = await store.attach(deckOf(['h1', 'h2-fixed', 'h4'], 'b'.repeat(24)));
  assert.deepEqual(again.book.counts(), { 1: 1, 3: 1 });
  assert.equal(again.changed, false);
});

test('draft pages go right after the page they were added to, with their own strokes and stable titles', () => {
  const book = new AnnotationBook({ path: 'p', deckKey: 'k', slideCount: 3, hashes: ['h1', 'h2', 'h3'] });
  const first = book.addDraft({ after: 2 });
  const second = book.addDraft({ after: 2 });
  const third = book.addDraft({ afterDraft: second.id });
  const top = book.addDraft({ after: 0 });
  assert.deepEqual(book.drafts().map((draft) => [draft.title, draft.after]), [
    ['New slide 4', 0], ['New slide 2', 2], ['New slide 3', 2], ['New slide 1', 2],
  ]);
  assert.deepEqual(book.drafts().map((draft) => draft.id), [top.id, second.id, third.id, first.id]);
  assert.throws(() => book.addDraft({ after: 4 }), (error) => error.code === 'slide_out_of_range');
  assert.throws(() => book.addDraft({ afterDraft: 'dnope00000000' }), (error) => error.code === 'draft_not_found');

  assert.equal(book.add(second.id, stroke('a', [[0.1, 0.1]])), true);
  assert.equal(book.strokes(second.id).length, 1);
  assert.deepEqual(book.counts(), {}, 'draft strokes are not counted as slide strokes');
  assert.throws(() => book.add('dnope00000000', stroke('b', [[0, 0]])), (error) => error.code === 'draft_not_found');
  assert.equal(book.remove(second.id, ['a']), 1);
  book.add(second.id, stroke('c', [[0.2, 0.2]]));
  assert.equal(book.clear(second.id), 1);

  assert.equal(book.unreadDrafts().length, 4, 'a draft is new to the agent even before anything is drawn on it');
  book.markRead([second.id]);
  assert.deepEqual(book.unreadDrafts().map((draft) => draft.id), [top.id, third.id, first.id]);

  const removed = book.removeDraft(first.id);
  assert.equal(removed.id, first.id);
  assert.equal(book.removeDraft(first.id), null);
  assert.equal(book.drafts().length, 3);
  for (let i = 0; i < 40; i++) {
    try {
      book.addDraft({ after: 1 });
    } catch (error) {
      assert.equal(error.code, 'annotation_limit');
      return;
    }
  }
  assert.fail('the number of drafts is limited');
});

test('moveDraft puts a draft right after a slide or another draft, and does nothing when it is already there', () => {
  const book = new AnnotationBook({ path: 'p', deckKey: 'k', slideCount: 4, hashes: ['h1', 'h2', 'h3', 'h4'] });
  const a = book.addDraft({ after: 1 });
  const b = book.addDraft({ afterDraft: a.id });
  const c = book.addDraft({ afterDraft: b.id });
  const order = () => book.drafts().map((draft) => `${draft.title.slice(-1)}@${draft.after}`).join(' ');
  assert.equal(order(), '1@1 2@1 3@1');

  const events = [];
  book.subscribe(() => events.push(book.rev));
  assert.deepEqual(book.moveDraft(c.id, { after: 1 }), { changed: true, after: 1 }, 'right after slide 1 comes before the drafts there');
  assert.equal(order(), '3@1 1@1 2@1');
  assert.deepEqual(book.moveDraft(a.id, { afterDraft: b.id }), { changed: true, after: 1 });
  assert.equal(order(), '3@1 2@1 1@1');
  assert.deepEqual(book.moveDraft(a.id, { after: 4 }), { changed: true, after: 4 });
  assert.equal(order(), '3@1 2@1 1@4');
  assert.equal(book.draft(a.id).afterHash, 'h4', 'it follows the picture of the slide it now comes after');
  assert.deepEqual(book.moveDraft(c.id, { after: 0 }), { changed: true, after: 0 });
  assert.equal(order(), '3@0 2@1 1@4');

  const calls = events.length;
  assert.deepEqual(book.moveDraft(c.id, { after: 0 }), { changed: false, after: 0 }, 'already there');
  assert.deepEqual(book.moveDraft(b.id, { after: 1 }), { changed: false, after: 1 });
  assert.equal(events.length, calls, 'an unchanged move tells nobody');
  assert.equal(order(), '3@0 2@1 1@4');

  book.markRead([a.id, b.id, c.id]);
  book.moveDraft(a.id, { after: 2 });
  assert.deepEqual(book.unreadDrafts().map((draft) => draft.id), [a.id], 'a moved draft is new to the agent again');

  assert.throws(() => book.moveDraft(a.id, { afterDraft: a.id }), (error) => error.code === 'invalid_input');
  assert.throws(() => book.moveDraft(a.id, { afterDraft: 'dnope00000000' }), (error) => error.code === 'draft_not_found');
  assert.throws(() => book.moveDraft('dnope00000000', { after: 1 }), (error) => error.code === 'draft_not_found');
  assert.throws(() => book.moveDraft(a.id, { after: 9 }), (error) => error.code === 'slide_out_of_range');
  assert.throws(() => book.moveDraft(a.id, { after: 1.5 }), (error) => error.code === 'slide_out_of_range');
  assert.equal(order(), '3@0 2@1 1@2', 'a refused move changes nothing');
});

test('draft titles written by the Japanese-only versions are read as "New slide"', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deckOf(['h1', 'h2']));
  for (let i = 0; i < 4; i++) book.addDraft({ after: 1 });
  const record = book.toRecord();
  // "新規スライド案 N" is the oldest name, "新しいスライド N" the one that followed; a title that was not generated stays as it is.
  const legacy = ['新規スライド案 7', '新しいスライド 2', '新しいスライド', 'My own title'];
  record.drafts.forEach((draft, i) => {
    draft.title = legacy[i];
  });
  await fs.mkdir(store.dir, { recursive: true });
  await fs.writeFile(store.fileFor('C:\\x\\deck.pptx'), JSON.stringify({ format: 'pptx-review-canvas.annotations', version: 1, data: record }));
  const reopened = await new AnnotationStore(root).attach(deckOf(['h1', 'h2']));
  assert.deepEqual(reopened.book.drafts().map((draft) => draft.title), ['New slide 7', 'New slide 2', 'New slide', 'My own title']);
});

test('AnnotationStore keeps draft pages across a rebuild, following the slide each was placed after', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const first = await store.attach(deckOf(['h1', 'h2', 'h3']));
  const afterTwo = first.book.addDraft({ after: 2 });
  const atTop = first.book.addDraft({ after: 0 });
  first.book.add(afterTwo.id, stroke('s', [[0.5, 0.5]]));
  await store.save(first.book);

  const reopened = await store.attach(deckOf(['h1', 'h2', 'h3']));
  assert.deepEqual(reopened.book.drafts().map((draft) => [draft.id, draft.after, draft.count]), [[atTop.id, 0, 0], [afterTwo.id, 2, 1]]);
  assert.equal(reopened.book.draftSeq, 2);
  reopened.book.addDraft({ after: 3 });
  assert.equal(reopened.book.drafts().at(-1).title, 'New slide 3', 'titles keep counting after a reopen');

  // A slide was inserted before "h2": the draft that followed h2 now follows slide 3.
  const rebuilt = await store.attach(deckOf(['h1', 'new', 'h2', 'h3'], 'b'.repeat(24)));
  assert.deepEqual(rebuilt.book.drafts().map((draft) => [draft.id, draft.after]), [[atTop.id, 0], [afterTwo.id, 3], [rebuilt.book.drafts().at(-1).id, 4]]);
  assert.equal(rebuilt.book.strokes(afterTwo.id).length, 1);

  // The slide it followed was rewritten: the draft stays where it was, never beyond the last slide.
  const rewritten = await store.attach(deckOf(['h1', 'x'], 'c'.repeat(24)));
  assert.ok(rewritten.book.drafts().every((draft) => draft.after <= 2));
  assert.equal(rewritten.book.drafts().length, 3);
});

test('removed draft pages are set aside in the archive with their strokes', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deckOf(['h1', 'h2']));
  const draft = book.addDraft({ after: 1 });
  book.add(draft.id, stroke('keep-me', [[0.4, 0.4]]));
  const removed = book.removeDraft(draft.id);
  await store.archiveDrafts(book, [removed]);
  await store.save(book);
  const names = await fs.readdir(store.archiveDir);
  assert.equal(names.length, 1);
  const archived = JSON.parse(await fs.readFile(`${store.archiveDir}/${names[0]}`, 'utf8'));
  assert.equal(archived.data.drafts[0].strokes[0].id, 'keep-me');
  assert.equal((await store.attach(deckOf(['h1', 'h2']))).book.drafts().length, 0);
});

test('AnnotationStore sets a damaged record aside instead of overwriting it', async () => {
  const root = await makeTempDir();
  const logs = [];
  const store = new AnnotationStore(root, { log: (message) => logs.push(message) });
  const deck = deckOf(['h1']);
  await fs.mkdir(store.dir, { recursive: true });
  await fs.writeFile(store.fileFor(deck.source.path), '{ broken');
  const attached = await store.attach(deck);
  assert.equal(attached.book.total(), 0);
  assert.equal(logs.length, 1);
  assert.equal((await fs.readdir(store.dir)).filter((name) => name.includes('.corrupt-')).length, 1);
});

test('clusterMarks joins strokes that touch and numbers the marks in reading order', () => {
  const strokes = [line('low', 0.6, 0.8, 0.9, 0.8), circle('c', 0.3, 0.3, 0.08), line('arrow', 0.4, 0.3, 0.5, 0.3), line('far', 0.8, 0.1, 0.9, 0.1)];
  const marks = clusterMarks(strokes, 16 / 9);
  assert.equal(marks.length, 3);
  assert.deepEqual(marks.map((mark) => mark.strokes.map((s) => s.id)), [['far'], ['c', 'arrow'], ['low']]);
});

test('analyzeSlide describes each mark and finds the shapes under it', () => {
  const shapes = [
    { name: 'slide/column/1', kind: 'group', depth: 0, x: 40, y: 100, w: 880, h: 300, text: '', alt: '' },
    { name: 'slide/column/1/title', kind: 'text', depth: 1, x: 40, y: 100, w: 400, h: 30, text: 'Get ready', alt: '' },
    { name: 'slide/column/1/box', kind: 'shape', depth: 1, x: 40, y: 140, w: 400, h: 200, text: '', alt: '' },
    { name: 'slide/column/1/body', kind: 'text', depth: 1, x: 50, y: 150, w: 380, h: 60, text: 'Build the instructions from the request and the state.', alt: '' },
    { name: 'background', kind: 'picture', depth: 0, x: 0, y: 0, w: 960, h: 540, text: '', alt: 'bg' },
  ];
  const underline = stroke('u', [[0.06, 0.34], [0.2, 0.345], [0.43, 0.34]]);
  const ring = circle('r', 0.3, 0.33, 0.07);
  const { marks, badges } = analyzeSlide({ strokes: [ring, underline], aspect: 16 / 9, shapes, slideW: 960, slideH: 540 });
  assert.equal(marks.length, 1, 'the underline touches the circle, so they are one mark');
  assert.equal(marks[0].kind, 'multiple');
  assert.deepEqual(marks[0].parts.sort(), ['enclosure', 'underline']);
  assert.equal(marks[0].strokeCount, 2);
  assert.equal(marks[0].number, 1);
  const names = marks[0].onShapes.map((shape) => shape.name);
  assert.ok(names.includes('slide/column/1/body'), `candidates: ${names}`);
  assert.ok(!names.includes('background'), 'the full-slide picture is not a useful answer');
  assert.equal(marks[0].onShapes.find((shape) => shape.name === 'slide/column/1/body').group, 'slide/column/1');
  assert.equal(badges.length, 1);

  const alone = analyzeSlide({ strokes: [underline], aspect: 16 / 9, shapes, slideW: 960, slideH: 540 });
  assert.equal(alone.marks[0].kind, 'underline');
  assert.equal(alone.marks[0].region, 'middle left');
  assert.equal(analyzeSlide({ strokes: [ring], aspect: 16 / 9, shapes: null, slideW: 960, slideH: 540 }).marks[0].onShapes.length, 0);
  assert.equal(analyzeSlide({ strokes: [ring], aspect: 16 / 9, shapes: [], slideW: 960, slideH: 540 }).marks[0].kind, 'enclosure');
});

test('analyzeSlide falls back to the nearest shapes for a thin mark between two rows', () => {
  const shapes = [
    { name: 'head', kind: 'text', depth: 0, x: 180, y: 139, w: 339, h: 24, text: 'What to decide', alt: '' },
    { name: 'rule', kind: 'line', depth: 0, x: 46, y: 163, w: 868, h: 0, text: '', alt: '' },
    { name: 'row-1', kind: 'text', depth: 0, x: 180, y: 163, w: 339, h: 52, text: 'Agent implementation', alt: '' },
    { name: 'far', kind: 'text', depth: 0, x: 180, y: 400, w: 339, h: 30, text: 'Far below', alt: '' },
  ];
  const underline = stroke('u', [[0.1, 0.3], [0.25, 0.3], [0.4, 0.3]]);
  const { marks } = analyzeSlide({ strokes: [underline], aspect: 16 / 9, shapes, slideW: 960, slideH: 540 });
  const names = marks[0].onShapes.map((shape) => shape.name);
  assert.deepEqual(names.slice(0, 2).sort(), ['head', 'row-1'], `candidates: ${names}`);
  assert.ok(marks[0].onShapes.every((shape) => shape.near === true));
  assert.ok(!names.includes('far'));
});
