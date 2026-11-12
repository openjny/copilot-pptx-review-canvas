import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { AnnotationBook, AnnotationStore } from '../lib/annotations.mjs';
import { MAX_COMMENTS_PER_PAGE, MAX_REPLIES_PER_COMMENT, commentKind, newComment, newReply, validateAnchor, validateAuthor, validateComment, validateCommentText } from '../lib/comments.mjs';
import { analyzeComments } from '../lib/marks.mjs';
import { decodePng } from '../lib/png.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makePackageFile, makePng, makeTempDir, postJson, request } from './helpers.mjs';

const IDS = [256, 257, 258, 259];
const SHAPES = { 2: [{ name: 'slide/2/title', kind: 'text', depth: 0, x: 10, y: 10, w: 200, h: 60, text: 'Second' }, { name: 'slide/2/body', kind: 'text', depth: 0, x: 10, y: 300, w: 400, h: 100, text: 'Body' }] };
const pen = (id) => ({ id, tool: 'pen', color: '#e5231b', width: 0.004, points: [[0.5, 0.5], [0.7, 0.6]] });
const deck = (key, hashes) => ({ key, source: { path: 'C:\\x\\deck.pptx' }, slideCount: hashes.length, slides: hashes.map((hash) => ({ hash })) });
const rejects = (fn, code) => assert.throws(fn, (error) => error.code === code);

test('comment text and anchors are checked and put in a canonical form', () => {
  assert.equal(validateCommentText('  hello\r\nworld  '), 'hello\nworld');
  rejects(() => validateCommentText('   '), 'invalid_input');
  rejects(() => validateCommentText(5), 'invalid_input');
  rejects(() => validateCommentText('x'.repeat(2001)), 'invalid_input');

  assert.equal(validateAnchor(null), null);
  assert.equal(validateAnchor(undefined), null);
  assert.deepEqual(validateAnchor({ x: 0.2, y: 0.3 }), { x: 0.2, y: 0.3, w: 0, h: 0 }, 'a click is a point');
  assert.deepEqual(validateAnchor({ x: 0.2, y: 0.3, w: 0.001, h: 0.002 }), { x: 0.2, y: 0.3, w: 0, h: 0 }, 'a drag that is too small is a point too');
  assert.deepEqual(validateAnchor({ x: 0.2, y: 0.3, w: 0.4, h: 0.25 }), { x: 0.2, y: 0.3, w: 0.4, h: 0.25 });
  assert.deepEqual(validateAnchor({ x: -1, y: 2, w: 5, h: 5 }), { x: 0, y: 1, w: 1, h: 0 }, 'it is kept inside the page');
  assert.deepEqual(validateAnchor({ x: 0.9, y: 0.1, w: 0.5, h: 0.2 }), { x: 0.9, y: 0.1, w: 0.1, h: 0.2 });
  for (const bad of [{ x: 'a', y: 0 }, { x: 0 }, [], 'x', { x: NaN, y: 0 }]) rejects(() => validateAnchor(bad), 'invalid_input');

  const made = newComment({ text: ' a ', anchor: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 } });
  assert.match(made.id, /^c[0-9a-f]{12}$/);
  assert.equal(commentKind(made), 'area');
  assert.equal(commentKind({ anchor: { x: 0, y: 0, w: 0, h: 0 } }), 'point');
  assert.equal(commentKind({ anchor: null }), 'page');
  assert.deepEqual(validateComment(made), made);
  rejects(() => validateComment({ ...made, id: 'x' }), 'invalid_input');
  rejects(() => validateComment({ ...made, createdAt: 'now' }), 'invalid_input');
});

test('a book holds comments on slides and draft pages, and counts them', () => {
  const book = new AnnotationBook({ path: 'C:\\x\\deck.pptx', deckKey: 'k'.repeat(24), slideCount: 4, hashes: ['h1', 'h2', 'h3', 'h4'] });
  assert.deepEqual(book.annotatedSlides(), []);
  const whole = book.addComment(3, { text: 'whole page' });
  const area = book.addComment(3, { text: 'this area', anchor: { x: 0.1, y: 0.1, w: 0.3, h: 0.2 } });
  const draft = book.addDraft({ after: 2 });
  book.addComment(draft.id, { text: 'idea' });

  assert.deepEqual(book.comments(3).map((comment) => comment.text), ['whole page', 'this area']);
  assert.deepEqual(book.commentCounts(), { 3: 2 });
  assert.deepEqual(book.annotatedSlides(), [3], 'a slide with only comments is annotated');
  assert.deepEqual(book.unreadSlides(), [3]);
  assert.equal(book.drafts()[0].comments, 1);

  const before = book.rev;
  assert.equal(book.updateComment(3, whole.id, 'whole page'), null, 'the same text changes nothing');
  assert.equal(book.rev, before);
  const changed = book.updateComment(3, whole.id, 'changed');
  assert.equal(changed.text, 'changed');
  assert.ok(changed.updatedAt >= whole.updatedAt);
  rejects(() => book.updateComment(3, 'cnone', 'x'), 'comment_not_found');
  rejects(() => book.addComment('d000000000000', { text: 'x' }), 'draft_not_found');

  book.markRead([3]);
  const removed = book.removeComment(3, area.id);
  assert.equal(removed.index, 1);
  assert.deepEqual(book.unreadSlides(), [3], 'removing a comment is a change the agent has not seen');
  assert.equal(book.removeComment(3, area.id), null);
  book.reinsertComment(3, removed.comment, removed.index);
  assert.deepEqual(book.comments(3).map((comment) => comment.id), [whole.id, area.id]);
  rejects(() => book.reinsertComment(3, removed.comment, 0), 'invalid_input');

  for (let i = book.comments(1).length; i < MAX_COMMENTS_PER_PAGE; i++) book.addComment(1, { text: `c${i}` });
  rejects(() => book.addComment(1, { text: 'one too many' }), 'annotation_limit');
});

test('comments are saved, survive a restart, and follow their slide through moves, deletions and undo', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2', 'h3', 'h4']));
  const kept = book.addComment(4, { text: 'on four', anchor: { x: 0.5, y: 0.5 } });
  book.addComment(2, { text: 'on two' });
  book.addComment(3, { text: 'on three' });
  book.add(3, pen('s3'));
  const draft = book.addDraft({ after: 3 });
  const onDraft = book.addComment(draft.id, { text: 'draft note' });
  await store.save(book);

  const reopened = await new AnnotationStore(root).attach(deck('a'.repeat(24), ['h1', 'h2', 'h3', 'h4']));
  assert.deepEqual(reopened.book.commentCounts(), { 2: 1, 3: 1, 4: 1 });
  assert.deepEqual(reopened.book.comments(4), [kept], 'the same comment comes back');
  assert.equal(reopened.book.comments(draft.id)[0].text, 'draft note');

  const saved = { strokes: [...book.strokes(3)], comments: [...book.comments(3)], unread: book.isUnread(3), draftIds: [draft.id] };
  const deleted = await store.attach(deck('b'.repeat(24), ['h1', 'h2', 'h4']), { order: [1, 2, 4] });
  assert.equal(deleted.archived, 2, 'the deleted slide had a stroke and a comment');
  assert.deepEqual(book.commentCounts(), { 2: 1, 3: 1 });
  assert.equal(book.comments(3)[0].text, 'on four');
  const [archiveFile] = await fs.readdir(path.join(root, 'annotations', 'archive'));
  const archive = JSON.parse(await fs.readFile(path.join(root, 'annotations', 'archive', archiveFile), 'utf8'));
  assert.equal(archive.data.slides['3'].comments[0].text, 'on three', 'the comment of the deleted slide is set aside');

  await store.attach(deck('c'.repeat(24), ['h1', 'h2', 'h3', 'h4']), { insert: { slide: 3, ...saved } });
  assert.deepEqual(book.commentCounts(), { 2: 1, 3: 1, 4: 1 });
  assert.equal(book.comments(3)[0].text, 'on three');
  assert.equal(book.comments(4)[0].text, 'on four');
  assert.equal(book.comments(draft.id)[0].id, onDraft.id);

  const moved = await store.attach(deck('d'.repeat(24), ['h1', 'h3', 'h2', 'h4']), { order: [1, 3, 2, 4] });
  assert.equal(moved.archived, 0);
  assert.equal(book.comments(2)[0].text, 'on three');
  assert.equal(book.comments(3)[0].text, 'on two');
});

test('after a rebuild, the comments stay on slides whose picture did not change and are set aside from the others', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2', 'h3']));
  book.addComment(1, { text: 'one' });
  book.addComment(2, { text: 'two' });
  book.addComment(3, { text: 'three' });
  await store.save(book);

  const fresh = new AnnotationStore(root);
  const rebuilt = await fresh.attach(deck('b'.repeat(24), ['h1', 'changed', 'h3']));
  assert.deepEqual(rebuilt.book.commentCounts(), { 1: 1, 3: 1 });
  assert.equal(rebuilt.kept, 2);
  assert.equal(rebuilt.archived, 1);
  assert.equal((await fs.readdir(path.join(root, 'annotations', 'archive'))).length, 1);
});

test('a record written before comments existed still loads', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2']));
  book.add(1, pen('s1'));
  await store.save(book);
  const file = store.fileFor('C:\\x\\deck.pptx');
  const record = JSON.parse(await fs.readFile(file, 'utf8'));
  for (const entry of Object.values(record.data.slides)) delete entry.comments;
  await fs.writeFile(file, JSON.stringify(record));
  const reopened = await new AnnotationStore(root).attach(deck('a'.repeat(24), ['h1', 'h2']));
  assert.deepEqual(reopened.book.counts(), { 1: 1 });
  assert.deepEqual(reopened.book.commentCounts(), {});

  record.data.slides['1'].comments = [{ id: 'bad' }];
  await fs.writeFile(file, JSON.stringify(record));
  const broken = await new AnnotationStore(root).attach(deck('a'.repeat(24), ['h1', 'h2']));
  assert.deepEqual(broken.book.counts(), {}, 'a record with a broken comment is set aside, not half loaded');
});

test('analyzeComments says where each comment is and what lies under it, and lists what to draw', () => {
  const comments = [
    newComment({ text: 'whole page' }),
    newComment({ text: 'this title', anchor: { x: 0.0, y: 0.0, w: 0.25, h: 0.15 } }),
    newComment({ text: 'here', anchor: { x: 0.2, y: 0.6 } }),
  ];
  const result = analyzeComments({ comments, aspect: 960 / 540, shapes: SHAPES[2], slideW: 960, slideH: 540 });
  assert.deepEqual(result.comments.map((entry) => [entry.number, entry.kind, entry.text]), [[1, 'page', 'whole page'], [2, 'area', 'this title'], [3, 'point', 'here']]);
  assert.equal(result.comments[0].boxPercent, undefined, 'a page comment has no place');
  assert.deepEqual(result.comments[1].boxPercent, { x: 0, y: 0, w: 25, h: 15 });
  assert.equal(result.comments[1].onShapes[0].name, 'slide/2/title');
  assert.deepEqual(result.comments[2].boxPercent, { x: 20, y: 60 });
  assert.deepEqual(result.rects, [{ x: 0, y: 0, w: 0.25, h: 0.15 }], 'only an area gets a rectangle');
  assert.deepEqual(result.badges.map((badge) => badge.number), [2, 3], 'and a point or an area gets a number');
});

async function setup() {
  const dataRoot = await makeTempDir();
  const decks = await makeTempDir();
  const calls = [];
  const registry = new ViewerRegistry({ root: dataRoot, watchIntervalMs: 0, renderer: fakeRenderer({ slides: 4, calls, shapes: SHAPES }) });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makePackageFile(decks, 'deck.pptx', IDS);
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  return { registry, instance, file, dataRoot };
}

const post = (instance, route, body) => postJson(`${instance.url}api/${route}`, body);

test('comments are written, read, changed and deleted through the page API', async () => {
  const { registry, instance } = await setup();
  try {
    const empty = await request(`${instance.url}api/comments?slide=2`);
    assert.deepEqual(json(empty).comments, []);

    const added = await post(instance, 'comments/add', { slide: 2, text: '  fix the title  ', anchor: { x: 0.05, y: 0.05, w: 0.3, h: 0.1 } });
    assert.equal(added.status, 200);
    const { comment, version } = json(added);
    assert.equal(comment.text, 'fix the title');
    assert.equal(comment.anchor.w, 0.3);
    assert.ok(version.endsWith(`:${instance.snapshot().annotations.version.split(':')[1]}`));
    const page = json(await post(instance, 'comments/add', { slide: 2, text: 'overall: too dense' })).comment;
    assert.equal(page.anchor, null);

    assert.deepEqual(json(await request(`${instance.url}api/comments?slide=2`)).comments.map((entry) => entry.text), ['fix the title', 'overall: too dense']);
    const state = instance.snapshot();
    assert.deepEqual(state.annotations.commentCounts, { 2: 2 });
    assert.deepEqual(state.annotations.unread, [2]);

    const updated = await post(instance, 'comments/update', { slide: 2, id: comment.id, text: 'fix the title and the date' });
    assert.equal(json(updated).comment.text, 'fix the title and the date');
    assert.equal((await post(instance, 'comments/update', { slide: 2, id: 'cnone', text: 'x' })).status, 404);
    assert.equal((await post(instance, 'comments/update', { slide: 2, id: comment.id, text: '  ' })).status, 400);
    assert.equal((await post(instance, 'comments/add', { slide: 2, text: 'x', anchor: 'nowhere' })).status, 400);
    assert.equal((await post(instance, 'comments/add', { slide: 9, text: 'x' })).status, 400);

    const draft = instance.addDraft({ after: 2 }).draft.id;
    assert.equal((await post(instance, 'comments/add', { draft, text: 'on the new slide' })).status, 200);
    assert.deepEqual(json(await request(`${instance.url}api/comments?draft=${draft}`)).comments.map((entry) => entry.text), ['on the new slide']);
    assert.equal(instance.snapshot().annotations.drafts[0].comments, 1);

    assert.equal((await post(instance, 'comments/remove', { slide: 2, id: page.id })).status, 200);
    assert.equal((await post(instance, 'comments/remove', { slide: 2, id: page.id })).status, 404);
    assert.deepEqual(instance.outline().slides[1], { slide: 2, title: 'Title 2', comments: 1 });
    assert.equal(instance.outline().draftPages[0].comments, 1);
  } finally {
    await registry.closeAll();
  }
});

test('the agent reads the comments: text, place, shapes under them, and an image with blue numbers', async () => {
  const { registry, instance } = await setup();
  try {
    instance.addComment(2, { text: 'rewrite this title', anchor: { x: 0, y: 0, w: 0.25, h: 0.15 } });
    instance.addComment(2, { text: 'overall: too dense' });
    instance.addComment(3, { text: 'only about the whole page' });
    instance.addStroke(1, pen('s1'));
    const draft = instance.addDraft({ after: 3 }).draft.id;
    instance.addComment(draft, { text: 'a summary slide', anchor: { x: 0.1, y: 0.1 } });

    const { info, images } = await instance.annotationReport();
    assert.equal(info.annotatedSlides, 3);
    const two = info.slides.find((entry) => entry.slide === 2);
    assert.equal(two.commentCount, 2);
    assert.deepEqual(two.comments.map((entry) => [entry.number, entry.kind, entry.text]), [[1, 'area', 'rewrite this title'], [2, 'page', 'overall: too dense']]);
    assert.equal(two.comments[0].onShapes[0].name, 'slide/2/title');
    assert.equal(info.slides.find((entry) => entry.slide === 3).comments[0].kind, 'page');
    assert.equal(info.draftPages[0].comments[0].kind, 'point');

    const imagesOf = (entry) => images.filter((image) => (entry.slide ? image.slide === entry.slide : image.draft === entry.draft));
    assert.equal(imagesOf({ slide: 2 }).length, 1, 'a page with an area comment comes as an image');
    assert.equal(imagesOf({ slide: 3 }).length, 0, 'a page with only a whole-page comment needs none');
    assert.equal(imagesOf({ draft }).length, 1);
    const drawn = decodePng(imagesOf({ slide: 2 })[0].data);
    let blue = 0;
    for (let i = 0; i < drawn.data.length; i += 4) if (drawn.data[i] < 80 && drawn.data[i + 2] > 200) blue += 1;
    assert.ok(blue > 0, 'the comment is drawn in blue');
    assert.notDeepEqual(imagesOf({ slide: 2 })[0].data, makePng(2));

    assert.deepEqual(instance.snapshot().annotations.unread, []);
    assert.equal(instance.unreadAnnotations(), null, 'reading marks them read');
    instance.addComment(3, { text: 'one more' });
    assert.deepEqual(instance.unreadAnnotations().slides, [3]);
    assert.equal((await instance.annotationReport({ slide: 3 })).info.slides[0].commentCount, 2);
  } finally {
    await registry.closeAll();
  }
});

test('deleting a comment can be undone and redone, and a deleted slide brings its comments back', async () => {
  const { registry, instance, file } = await setup();
  try {
    const original = await fs.readFile(file);
    const first = instance.addComment(2, { text: 'first' }).comment;
    instance.addComment(2, { text: 'second' });
    await post(instance, 'comments/remove', { slide: 2, id: first.id });
    assert.equal(instance.snapshot().history.undo.label, 'Delete comment');
    assert.deepEqual(instance.commentsFor(2).comments.map((entry) => entry.text), ['second']);

    const undo = await post(instance, 'history/undo', { id: instance.snapshot().history.undo.id });
    assert.equal(undo.status, 200);
    assert.deepEqual(instance.commentsFor(2).comments.map((entry) => entry.text), ['first', 'second'], 'it is back in the same place');
    assert.equal((await post(instance, 'history/redo', { id: instance.snapshot().history.redo.id })).status, 200);
    assert.deepEqual(instance.commentsFor(2).comments.map((entry) => entry.text), ['second']);
    await post(instance, 'history/undo', { id: instance.snapshot().history.undo.id });

    assert.equal((await post(instance, 'slides/delete', { slide: 2 })).status, 200);
    assert.deepEqual(instance.snapshot().annotations.commentCounts, {});
    assert.equal((await post(instance, 'history/undo', { id: instance.snapshot().history.undo.id })).status, 200);
    assert.deepEqual(await fs.readFile(file), original);
    assert.deepEqual(instance.commentsFor(2).comments.map((entry) => entry.text), ['first', 'second']);
    assert.deepEqual(instance.snapshot().annotations.commentCounts, { 2: 2 });
  } finally {
    await registry.closeAll();
  }
});

test('a comment from an older record reads as the user\'s, with no replies and not resolved; bad authors and replies are refused', () => {
  const old = { id: 'c000000000001', text: 'old', anchor: null, createdAt: 1, updatedAt: 1 };
  assert.deepEqual(validateComment(old), { ...old, author: 'user', resolved: false, replies: [] });
  assert.equal(validateComment({ ...old, author: 'agent', resolved: true, resolvedAt: 5 }).resolvedAt, 5);
  assert.equal(validateAuthor(undefined), 'user');
  rejects(() => validateAuthor('robot'), 'invalid_input');
  rejects(() => validateComment({ ...old, author: 'robot' }), 'invalid_input');
  rejects(() => validateComment({ ...old, resolved: 'yes' }), 'invalid_input');
  rejects(() => validateComment({ ...old, replies: 'none' }), 'invalid_input');
  rejects(() => validateComment({ ...old, replies: [{ id: 'x', text: 'a', createdAt: 1 }] }), 'invalid_input');
  rejects(() => validateComment({ ...old, replies: Array.from({ length: MAX_REPLIES_PER_COMMENT + 1 }, () => ({})) }), 'invalid_input');
  const reply = newReply({ text: ' yes ', author: 'agent' });
  assert.match(reply.id, /^r[0-9a-f]{12}$/);
  assert.deepEqual(validateComment({ ...old, replies: [reply] }).replies, [reply]);
});

test('threads: who wrote what decides what the agent has not read, and a resolved thread no longer counts', () => {
  const book = new AnnotationBook({ path: 'C:\\x\\deck.pptx', deckKey: 'k'.repeat(24), slideCount: 4, hashes: ['h1', 'h2', 'h3', 'h4'] });
  const fromAgent = book.addComment(2, { text: 'I changed the title. Does it match?', anchor: { x: 0.1, y: 0.1 }, author: 'agent' });
  assert.equal(fromAgent.author, 'agent');
  assert.deepEqual(book.unreadSlides(), [], 'what the agent writes is not unread for the agent');
  assert.deepEqual(book.commentCounts(), { 2: 1 });
  rejects(() => book.updateComment(2, fromAgent.id, 'edited by the user'), 'comment_not_editable');

  const rev = book.rev;
  const { comment, reply } = book.replyComment(2, fromAgent.id, { text: 'Yes, but shorten it.', author: 'user' });
  assert.ok(book.rev > rev);
  assert.equal(reply.author, 'user');
  assert.deepEqual(comment.replies.map((entry) => entry.text), ['Yes, but shorten it.']);
  assert.deepEqual(book.unreadSlides(), [2], 'a reply from the user is something to read');

  book.markRead([2]);
  book.replyComment(2, fromAgent.id, { text: 'Shortened.', author: 'agent' });
  assert.deepEqual(book.unreadSlides(), [], 'a reply from the agent is not');
  assert.deepEqual(book.comments(2)[0].replies.map((entry) => entry.author), ['user', 'agent']);

  const resolved = book.setResolved(2, fromAgent.id, true);
  assert.equal(resolved.resolved, true);
  assert.ok(resolved.resolvedAt > 0);
  assert.deepEqual(book.unreadSlides(), [], 'resolving is not feedback to read');
  assert.deepEqual(book.commentCounts(), {}, 'a resolved thread is not counted');
  assert.deepEqual(book.annotatedSlides(), [], 'and a slide with only resolved threads has nothing to say');
  assert.equal(book.setResolved(2, fromAgent.id, true), null, 'resolving twice changes nothing');
  assert.equal(book.setResolved(2, fromAgent.id, false).resolvedAt, undefined);
  assert.deepEqual(book.commentCounts(), { 2: 1 }, 'and it can be opened again');

  assert.equal(book.findComment(fromAgent.id).target, 2);
  const draft = book.addDraft({ after: 1 });
  const onDraft = book.addComment(draft.id, { text: 'on the draft', author: 'agent' });
  assert.equal(book.findComment(onDraft.id).target, draft.id);
  assert.equal(book.findComment('c999999999999'), null);
  rejects(() => book.replyComment(2, 'cnone', { text: 'x' }), 'comment_not_found');
  rejects(() => book.setResolved(2, 'cnone', true), 'comment_not_found');

  const busy = book.addComment(3, { text: 'busy' });
  for (let i = 0; i < MAX_REPLIES_PER_COMMENT; i++) book.replyComment(3, busy.id, { text: `r${i}` });
  rejects(() => book.replyComment(3, busy.id, { text: 'one too many' }), 'annotation_limit');
});

test('replies and resolved threads are saved and come back', async () => {
  const root = await makeTempDir();
  const store = new AnnotationStore(root);
  const { book } = await store.attach(deck('a'.repeat(24), ['h1', 'h2']));
  const thread = book.addComment(1, { text: 'from the agent', author: 'agent' });
  book.replyComment(1, thread.id, { text: 'from the user', author: 'user' });
  book.setResolved(1, thread.id, true);
  await store.save(book);
  const [back] = (await new AnnotationStore(root).attach(deck('a'.repeat(24), ['h1', 'h2']))).book.comments(1);
  assert.equal(back.author, 'agent');
  assert.equal(back.resolved, true);
  assert.deepEqual(back.replies.map((entry) => [entry.author, entry.text]), [['user', 'from the user']]);

  const file = store.fileFor('C:\\x\\deck.pptx');
  const record = JSON.parse(await fs.readFile(file, 'utf8'));
  delete record.data.slides['1'].comments[0].author;
  delete record.data.slides['1'].comments[0].replies;
  delete record.data.slides['1'].comments[0].resolved;
  await fs.writeFile(file, JSON.stringify(record));
  const [legacy] = (await new AnnotationStore(root).attach(deck('a'.repeat(24), ['h1', 'h2']))).book.comments(1);
  assert.deepEqual([legacy.author, legacy.resolved, legacy.replies], ['user', false, []], 'a comment saved before threads existed is still the user\'s');
});

test('the page cannot write as the agent, and it replies and resolves through the API', async () => {
  const { registry, instance } = await setup();
  try {
    const spoof = json(await post(instance, 'comments/add', { slide: 2, text: 'I claim to be the agent', author: 'agent' })).comment;
    assert.equal(spoof.author, 'user');
    const root = (await instance.addAgentComment({ slide: 2, text: 'Check this number', anchor: { x: 50, y: 50 } }));
    const replied = await post(instance, 'comments/reply', { slide: 2, id: root.id, text: '  it is right  ' });
    assert.equal(replied.status, 200);
    assert.equal(json(replied).reply.text, 'it is right');
    assert.equal(json(replied).reply.author, 'user');
    assert.equal(json(replied).comment.replies.length, 1);
    assert.equal((await post(instance, 'comments/reply', { slide: 2, id: root.id, text: ' ' })).status, 400);
    assert.equal((await post(instance, 'comments/reply', { slide: 2, id: 'cnone', text: 'x' })).status, 404);

    assert.equal((await post(instance, 'comments/update', { slide: 2, id: root.id, text: 'rewrite the agent' })).status, 409, 'the agent\'s words cannot be edited');
    const resolved = await post(instance, 'comments/resolve', { slide: 2, id: root.id, resolved: true });
    assert.equal(json(resolved).comment.resolved, true);
    assert.equal((await post(instance, 'comments/resolve', { slide: 2, id: root.id, resolved: 'yes' })).status, 400);
    assert.deepEqual(instance.snapshot().annotations.commentCounts, { 2: 1 }, 'only the user\'s own comment is still open');
    assert.deepEqual(json(await request(`${instance.url}api/comments?slide=2`)).comments.map((entry) => [entry.author, entry.resolved, entry.replies.length]), [['user', false, 0], ['agent', true, 1]]);
    assert.equal(json(await post(instance, 'comments/resolve', { slide: 2, id: root.id, resolved: false })).comment.resolved, false);
  } finally {
    await registry.closeAll();
  }
});

test('the agent starts threads and replies; they come back marked as the agent\'s and are not unread', async () => {
  const { registry, instance } = await setup();
  try {
    const slideBox = instance.snapshot().deck;
    const area = await instance.addAgentComment({ slide: 2, text: 'I rewrote the title here.', shape: 'slide/2/title' });
    assert.equal(area.kind, 'area');
    const [stored] = instance.commentsFor(2).comments;
    assert.equal(stored.author, 'agent');
    assert.ok(Math.abs(stored.anchor.x - 10 / slideBox.slideWidth) < 0.001 && Math.abs(stored.anchor.w - 200 / slideBox.slideWidth) < 0.001, 'the shape\'s box, as fractions of the slide');
    const point = await instance.addAgentComment({ slide: 2, text: 'Weak evidence for this claim. Please review closely.', anchor: { x: 25, y: 75 } });
    assert.equal(point.kind, 'point');
    assert.deepEqual(instance.commentsFor(2).comments[1].anchor, { x: 0.25, y: 0.75, w: 0, h: 0 });
    const whole = await instance.addAgentComment({ slide: 3, text: 'Overall, this slide is dense.' });
    assert.equal(whole.kind, 'page');
    const draft = instance.addDraft({ after: 2 }).draft.id;
    await instance.annotationReport({ includeImages: false });
    assert.equal((await instance.addAgentComment({ draft, text: 'Is this the layout you meant?', anchor: { x: 10, y: 10, w: 30, h: 20 } })).draft, draft);
    assert.equal(instance.unreadAnnotations(), null, 'the agent\'s own comments are not unread');
    assert.deepEqual(instance.snapshot().annotations.commentCounts, { 2: 2, 3: 1 });

    for (const bad of [{ text: 'x' }, { slide: 2, draft, text: 'x' }, { slide: 2, text: 'x', anchor: { x: 1, y: 1 }, shape: 'slide/2/body' }, { slide: 2, text: 'x', shape: 'nope' }, { draft, text: 'x', shape: 'slide/2/body' }, { slide: 2, text: 'x', anchor: { x: 'a', y: 1 } }, { slide: 2, text: ' ' }, { slide: 9, text: 'x' }]) {
      await assert.rejects(() => instance.addAgentComment(bad), (error) => error.code === 'invalid_input' || error.code === 'slide_out_of_range', JSON.stringify(bad));
    }

    const answer = instance.replyAsAgent({ id: stored.id, text: 'Done.' });
    assert.equal(answer.slide, 2);
    assert.equal(answer.replies, 1);
    assert.throws(() => instance.replyAsAgent({ id: 'c999999999999', text: 'x' }), (error) => error.code === 'comment_not_found');
    assert.throws(() => instance.replyAsAgent({ id: 5, text: 'x' }), (error) => error.code === 'invalid_input');
    assert.equal(instance.unreadAnnotations(), null);

    // A reply from the user is read by the agent: it shows in the report with its author, and reading clears the unread state.
    instance.replyComment(2, stored.id, 'Looks good.');
    assert.deepEqual(instance.unreadAnnotations().slides, [2]);
    const { info } = await instance.annotationReport({ slide: 2, includeImages: false });
    const thread = info.slides[0].comments.find((entry) => entry.id === stored.id);
    assert.equal(thread.author, 'agent');
    assert.deepEqual(thread.replies, [{ author: 'agent', text: 'Done.' }, { author: 'user', text: 'Looks good.' }]);
    assert.equal(instance.unreadAnnotations(), null);

    instance.resolveComment(2, stored.id, true);
    const after = (await instance.annotationReport({ slide: 2, includeImages: false })).info.slides[0];
    assert.deepEqual(after.comments.map((entry) => entry.id), [instance.commentsFor(2).comments[1].id], 'a resolved thread is left out of the report');
    assert.equal(after.resolvedCommentCount, 1);
    assert.equal(after.comments[0].number, 1, 'and the numbers count only the open threads');
  } finally {
    await registry.closeAll();
  }
});

test('the theme is chosen on the page, kept for every viewer, and survives a restart', async () => {
  const { registry, instance, dataRoot } = await setup();
  try {
    assert.equal(instance.snapshot().theme, 'auto');
    assert.equal((await post(instance, 'theme', { theme: 'neon' })).status, 400);
    assert.equal((await post(instance, 'theme', {})).status, 400);
    const other = await registry.ensure('panel-2', { baseDir: dataRoot });
    const rev = other.snapshot().rev;
    const chosen = await post(instance, 'theme', { theme: 'dark' });
    assert.equal(json(chosen).theme, 'dark');
    assert.equal(instance.snapshot().theme, 'dark');
    assert.equal(other.snapshot().theme, 'dark', 'another viewer follows');
    assert.ok(other.snapshot().rev > rev, 'and is told');
  } finally {
    await registry.closeAll();
  }
  const again = new ViewerRegistry({ root: dataRoot, watchIntervalMs: 0, renderer: fakeRenderer({ slides: 4, calls: [], shapes: SHAPES }) });
  await again.ready();
  assert.equal(again.preferences.theme, 'dark');
  await again.closeAll();
});
