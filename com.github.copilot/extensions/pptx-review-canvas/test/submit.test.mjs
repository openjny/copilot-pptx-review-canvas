import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makeDeckFile, makeTempDir, postJson, request } from './helpers.mjs';

const ring = (id, cx, cy, r = 0.05) => ({
  id,
  tool: 'pen',
  color: '#e5231b',
  width: 0.004,
  points: Array.from({ length: 25 }, (_, i) => [cx + r * Math.cos((i / 24) * 2 * Math.PI), cy + r * Math.sin((i / 24) * 2 * Math.PI) * 1.6]),
});

/** A registry whose "agent" is a list: every message the viewer sends to it is kept in `sent`. */
async function setup({ sendToAgent, root, language } = {}) {
  const dataRoot = root ?? (await makeTempDir());
  const decks = await makeTempDir();
  const sent = [];
  const registry = new ViewerRegistry({
    root: dataRoot,
    renderer: fakeRenderer({ slides: 3 }),
    watchIntervalMs: 0,
    env: language ? { PPTX_CANVAS_LANG: language } : {},
    ...(sendToAgent === null ? {} : { sendToAgent: sendToAgent ?? (async (message) => { sent.push(message); }) }),
  });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makeDeckFile(decks, 'review.pptx', 'v1');
  await instance.openDeck(file);
  assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  const state = async () => json(await request(`${instance.url}api/state`));
  const submit = (body = {}) => postJson(`${instance.url}api/review/submit`, body);
  return { registry, instance, sent, state, submit, dataRoot, file };
}

test('the review can be submitted once for what the agent has not read, and again after the marks change', async () => {
  const { registry, instance, sent, state, submit } = await setup();
  try {
    assert.deepEqual((await state()).review, { canSubmit: false, submitted: false, slides: [], drafts: 0 });
    assert.equal((await submit()).status, 409, 'with nothing marked there is nothing to submit');
    assert.equal(json(await submit()).error.code, 'nothing_to_submit');

    instance.addStroke(2, ring('a', 0.3, 0.3));
    assert.deepEqual((await state()).review, { canSubmit: true, submitted: false, slides: [2], drafts: 0 });

    const response = await submit();
    assert.equal(response.status, 200);
    assert.deepEqual(json(response), { submitted: true });
    assert.equal(sent.length, 1);
    assert.match(sent[0].prompt, /"review\.pptx"/);
    assert.match(sent[0].prompt, /instanceId panel-1/);
    assert.match(sent[0].prompt, /pptx-review-canvas/);
    assert.match(sent[0].prompt, /Slides 2/);
    assert.match(sent[0].prompt, /get_annotations/);
    assert.match(sent[0].displayPrompt, /^Review submitted: review\.pptx \(Slides 2\)$/);

    assert.deepEqual((await state()).review, { canSubmit: false, submitted: true, slides: [2], drafts: 0 });
    assert.equal(json(await submit()).error.code, 'already_submitted');
    assert.equal(sent.length, 1, 'the same marks are not sent twice');

    instance.addStroke(3, ring('b', 0.6, 0.6));
    assert.deepEqual((await state()).review, { canSubmit: true, submitted: false, slides: [2, 3], drafts: 0 });

    await instance.annotationReport();
    assert.deepEqual((await state()).review, { canSubmit: false, submitted: false, slides: [], drafts: 0 }, 'once the agent has read the marks there is nothing new');
  } finally {
    await registry.closeAll();
  }
});

test('new slides count as feedback, and the message to the agent carries no free text', async () => {
  const { registry, sent, state, submit } = await setup();
  try {
    await postJson(`${(await registry.ensure('panel-1')).url}api/drafts/add`, { after: 1 });
    assert.deepEqual((await state()).review, { canSubmit: true, submitted: false, slides: [], drafts: 1 });

    assert.equal((await submit()).status, 200);
    assert.match(sent[0].prompt, /New slides: 1/);
    assert.match(sent[0].displayPrompt, /\(New slides: 1\)$/);
    assert.doesNotMatch(sent[0].prompt, /instructions:/, 'the message carries no free text of the user');
  } finally {
    await registry.closeAll();
  }
});

test('the message follows the language that is current', async () => {
  const { registry, instance, sent, submit } = await setup({ language: 'ja' });
  try {
    instance.addStroke(2, ring('a', 0.3, 0.3));
    assert.equal((await submit()).status, 200);
    assert.match(sent[0].prompt, /「review\.pptx」のレビューを終えたので、提出します/);
    assert.match(sent[0].displayPrompt, /^レビューを提出しました: review\.pptx（スライド 2）$/);
  } finally {
    await registry.closeAll();
  }
});

test('without a session to send to, or when sending fails, nothing is marked as submitted', async () => {
  const none = await setup({ sendToAgent: null });
  try {
    none.instance.addStroke(2, ring('a', 0.3, 0.3));
    assert.equal((await none.state()).review.canSubmit, false, 'the button stays off when the viewer cannot reach an agent');
    const refused = await none.submit();
    assert.equal(refused.status, 409);
    assert.equal(json(refused).error.code, 'agent_unavailable');
  } finally {
    await none.registry.closeAll();
  }

  let fail = true;
  const sent = [];
  const flaky = await setup({ sendToAgent: async (message) => { if (fail) throw new Error('session closed'); sent.push(message); } });
  try {
    flaky.instance.addStroke(2, ring('a', 0.3, 0.3));
    const failed = await flaky.submit();
    assert.equal(failed.status, 409);
    assert.match(json(failed).error.message, /session closed/);
    assert.equal((await flaky.state()).review.canSubmit, true, 'a failed send can be tried again');
    fail = false;
    assert.equal((await flaky.submit()).status, 200);
    assert.equal(sent.length, 1);
  } finally {
    await flaky.registry.closeAll();
  }
});

test('the height of the thumbnail strip is saved for every viewer and survives a restart', async () => {
  const { registry, instance, state, dataRoot } = await setup();
  try {
    assert.equal((await state()).railHeight, null);
    assert.deepEqual(json(await postJson(`${instance.url}api/rail-height`, { height: 160 })), { height: 160 });
    assert.equal((await state()).railHeight, 160);
    const other = await registry.ensure('panel-2');
    assert.equal(json(await request(`${other.url}api/state`)).railHeight, 160, 'another viewer of the same session sees it');

    const restarted = new ViewerRegistry({ root: dataRoot, renderer: fakeRenderer(), watchIntervalMs: 0, env: {} });
    const again = await restarted.ensure('panel-1');
    try {
      assert.equal(json(await request(`${again.url}api/state`)).railHeight, 160);
    } finally {
      await restarted.closeAll();
    }

    for (const height of [10, 5000, 120.5, 'tall', true]) {
      assert.equal((await postJson(`${instance.url}api/rail-height`, { height })).status, 400, `height ${JSON.stringify(height)}`);
    }
    assert.equal((await state()).railHeight, 160, 'a refused height changes nothing');
    assert.deepEqual(json(await postJson(`${instance.url}api/rail-height`, { height: null })), { height: null });
    assert.equal((await state()).railHeight, null, 'null goes back to the default');
  } finally {
    await registry.closeAll();
  }
});

test('the width of the comments pane is saved for every viewer and survives a restart', async () => {
  const { registry, instance, state, dataRoot } = await setup();
  try {
    assert.equal((await state()).commentsWidth, null);
    const saved = await postJson(`${instance.url}api/comments-width`, { width: 420 });
    assert.equal(saved.status, 200);
    assert.deepEqual(json(saved), { width: 420 });
    assert.equal((await state()).commentsWidth, 420);
    const other = await registry.ensure('panel-2');
    assert.equal(json(await request(`${other.url}api/state`)).commentsWidth, 420, 'another viewer of the same session sees it');

    const restarted = new ViewerRegistry({ root: dataRoot, renderer: fakeRenderer(), watchIntervalMs: 0, env: {} });
    const again = await restarted.ensure('panel-1');
    try {
      assert.equal(json(await request(`${again.url}api/state`)).commentsWidth, 420);
    } finally {
      await restarted.closeAll();
    }

    for (const width of [50, 5000, 410.5, 'wide', true]) {
      assert.equal((await postJson(`${instance.url}api/comments-width`, { width })).status, 400, `width ${JSON.stringify(width)}`);
    }
    assert.equal((await state()).commentsWidth, 420, 'a refused width changes nothing');
    assert.deepEqual(json(await postJson(`${instance.url}api/comments-width`, { width: null })), { width: null });
    assert.equal((await state()).commentsWidth, null, 'null goes back to the default');
  } finally {
    await registry.closeAll();
  }
});
