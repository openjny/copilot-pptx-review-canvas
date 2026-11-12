import assert from 'node:assert/strict';
import http from 'node:http';
import path from 'node:path';
import { test } from 'node:test';
import { ViewerError } from '../lib/errors.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { fakeRenderer, json, makeDeckFile, makeTempDir, postJson, request } from './helpers.mjs';

async function setup(options = {}) {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const registry = new ViewerRegistry({ root, renderer: fakeRenderer(options) });
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  const file = await makeDeckFile(decks, 'sample.pptx');
  return { registry, instance, decks, file };
}

test('serves the page with a strict policy and starts empty', async () => {
  const { registry, instance } = await setup();
  try {
    const page = await request(instance.url);
    assert.equal(page.status, 200);
    assert.match(page.headers['content-type'], /text\/html/);
    assert.match(page.headers['content-security-policy'], /default-src 'none'/);
    assert.equal(json(await request(`${instance.url}api/state`)).phase, 'empty');
    assert.equal((await request(`${instance.url}missing`)).status, 404);
  } finally {
    await registry.closeAll();
  }
});

test('opens a deck, serves its images and keeps the current slide in sync', async () => {
  const { registry, instance, file } = await setup({ slides: 3 });
  try {
    await instance.openDeck(file);
    const summary = await instance.whenSettled(5000);
    assert.equal(summary.phase, 'ready');
    assert.equal(summary.slideCount, 3);

    const state = json(await request(`${instance.url}api/state`));
    assert.equal(state.deck.name, 'sample.pptx');
    assert.equal(state.deck.slides[1].title, 'Title 2');
    assert.equal(state.current, 1);

    const image = await request(`${instance.url}deck/${state.deck.key}/s0002.png`);
    assert.equal(image.status, 200);
    assert.equal(image.headers['content-type'], 'image/png');
    assert.equal((await request(`${instance.url}deck/${state.deck.key}/s0009.png`)).status, 404);
    assert.equal((await request(`${instance.url}deck/${state.deck.key}/..%2Fmanifest.json`)).status, 404);

    const moved = await postJson(`${instance.url}api/current`, { slide: 3 });
    assert.equal(moved.status, 200);
    assert.equal(json(moved).title, 'Title 3');
    assert.equal(instance.summary().current, 3);
    assert.equal((await postJson(`${instance.url}api/current`, { slide: 4 })).status, 400);
    assert.throws(() => instance.setCurrent(0), (error) => error instanceof ViewerError && error.code === 'slide_out_of_range');

    assert.equal(instance.outline({ includeNotes: true }).slides[0].notes, 'Notes 1\nsecond line');
    assert.ok((await instance.slideImage(2)).length > 0);
  } finally {
    await registry.closeAll();
  }
});

test('refuses foreign hosts, foreign origins and non-JSON bodies', async () => {
  const { registry, instance, file } = await setup();
  try {
    const port = new URL(instance.url).port;
    assert.equal((await request(`${instance.url}api/state`, { headers: { Host: `evil.example:${port}` } })).status, 403);
    assert.equal((await request(`${instance.url}api/state`, { headers: { Host: '127.0.0.1:1' } })).status, 403);
    assert.equal((await postJson(`${instance.url}api/open`, { path: file }, { Origin: 'http://evil.example' })).status, 403);
    assert.equal((await request(`${instance.url}api/open`, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ path: file }) })).status, 415);
    assert.equal((await postJson(`${instance.url}api/open`, { path: path.join(path.dirname(file), 'nope.pptx') })).status, 404);
    assert.equal((await postJson(`${instance.url}api/open`, { path: 5 })).status, 400);
    const sameOrigin = await postJson(`${instance.url}api/open`, { path: file }, { Origin: `http://127.0.0.1:${port}` });
    assert.equal(sameOrigin.status, 202);
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
  } finally {
    await registry.closeAll();
  }
});

test('pushes state to event-stream clients and restores the slide the deck was left on', async () => {
  const { registry, instance, file } = await setup({ slides: 4 });
  try {
    await instance.openDeck(file);
    await instance.whenSettled(5000);
    instance.setCurrent(4);
    await instance._flushPosition();

    const received = await new Promise((resolve, reject) => {
      const target = new URL(`${instance.url}api/events`);
      const req = http.get({ host: target.hostname, port: target.port, path: target.pathname }, (res) => {
        let text = '';
        res.setEncoding('utf8');
        res.on('data', (chunk) => {
          text += chunk;
          if (text.includes('"current":4') && text.includes('event: state')) {
            req.destroy();
            resolve(text);
          }
        });
      });
      req.on('error', (error) => (error.code === 'ECONNRESET' ? undefined : reject(error)));
    });
    assert.match(received, /event: state/);

    await registry.close('panel-1');
    const reopened = await registry.ensure('panel-1', { baseDir: path.dirname(file) });
    await reopened.openDeck(file);
    const summary = await reopened.whenSettled(5000);
    assert.equal(summary.current, 4);
  } finally {
    await registry.closeAll();
  }
});

test('whenReady waits for a render that is still running and gives up after the timeout', async () => {
  const slow = await setup({ slides: 3, delayMs: 40 });
  try {
    await slow.instance.openDeck(slow.file);
    assert.throws(() => slow.instance.outline(), (error) => error.code === 'no_deck');
    const ready = await slow.instance.whenReady(5000);
    assert.equal(ready, slow.instance);
    assert.equal(ready.outline().slideCount, 3);
  } finally {
    await slow.registry.closeAll();
  }

  const stuck = await setup({ slides: 3, delayMs: 300 });
  try {
    await stuck.instance.openDeck(stuck.file);
    await assert.rejects(stuck.instance.whenReady(30), (error) => error instanceof ViewerError && error.code === 'still_rendering');
  } finally {
    await stuck.registry.closeAll();
  }
});

test('whenReady reports a failed render only when there is no deck to read', async () => {
  const failing = await setup({ failWith: new ViewerError('render_failed', 'PowerPoint is unavailable') });
  try {
    await failing.instance.openDeck(failing.file);
    await assert.rejects(failing.instance.whenReady(5000), (error) => error instanceof ViewerError && error.code === 'render_failed');
  } finally {
    await failing.registry.closeAll();
  }

  const root = await makeTempDir();
  const decks = await makeTempDir();
  let broken = false;
  const working = fakeRenderer({ slides: 2 });
  const registry = new ViewerRegistry({ root, renderer: (args) => (broken ? Promise.reject(new ViewerError('render_failed', 'boom')) : working(args)) });
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    await instance.openDeck(await makeDeckFile(decks, 'sample.pptx'));
    await instance.whenReady(5000);
    broken = true;
    await instance.refresh();
    assert.equal((await instance.whenSettled(5000)).phase, 'error');
    assert.equal((await instance.whenReady(5000)).outline().slideCount, 2);
  } finally {
    await registry.closeAll();
  }
});

test('reports a failed render as an error state and reopening the same deck does not render again', async () => {
  const failing = await setup({ failWith: new ViewerError('render_failed', 'PowerPoint is unavailable') });
  try {
    await failing.instance.openDeck(failing.file);
    const summary = await failing.instance.whenSettled(5000);
    assert.equal(summary.phase, 'error');
    assert.equal(summary.error.code, 'render_failed');
    assert.throws(() => failing.instance.slideInfo(), (error) => error.code === 'no_deck');
  } finally {
    await failing.registry.closeAll();
  }

  const calls = [];
  const ok = await setup({ calls });
  try {
    await ok.instance.openDeck(ok.file);
    await ok.instance.whenSettled(5000);
    await ok.instance.openDeck(ok.file, { slide: 2 });
    assert.equal(ok.instance.summary().current, 2);
    assert.equal(calls.length, 1);
    const again = await ok.registry.ensure('panel-1');
    assert.equal(again, ok.instance);
  } finally {
    await ok.registry.closeAll();
  }
});
