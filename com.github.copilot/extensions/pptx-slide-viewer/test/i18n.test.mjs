import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { readdirSync, readFileSync } from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { AnnotationStore, validateStroke } from '../lib/annotations.mjs';
import { ViewerError } from '../lib/errors.mjs';
import { DEFAULT_LANGUAGE, SUPPORTED_LANGUAGES, createI18n, msg, normalizeLanguage, translate } from '../lib/i18n.mjs';
import en from '../lib/locales/en.mjs';
import ja from '../lib/locales/ja.mjs';
import { analyzeSlide } from '../lib/marks.mjs';
import { suggestNewDeckPath } from '../lib/new-deck.mjs';
import { ViewerRegistry } from '../lib/registry.mjs';
import { sketchbookName, sketchbooksDir } from '../lib/sketchbook.mjs';
import { fakeRenderer, json, makeDeckFile, makeTempDir, postJson, request } from './helpers.mjs';

const EXTENSION_DIR = fileURLToPath(new URL('..', import.meta.url));
const sourceFiles = [
  path.join(EXTENSION_DIR, 'extension.mjs'),
  ...readdirSync(path.join(EXTENSION_DIR, 'lib')).filter((name) => name.endsWith('.mjs')).map((name) => path.join(EXTENSION_DIR, 'lib', name)),
];
const placeholders = (text) => [...text.matchAll(/\{(\w+)\}/g)].map((match) => match[1]).sort();

async function until(condition, ms = 4000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  assert.fail('condition was not met in time');
}

/** Listens to the event stream of a viewer and keeps every state it is sent. */
function listen(instance) {
  const states = [];
  const target = new URL(`${instance.url}api/events`);
  const req = http.get({ host: target.hostname, port: target.port, path: target.pathname }, (res) => {
    let pending = '';
    res.setEncoding('utf8');
    res.on('data', (chunk) => {
      pending += chunk;
      const frames = pending.split('\n\n');
      pending = frames.pop();
      for (const frame of frames) {
        const data = /^data: (.*)$/m.exec(frame);
        if (/^event: state$/m.test(frame) && data) states.push(JSON.parse(data[1]));
      }
    });
  });
  req.on('error', () => {});
  return { states, stop: () => req.destroy() };
}

async function setup({ env = {}, root, renderer = fakeRenderer({ slides: 2 }) } = {}) {
  const dataRoot = root ?? (await makeTempDir());
  const decks = await makeTempDir();
  const registry = new ViewerRegistry({ root: dataRoot, env, renderer, watchIntervalMs: 0 });
  return { registry, dataRoot, decks };
}

test('the English and Japanese catalogs have the same keys and the same placeholders, and no text is empty', () => {
  assert.deepEqual(Object.keys(ja).sort(), Object.keys(en).sort());
  for (const key of Object.keys(en)) {
    assert.ok(en[key].trim().length > 0, `en: ${key} is empty`);
    assert.ok(ja[key].trim().length > 0, `ja: ${key} is empty`);
    assert.deepEqual(placeholders(ja[key]), placeholders(en[key]), `${key}: the placeholders differ`);
  }
});

test('every key the code uses is in the English catalog, and every English key is used', () => {
  const areas = [...new Set(Object.keys(en).map((key) => key.split('.')[0]))];
  const literal = new RegExp(`'((?:${areas.join('|')})\\.[A-Za-z0-9.]+)'`, 'g');
  const used = new Set();
  for (const file of sourceFiles) {
    for (const match of readFileSync(file, 'utf8').matchAll(literal)) {
      if (!/\.(json|mjs|js|png|xml|rels|flag)$/.test(match[1])) used.add(match[1]);
    }
  }
  assert.deepEqual([...used].filter((key) => !Object.hasOwn(en, key)).sort(), [], 'keys the code uses but the catalog lacks');
  assert.deepEqual(Object.keys(en).filter((key) => !used.has(key)).sort(), [], 'catalog keys nothing uses');
});

test('no Japanese is written in the code outside the Japanese catalog and the names of legacy data', () => {
  const japanese = /[\u3000-\u30ff\u4e00-\u9fff\uff00-\uffef]/;
  const found = [];
  for (const file of sourceFiles) {
    readFileSync(file, 'utf8').split(/\r?\n/).forEach((line, index) => {
      if (japanese.test(line) && !/LEGACY_DRAFT_TITLE|Japanese-only versions/.test(line)) found.push(`${path.basename(file)}:${index + 1}`);
    });
  }
  assert.deepEqual(found, []);
});

test('normalizeLanguage accepts supported languages in any case or region form and nothing else', () => {
  assert.deepEqual(SUPPORTED_LANGUAGES, ['en', 'ja']);
  assert.equal(DEFAULT_LANGUAGE, 'en');
  for (const [value, expected] of [['en', 'en'], ['ja', 'ja'], ['EN', 'en'], ['Ja', 'ja'], ['ja-JP', 'ja'], ['en-US', 'en'], ['en_GB', 'en'], [' ja ', 'ja']]) {
    assert.equal(normalizeLanguage(value), expected, String(value));
  }
  for (const value of ['', 'fr', 'jp', 'english', 'ja JP', undefined, null, 5, {}, ['ja']]) assert.equal(normalizeLanguage(value), null, String(value));
});

test('t looks the key up in the current language, fills placeholders, and falls back to English and then to the key', () => {
  const i18n = createI18n();
  assert.equal(i18n.language, 'en');
  assert.equal(i18n.t('status.slideCount', { count: 3 }), '3 slides');
  assert.equal(i18n.t('status.error', {}), 'Error: {message}', 'a placeholder without a value is left as it is');
  assert.equal(i18n.t('no.such.key'), 'no.such.key');

  assert.equal(i18n.setLanguage('ja-JP'), 'ja');
  assert.equal(i18n.language, 'ja');
  assert.equal(i18n.t('status.slideCount', { count: 3 }), '3 枚のスライド');
  assert.throws(() => i18n.setLanguage('fr'), RangeError);
  assert.equal(i18n.language, 'ja', 'an unsupported language changes nothing');
  assert.equal(createI18n({ language: 'fr' }).language, 'en');
  assert.equal(createI18n({ language: 'ja' }).language, 'ja');

  en['test.onlyEnglish'] = 'only {name}';
  try {
    assert.equal(i18n.t('test.onlyEnglish', { name: 'here' }), 'only here', 'a key missing from the Japanese catalog is shown in English');
    assert.equal(translate('xx', 'test.onlyEnglish', { name: 'x' }), 'only x', 'so is any other language');
  } finally {
    delete en['test.onlyEnglish'];
  }
});

test('a param can be a message or a list, and both are rendered in the language of the text around them', () => {
  const nested = { where: 'slides[1]', detail: new ViewerError('invalid_input', 'stroke.idInvalid') };
  assert.equal(translate('en', 'record.strokesInvalid', nested), 'Invalid strokes in slides[1] (stroke.id is not valid.)');
  assert.equal(translate('ja', 'record.strokesInvalid', nested), 'slides[1] のストロークが不正です（stroke.id が不正です。）');
  assert.equal(translate('ja', 'error.invalidPptx', { detail: msg('zip.tooSmall') }), 'PowerPoint のファイルを読めませんでした（ファイルが小さすぎます）。');

  assert.equal(translate('en', 'stroke.toolInvalid', { tools: ['pen', 'highlighter'] }), 'stroke.tool must be one of: pen, highlighter.');
  assert.equal(translate('ja', 'stroke.toolInvalid', { tools: ['pen', 'highlighter'] }), 'stroke.tool は pen、highlighter のいずれかにしてください。');
  assert.equal(translate('ja', 'unread.slides', { list: [2, 3, 5] }), 'スライド 2、3、5');
});

test('a ViewerError keeps its English text and carries the key for the language in use', () => {
  const error = new ViewerError('slide_out_of_range', 'error.slideOutOfRange', { count: 5 });
  assert.equal(error.code, 'slide_out_of_range');
  assert.equal(error.message, 'The slide number must be between 1 and 5.');
  assert.deepEqual([error.messageKey, error.params], ['error.slideOutOfRange', { count: 5 }]);
  const i18n = createI18n({ language: 'ja' });
  assert.equal(i18n.messageOf(error), 'スライド番号は 1 から 5 の範囲で指定してください。');
  assert.equal(i18n.messageOf(new Error('plain')), 'plain', 'an error without a key keeps its own text');
  assert.equal(new ViewerError('x', 'a text that is no key').message, 'a text that is no key');

  assert.throws(() => validateStroke(null), (thrown) => thrown.code === 'invalid_input' && thrown.messageKey === 'stroke.notObject');
});

test('names that are made on the spot follow the language, and Japanese names made before still work', async () => {
  const at = new Date(2024, 0, 2, 3, 4, 5);
  const japanese = createI18n({ language: 'ja' });
  assert.equal(sketchbookName('', at), 'Sketch 20240102-030405');
  assert.equal(sketchbookName('', at, japanese.t), 'スケッチ 20240102-030405');
  assert.equal(sketchbookName('Ideas', at, japanese.t), 'Ideas', 'a name given by the user is never translated');

  const dir = await makeTempDir();
  assert.equal(await suggestNewDeckPath(dir, japanese.t), path.join(dir, '新しい資料.pptx'));
  assert.equal(await suggestNewDeckPath(dir), path.join(dir, 'New deck.pptx'));

  const line = { id: 'u', tool: 'pen', color: '#e5231b', width: 0.004, points: [[0.06, 0.34], [0.2, 0.345], [0.43, 0.34]] };
  const marks = (t) => analyzeSlide({ strokes: [line], aspect: 16 / 9, shapes: null, slideW: 960, slideH: 540, t }).marks;
  assert.equal(marks(japanese.t)[0].region, '左中');
  assert.equal(marks(createI18n().t)[0].region, 'middle left');
});

test('POST /api/language saves the choice, switches every viewer at once and is remembered by the next registry', async () => {
  const { registry, dataRoot, decks } = await setup();
  const first = await registry.ensure('panel-1', { baseDir: decks });
  const second = await registry.ensure('panel-2', { baseDir: decks });
  const stream = listen(second);
  try {
    assert.equal(json(await request(`${first.url}api/state`)).language, 'en');
    assert.equal(first.title, 'PowerPoint Viewer');

    const response = await postJson(`${first.url}api/language`, { language: 'ja-JP' });
    assert.equal(response.status, 200);
    assert.deepEqual(json(response), { language: 'ja' }, 'the answer is the normalized language');
    assert.equal(json(await request(`${first.url}api/state`)).language, 'ja');
    assert.equal(json(await request(`${second.url}api/state`)).language, 'ja');
    assert.equal(first.title, 'PowerPoint ビューア');
    assert.equal(registry.language, 'ja');
    await until(() => stream.states.at(-1)?.language === 'ja');
    assert.equal(stream.states.at(-1).rev > stream.states[0].rev, true, 'the other viewer was sent a new state');

    const saved = JSON.parse(await fs.readFile(path.join(dataRoot, 'settings.json'), 'utf8'));
    assert.deepEqual(saved, { format: 'pptx-slide-viewer.settings', version: 1, data: { language: 'ja' } });
    assert.deepEqual(json(await postJson(`${second.url}api/language`, { language: 'en' })), { language: 'en' });
    assert.equal(json(await request(`${first.url}api/state`)).language, 'en', 'any viewer can change it');
    assert.equal(json(await postJson(`${second.url}api/language`, { language: 'JA' })).language, 'ja');
  } finally {
    stream.stop();
    await registry.closeAll();
  }

  // The saved choice wins over the environment, and an environment default applies only where nothing was saved.
  const next = new ViewerRegistry({ root: dataRoot, env: { PPTX_CANVAS_LANG: 'en' }, renderer: fakeRenderer(), watchIntervalMs: 0 });
  try {
    await next.ready();
    assert.equal(next.language, 'ja');
    assert.equal(json(await request(`${(await next.ensure('panel-1', { baseDir: decks })).url}api/state`)).language, 'ja');
  } finally {
    await next.closeAll();
  }
  const fresh = new ViewerRegistry({ root: await makeTempDir(), env: { PPTX_CANVAS_LANG: 'ja_JP' }, renderer: fakeRenderer(), watchIntervalMs: 0 });
  await fresh.ready();
  assert.equal(fresh.language, 'ja', 'PPTX_CANVAS_LANG is the default when nothing is saved');
  const plain = new ViewerRegistry({ root: await makeTempDir(), env: { PPTX_CANVAS_LANG: 'klingon' }, renderer: fakeRenderer(), watchIntervalMs: 0 });
  await plain.ready();
  assert.equal(plain.language, 'en', 'and an unknown value is ignored');
});

test('POST /api/language refuses an unsupported language and the usual unsafe requests, and changes nothing', async () => {
  const { registry, dataRoot, decks } = await setup();
  const instance = await registry.ensure('panel-1', { baseDir: decks });
  try {
    for (const body of [{ language: 'fr' }, { language: '' }, { language: 5 }, { language: null }, {}]) {
      const response = await postJson(`${instance.url}api/language`, body);
      assert.equal(response.status, 400, JSON.stringify(body));
      assert.equal(json(response).error.code, 'invalid_input');
    }
    assert.equal(json(await postJson(`${instance.url}api/language`, { language: 'fr' })).error.message, 'language must be one of: en, ja.');
    assert.equal((await postJson(`${instance.url}api/language`, { language: 'ja' }, { Origin: 'http://evil.example' })).status, 403);
    assert.equal((await request(`${instance.url}api/language`, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: '{"language":"ja"}' })).status, 415);
    assert.equal((await request(`${instance.url}api/language`)).status, 404, 'it is not a GET route');
    assert.equal(json(await request(`${instance.url}api/state`)).language, 'en');
    assert.equal(await fs.access(path.join(dataRoot, 'settings.json')).then(() => true, () => false), false, 'nothing was saved');
    await assert.rejects(() => registry.setLanguage('fr'), (error) => error.code === 'invalid_input' && error.messageKey === 'input.languageInvalid');
  } finally {
    await registry.closeAll();
  }
});

test('the texts in the state follow the language: errors, status, notices, history labels and page titles', async () => {
  const failing = await setup({ renderer: fakeRenderer({ failWith: new ViewerError('render_failed', 'error.renderTimeout') }) });
  try {
    const instance = await failing.registry.ensure('panel-1', { baseDir: failing.decks });
    await instance.openDeck(await makeDeckFile(failing.decks, 'broken.pptx'));
    assert.equal((await instance.whenSettled(5000)).phase, 'error');
    assert.deepEqual(instance.snapshot().error, { code: 'render_failed', message: 'Rendering the slides did not finish in time.' });
    assert.equal(instance.statusText, 'Error: Rendering the slides did not finish in time.');

    await postJson(`${instance.url}api/language`, { language: 'ja' });
    assert.deepEqual(json(await request(`${instance.url}api/state`)).error, { code: 'render_failed', message: 'スライドの描画が時間内に終わりませんでした。' }, 'the error that was already in the state is shown in the new language');
    assert.equal(instance.statusText, 'エラー: スライドの描画が時間内に終わりませんでした。');
    assert.equal(instance.summary().error.message, 'スライドの描画が時間内に終わりませんでした。');
    await assert.rejects(instance.whenReady(5000), (error) => error.code === 'render_failed' && instance.i18n.messageOf(error) === 'スライドの描画が時間内に終わりませんでした。');
    const refused = await postJson(`${instance.url}api/slides/move`, { from: 'x', to: 2 });
    assert.equal(json(refused).error.message, 'from と to を整数で指定してください。', 'so are the answers to the page');
  } finally {
    await failing.registry.closeAll();
  }

  const { registry, decks } = await setup({ renderer: fakeRenderer({ slides: 2 }) });
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    assert.equal(instance.statusText, 'Open a deck to get started');
    const file = await makeDeckFile(decks, 'sample.pptx', 'v1');
    await instance.openDeck(file);
    await instance.whenSettled(5000);
    assert.equal(instance.statusText, '2 slides');
    assert.throws(() => instance.setCurrent(9), (error) => registry.i18n.messageOf(error) === 'The slide number must be between 1 and 2.');

    const comment = instance.addComment(1, { text: 'check this' }).comment;
    instance.removeComment(1, comment.id);
    assert.equal(instance.snapshot().history.undo.label, 'Delete comment');
    await registry.setLanguage('ja');
    assert.equal(instance.snapshot().history.undo.label, 'コメントの削除', 'the label of an action that was already recorded follows the language too');
    await registry.setLanguage('en');

    await fs.writeFile(file, 'v2 is a different size');
    await instance.openDeck(file);
    await instance.whenSettled(5000);
    assert.equal(instance.snapshot().notice.text, 'Rendered the deck again.');
    const noticeId = instance.snapshot().notice.id;

    await registry.setLanguage('ja');
    const state = instance.snapshot();
    assert.equal(state.language, 'ja');
    assert.equal(instance.statusText, '2 枚のスライド');
    assert.deepEqual(state.notice, { id: noticeId, text: '再描画しました。' }, 'the same notice, in the new language');
    assert.throws(() => instance.setCurrent(9), (error) => registry.i18n.messageOf(error) === 'スライド番号は 1 から 2 の範囲で指定してください。');

    await instance.newSketchbook('');
    await instance.whenSettled(5000);
    assert.match(instance.snapshot().deck.name, /^スケッチ \d{8}-\d{6}$/, 'a sketchbook made now is named in the language in use');
    assert.equal(instance.snapshot().deck.slides[0].title, '白紙のページ');
    assert.equal(instance.slideInfo(1).title, '白紙のページ');
    await registry.setLanguage('en');
    assert.equal(instance.snapshot().deck.slides[0].title, 'Blank page');
    assert.equal(instance.outline().slides[0].title, 'Blank page');
  } finally {
    await registry.closeAll();
  }
});

test('data written by the Japanese-only versions still opens: sketchbooks, draft titles and cached page titles', async () => {
  const { registry, dataRoot, decks } = await setup();
  const sketchPath = path.join(sketchbooksDir(dataRoot), 'スケッチ 20240102-030405.sketch');
  await fs.mkdir(path.dirname(sketchPath), { recursive: true });
  await fs.writeFile(sketchPath, `${JSON.stringify({ format: 'pptx-slide-viewer.sketchbook', version: 1, createdAt: 1704160000000 })}\n`);
  const records = new AnnotationStore(dataRoot);
  await fs.mkdir(records.dir, { recursive: true });
  const draft = (id, title, after) => ({ id, title, after, afterHash: '', strokes: [], comments: [], createdAt: 1704160000000, unread: true });
  await fs.writeFile(records.fileFor(sketchPath), JSON.stringify({
    format: 'pptx-slide-viewer.annotations',
    version: 1,
    data: { path: sketchPath, deckKey: 'a'.repeat(24), updatedAt: 1704160000000, draftSeq: 2, slides: {}, drafts: [draft('d00000000000a', '新しいスライド 1', 0), draft('d00000000000b', '新規スライド案 2', 1)] },
  }));
  try {
    const instance = await registry.ensure('panel-1', { baseDir: decks });
    await instance.openDeck(sketchPath);
    assert.equal((await instance.whenSettled(5000)).phase, 'ready');
    let state = instance.snapshot();
    assert.deepEqual([state.deck.kind, state.deck.name, state.deck.slides[0].title], ['sketchbook', 'スケッチ 20240102-030405', 'Blank page']);
    assert.deepEqual(state.annotations.drafts.map((entry) => entry.title), ['New slide 1', 'New slide 2']);
    assert.deepEqual(instance.unreadAnnotations().drafts, ['New slide 1', 'New slide 2']);
    assert.equal(instance.addDraft({ after: 1 }).draft.title, 'New slide 3', 'numbering goes on after the old titles');

    // The page title that the Japanese-only versions cached in the rendered manifest is not what is shown now.
    const manifestFile = path.join(registry.deckCache.keyDir(state.deck.key), 'manifest.json');
    const manifest = JSON.parse(await fs.readFile(manifestFile, 'utf8'));
    manifest.data.slides[0].title = '白紙のページ';
    await fs.writeFile(manifestFile, JSON.stringify(manifest));
    await registry.close('panel-1');
    const reopened = await registry.ensure('panel-1', { baseDir: decks });
    await reopened.openDeck(sketchPath);
    await reopened.whenSettled(5000);
    state = reopened.snapshot();
    assert.equal(state.deck.slides[0].title, 'Blank page');
    await registry.setLanguage('ja');
    assert.equal(reopened.snapshot().deck.slides[0].title, '白紙のページ');
  } finally {
    await registry.closeAll();
  }
});
