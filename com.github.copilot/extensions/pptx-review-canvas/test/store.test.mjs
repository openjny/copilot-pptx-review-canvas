import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { test } from 'node:test';
import { PositionStore, SettingsStore, dataRootFor, hashPath, readRecord, resolveCopilotHome, writeFileAtomic, writeRecord } from '../lib/store.mjs';
import { makeDeckFile, makeTempDir } from './helpers.mjs';

test('resolveCopilotHome uses COPILOT_HOME exactly as given, else ~/.copilot', () => {
  assert.equal(resolveCopilotHome({ COPILOT_HOME: 'c:\\Custom\\Home' }, 'C:\\Users\\x'), 'c:\\Custom\\Home');
  assert.equal(resolveCopilotHome({ COPILOT_HOME: '' }, path.join('/home', 'x')), path.join('/home', 'x', '.copilot'));
  assert.equal(resolveCopilotHome({}, path.join('/home', 'x')), path.join('/home', 'x', '.copilot'));
  assert.equal(dataRootFor('/h'), path.join('/h', 'extension-data', 'pptx-review-canvas'));
});

test('writeFileAtomic replaces the target and leaves no temporary file', async () => {
  const dir = await makeTempDir();
  const target = path.join(dir, 'nested', 'a.txt');
  await writeFileAtomic(target, 'one');
  await writeFileAtomic(target, 'two');
  assert.equal(await fs.readFile(target, 'utf8'), 'two');
  assert.deepEqual(await fs.readdir(path.dirname(target)), ['a.txt']);
});

test('readRecord reports missing, ok and every kind of invalid record', async () => {
  const dir = await makeTempDir();
  const file = path.join(dir, 'r.json');
  const options = { format: 'f', version: 1, validate: (data) => (data.value === 1 ? null : 'bad value') };
  assert.deepEqual(await readRecord(file, options), { status: 'missing' });
  await writeRecord(file, 'f', 1, { value: 1 });
  assert.deepEqual(await readRecord(file, options), { status: 'ok', data: { value: 1 } });
  await writeRecord(file, 'f', 1, { value: 2 });
  assert.equal((await readRecord(file, options)).reason, 'bad value');
  await writeRecord(file, 'other', 1, { value: 1 });
  assert.equal((await readRecord(file, options)).status, 'invalid');
  await writeRecord(file, 'f', 2, { value: 1 });
  assert.deepEqual((await readRecord(file, options)).reason, { messageKey: 'record.unsupportedVersion', params: { version: '2' } });
  await fs.writeFile(file, '{ not json');
  assert.equal((await readRecord(file, options)).status, 'invalid');
});

test('PositionStore remembers slides per deck, lists recent decks and keeps corrupt files aside', async () => {
  const root = await makeTempDir();
  const decks = await makeTempDir();
  const a = await makeDeckFile(decks, 'a.pptx');
  const b = await makeDeckFile(decks, 'b.pptx');
  const logs = [];
  const store = new PositionStore(root, { log: (message) => logs.push(message) });

  assert.equal(await store.load(a), null);
  await store.save(a, 3);
  await new Promise((resolve) => setTimeout(resolve, 5));
  await store.save(b, 7);
  assert.equal((await store.load(a)).slide, 3);
  assert.equal((await store.load(b)).slide, 7);
  assert.deepEqual((await store.recent()).map((entry) => entry.name), ['b.pptx', 'a.pptx']);

  await fs.rm(a);
  assert.deepEqual((await store.recent()).map((entry) => entry.name), ['b.pptx']);

  await fs.writeFile(store.fileFor(b), '{ broken');
  assert.equal(await store.load(b), null);
  assert.equal(logs.length, 1);
  const kept = (await fs.readdir(store.dir)).filter((name) => name.includes('.corrupt-'));
  assert.equal(kept.length, 1);
  assert.notEqual(hashPath(a), hashPath(b));
});

test('SettingsStore saves the language in the order it was asked and keeps an unusable file aside', async () => {
  const root = await makeTempDir();
  const logs = [];
  const store = new SettingsStore(root, { log: (message) => logs.push(message) });
  assert.deepEqual(await store.load(), { language: null, commentsWidth: null, railHeight: null, theme: null });

  await Promise.all([store.save({ language: 'ja' }), store.save({ language: 'en-US' }), store.save({ language: 'ja-JP' })]);
  assert.deepEqual(await store.load(), { language: 'ja', commentsWidth: null, railHeight: null, theme: null });
  assert.deepEqual(JSON.parse(await fs.readFile(path.join(root, 'settings.json'), 'utf8')).data, { language: 'ja' });

  await fs.writeFile(store.file, JSON.stringify({ format: 'pptx-review-canvas.settings', version: 1, data: { language: 'fr' } }));
  assert.deepEqual(await store.load(), { language: null, commentsWidth: null, railHeight: null, theme: null });
  assert.equal(logs.length, 1);
  assert.deepEqual((await fs.readdir(root)).filter((name) => name.includes('.corrupt-')).length, 1);
  assert.equal(await fs.access(store.file).then(() => true, () => false), false, 'the next save starts from a clean file');
});

test('SettingsStore keeps the sizes of the comments pane and the thumbnail strip and the theme beside the language, and ignores values that make no sense', async () => {
  const root = await makeTempDir();
  const store = new SettingsStore(root);
  const saved = async () => JSON.parse(await fs.readFile(store.file, 'utf8')).data;
  await store.load();

  await store.save({ commentsWidth: 420 });
  assert.deepEqual(await saved(), { commentsWidth: 420 }, 'no language is written until the user chooses one');
  await store.save({ language: 'ja' });
  assert.deepEqual(await new SettingsStore(root).load(), { language: 'ja', commentsWidth: 420, railHeight: null, theme: null });

  await store.save({ commentsWidth: 99999 });
  await store.save({ commentsWidth: 'wide' });
  assert.deepEqual(await saved(), { language: 'ja', commentsWidth: 420 }, 'a width out of range is not saved');
  await store.save({ commentsWidth: null });
  assert.deepEqual(await saved(), { language: 'ja' }, 'null goes back to the default');

  await store.save({ railHeight: 150 });
  assert.deepEqual(await saved(), { language: 'ja', railHeight: 150 });
  await store.save({ railHeight: 5 });
  await store.save({ railHeight: 'tall' });
  assert.deepEqual(await saved(), { language: 'ja', railHeight: 150 }, 'a height out of range is not saved');
  await store.save({ railHeight: null });
  assert.deepEqual(await saved(), { language: 'ja' }, 'null goes back to the default');

  await store.save({ theme: 'dark' });
  assert.deepEqual(await saved(), { language: 'ja', theme: 'dark' });
  await store.save({ theme: 'sepia' });
  await store.save({ theme: null });
  assert.deepEqual(await saved(), { language: 'ja', theme: 'dark' }, 'a theme that does not exist is not saved');
  assert.deepEqual(await new SettingsStore(root).load(), { language: 'ja', commentsWidth: null, railHeight: null, theme: 'dark' });

  await fs.writeFile(store.file, JSON.stringify({ format: 'pptx-review-canvas.settings', version: 1, data: { language: 'en', commentsWidth: 5, railHeight: 1, theme: 'neon' } }));
  assert.deepEqual(await new SettingsStore(root).load(), { language: 'en', commentsWidth: null, railHeight: null, theme: null }, 'a bad size or theme does not cost the language');
});
