import assert from 'node:assert/strict';
import fs from 'node:fs';
import { afterEach, test } from 'node:test';
import { fileURLToPath } from 'node:url';

import * as i18n from '../assets/i18n.js';
import en from '../assets/locales/en.js';
import ja from '../assets/locales/ja.js';

const assetsDir = fileURLToPath(new URL('../assets/', import.meta.url));
const read = (name) => fs.readFileSync(`${assetsDir}${name}`, 'utf8');
const placeholdersOf = (text) => [...text.matchAll(/\{(\w+)\}/g)].map((match) => match[1]).sort();
const JAPANESE = /[\u3000-\u30ff\u4e00-\u9fff\uff00-\uffef]/;
const SCRIPTS = fs.readdirSync(assetsDir).filter((name) => name.endsWith('.js'));
const CODE_FILES = SCRIPTS.filter((name) => name !== 'i18n.js');
const html = read('index.html');

afterEach(() => {
  i18n.setLanguage('en');
  delete globalThis.document;
});

test('the English and Japanese catalogs have the same keys and placeholders, and no empty text', () => {
  assert.deepEqual(Object.keys(ja).sort(), Object.keys(en).sort());
  for (const [key, text] of Object.entries(en)) {
    assert.ok(text.trim(), `en: ${key} is empty`);
    assert.ok(ja[key].trim(), `ja: ${key} is empty`);
    assert.deepEqual(placeholdersOf(ja[key]), placeholdersOf(text), `placeholders of ${key}`);
  }
});

test('normalizeLanguage accepts the supported languages in any case and region form', () => {
  assert.deepEqual(i18n.SUPPORTED_LANGUAGES, ['en', 'ja']);
  assert.equal(i18n.DEFAULT_LANGUAGE, 'en');
  for (const [value, expected] of [['en', 'en'], ['ja', 'ja'], ['JA', 'ja'], ['ja-JP', 'ja'], ['en_US', 'en'], [' en-gb ', 'en'], ['fr', null], ['', null], ['jav', null], [null, null], [undefined, null], [42, null]]) {
    assert.equal(i18n.normalizeLanguage(value), expected, String(value));
  }
});

test('t() follows the language, fills placeholders and falls back to English and then to the key', () => {
  assert.equal(i18n.getLanguage(), 'en');
  assert.equal(i18n.t('slide.label', { n: 3, title: 'Plan' }), 'Slide 3: Plan');
  assert.equal(i18n.t('slide.label', { n: 3 }), 'Slide 3: {title}');
  assert.equal(i18n.t('slide.untitled'), en['slide.untitled']);
  assert.equal(i18n.t('no.such.key'), 'no.such.key');
  assert.equal(i18n.t('constructor'), 'constructor');

  assert.equal(i18n.setLanguage('ja-JP'), 'ja');
  assert.equal(i18n.t('slide.label', { n: 3, title: 'Plan' }), 'スライド 3: Plan');
  assert.equal(i18n.setLanguage('fr'), 'ja', 'an unsupported language is ignored');

  en['test.onlyEnglish'] = 'Only in English: {value}';
  try {
    assert.equal(i18n.t('test.onlyEnglish', { value: 0 }), 'Only in English: 0');
  } finally {
    delete en['test.onlyEnglish'];
  }
});

test('tn() picks the singular or the plural text and fills {count}', () => {
  assert.equal(i18n.tn('thumb.comments', 1), '1 comment');
  assert.equal(i18n.tn('thumb.comments', 2), '2 comments');
  assert.equal(i18n.tn('thumb.marks', 3, { status: 'x' }), '3 marks (x)');
  i18n.setLanguage('ja');
  assert.equal(i18n.tn('thumb.comments', 1), 'コメント 1 件');
});

test('setLanguage sets the lang of the document, and applyTranslations fills text, title, aria-label and placeholder', () => {
  const makeElement = (attributes) => {
    const element = {
      textContent: 'static',
      attributes: { ...attributes },
      getAttribute: (name) => element.attributes[name] ?? null,
      setAttribute: (name, value) => {
        element.attributes[name] = value;
      },
    };
    return element;
  };
  const elements = [
    makeElement({ 'data-i18n': 'open.submit' }),
    makeElement({ 'data-i18n-title': 'ink.undoTitle', 'data-i18n-aria-label': 'ink.undo' }),
    makeElement({ 'data-i18n-placeholder': 'open.sketchPlaceholder' }),
  ];
  const root = { querySelectorAll: (selector) => elements.filter((element) => element.getAttribute(selector.slice(1, -1)) !== null) };
  globalThis.document = { documentElement: { lang: 'en' }, ...root };

  i18n.setLanguage('ja');
  assert.equal(document.documentElement.lang, 'ja');
  i18n.applyTranslations();
  assert.equal(elements[0].textContent, '開く');
  assert.equal(elements[1].attributes.title, '元に戻す（Ctrl+Z）');
  assert.equal(elements[1].attributes['aria-label'], '元に戻す');
  assert.equal(elements[2].attributes.placeholder, '名前（省略できます）');

  i18n.setLanguage('en');
  assert.equal(document.documentElement.lang, 'en');
  i18n.applyTranslations(root);
  assert.equal(elements[0].textContent, 'Open');
  assert.equal(elements[2].attributes.placeholder, 'Name (optional)');
});

test('index.html is English by default and every data-i18n key exists with the same English text', () => {
  assert.match(html, /<html lang="en">/);
  let checked = 0;
  for (const [, tag, attributes, content] of html.matchAll(/<(\w+)\b([^>]*\bdata-i18n(?:-[a-z-]+)?=[^>]*)>([^<]*)/g)) {
    const found = Object.fromEntries([...attributes.matchAll(/([\w-]+)="([^"]*)"/g)].map((match) => [match[1], match[2]]));
    for (const [name, key] of Object.entries(found)) {
      if (!name.startsWith('data-i18n')) continue;
      assert.ok(Object.hasOwn(en, key), `${tag}: ${name}="${key}" is not in the English catalog`);
      const target = name === 'data-i18n' ? content : found[name.replace('data-i18n-', '')];
      assert.equal(target, en[key], `${tag}: static text of ${key}`);
      checked += 1;
    }
  }
  const total = [...html.matchAll(/\sdata-i18n(?:-[a-z-]+)?="/g)].length;
  assert.ok(total > 0);
  assert.equal(checked, total, 'every data-i18n attribute is checked');
});

test('every key used in code or markup is in the English catalog, and every catalog key is used', () => {
  const used = new Set();
  for (const [, key] of html.matchAll(/data-i18n(?:-[a-z-]+)?="([^"]+)"/g)) used.add(key);
  for (const name of CODE_FILES) {
    const source = read(name);
    assert.doesNotMatch(source, /\b(?:t|tn)\(\s*(?!['"])/, `${name}: t() and tn() need a literal key so that it can be checked`);
    for (const [, fn, key] of source.matchAll(/\b(t|tn)\(\s*'([^']+)'/g)) {
      if (fn === 'tn') {
        assert.ok(Object.hasOwn(en, `${key}One`) && Object.hasOwn(en, `${key}Other`), `${name}: tn('${key}') needs ${key}One and ${key}Other`);
        used.add(`${key}One`).add(`${key}Other`);
      } else {
        assert.ok(Object.hasOwn(en, key), `${name}: t('${key}') is not in the English catalog`);
        used.add(key);
      }
    }
  }
  assert.deepEqual([...used].filter((key) => !Object.hasOwn(en, key)), []);
  assert.deepEqual(Object.keys(en).filter((key) => !used.has(key)), [], 'catalog keys nothing uses');
});

test('no Japanese is left in the markup and the code; only ja.js and the language name hold it', () => {
  for (const name of ['index.html', ...SCRIPTS, 'locales/en.js']) {
    const source = read(name).replace('>日本語</button>', '>Japanese</button>');
    assert.doesNotMatch(source, JAPANESE, name);
  }
  assert.match(html, /data-lang="ja"[^>]*>日本語<\/button>/);
});
