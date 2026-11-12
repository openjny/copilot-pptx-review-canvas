// Text of the page in English (the default) and Japanese. Every visible string is a key in `locales/en.js` and `locales/ja.js`;
// markup asks for one with `data-i18n` (text), `data-i18n-title`, `data-i18n-aria-label` or `data-i18n-placeholder`, and code with `t('key')`.
// Nothing here touches the DOM at import time, so the module also loads in Node (the tests do that).

import en from './locales/en.js';
import ja from './locales/ja.js';

export const SUPPORTED_LANGUAGES = ['en', 'ja'];
export const DEFAULT_LANGUAGE = 'en';

const CATALOGS = { en, ja };
const ATTRIBUTES = [
  ['data-i18n-title', 'title'],
  ['data-i18n-aria-label', 'aria-label'],
  ['data-i18n-placeholder', 'placeholder'],
];

let current = DEFAULT_LANGUAGE;

/** 'ja-JP', 'JA' and 'en_US' give 'ja', 'ja' and 'en'; anything else gives null. */
export function normalizeLanguage(value) {
  if (typeof value !== 'string') return null;
  const language = value.trim().toLowerCase().split(/[-_]/)[0];
  return SUPPORTED_LANGUAGES.includes(language) ? language : null;
}

export function getLanguage() {
  return current;
}

/** Switches the language of `t()`; a value that is not supported is ignored. Returns the language now in effect. */
export function setLanguage(language) {
  const next = normalizeLanguage(language);
  if (next) {
    current = next;
    if (typeof document !== 'undefined') document.documentElement.lang = next;
  }
  return current;
}

const lookup = (language, key) => (Object.hasOwn(CATALOGS[language], key) ? CATALOGS[language][key] : undefined);

/** The text for `key` in the current language, else in English, else the key itself. `{name}` in the text is replaced from `params`. */
export function t(key, params) {
  const template = lookup(current, key) ?? lookup(DEFAULT_LANGUAGE, key) ?? key;
  if (!params) return template;
  return template.replace(/\{(\w+)\}/g, (placeholder, name) => (Object.hasOwn(params, name) ? String(params[name] ?? '') : placeholder));
}

/** Like `t()` for a text that depends on a count: uses `keyOne` when `count` is 1 and `keyOther` otherwise, and `{count}` is filled in. */
export function tn(key, count, params) {
  return t(`${key}${count === 1 ? 'One' : 'Other'}`, { count, ...params });
}

function within(root, selector) {
  const found = [...root.querySelectorAll(selector)];
  return root.matches?.(selector) ? [root, ...found] : found;
}

/** Writes the current language into every element of `root` that carries a `data-i18n*` attribute (this includes `<title>`). */
export function applyTranslations(root = document) {
  for (const element of within(root, '[data-i18n]')) element.textContent = t(element.getAttribute('data-i18n'));
  for (const [source, attribute] of ATTRIBUTES) {
    for (const element of within(root, `[${source}]`)) element.setAttribute(attribute, t(element.getAttribute(source)));
  }
}
