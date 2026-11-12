// Server-side i18n: every text that a person or the agent reads comes from a catalog (lib/locales/<language>.mjs), addressed by a
// dotted key such as 'error.slideOutOfRange'. A template takes {name} placeholders.
//
// - createI18n() holds the current language. The ViewerRegistry owns the one instance that every viewer, every store and
//   extension.mjs share, so a language change is seen everywhere at once.
// - Code with no access to that instance (validators, file rewriters, the stores) throws ViewerError(code, messageKey, params).
//   Its `message` is the English text, which keeps logs and test failures readable; messageKey and params let the layer that
//   reports it (the state snapshot, a CanvasError, a log line) render it in the language that is current at that moment. An error
//   held in the state is therefore shown in the new language after a switch, without being created again.
// - A param can itself be a message ({ messageKey, params }: a ViewerError, or msg()), which is rendered in the same language,
//   and an array is joined with the language's list separator.
// - A missing key falls back to the English text and finally to the key itself, so a gap in a catalog never breaks a message.

import en from './locales/en.mjs';
import ja from './locales/ja.mjs';

export const SUPPORTED_LANGUAGES = ['en', 'ja'];
export const DEFAULT_LANGUAGE = 'en';
const CATALOGS = { en, ja };

/** 'en' or 'ja' for a language tag such as 'ja', 'JA-jp' or 'en_US'; null for anything else. */
export function normalizeLanguage(value) {
  if (typeof value !== 'string') return null;
  const primary = value.trim().toLowerCase().split(/[-_]/)[0];
  return SUPPORTED_LANGUAGES.includes(primary) ? primary : null;
}

const isMessage = (value) => Boolean(value) && typeof value === 'object' && typeof value.messageKey === 'string';

/** A message to render later, for text that is built far from the code that knows the language (a validation problem, a nested detail). */
export const msg = (messageKey, params) => ({ messageKey, params });

function templateFor(language, key) {
  for (const catalog of [Object.hasOwn(CATALOGS, language) ? CATALOGS[language] : null, en]) {
    if (catalog && Object.hasOwn(catalog, key)) return catalog[key];
  }
  return key;
}

/** Renders `key` in `language` with no instance needed. */
export function translate(language, key, params) {
  const template = templateFor(language, key);
  if (!params) return template;
  return template.replace(/\{(\w+)\}/g, (placeholder, name) => {
    if (!Object.hasOwn(params, name)) return placeholder;
    const value = params[name];
    if (isMessage(value)) return translate(language, value.messageKey, value.params);
    if (Array.isArray(value)) return value.map(String).join(templateFor(language, 'common.listSeparator'));
    return String(value);
  });
}

/** English only, for code that is given no instance. */
export const tEn = (key, params) => translate('en', key, params);

export function createI18n({ language = DEFAULT_LANGUAGE } = {}) {
  let current = normalizeLanguage(language) ?? DEFAULT_LANGUAGE;
  return {
    get language() {
      return current;
    },
    setLanguage(next) {
      const normalized = normalizeLanguage(next);
      if (!normalized) throw new RangeError(`Unsupported language: ${String(next)}`);
      current = normalized;
      return current;
    },
    t: (key, params) => translate(current, key, params),
    /** The text of an error in the current language: a ViewerError (or any message) is rendered from its key, anything else keeps its own text. */
    messageOf: (error) => (isMessage(error) ? translate(current, error.messageKey, error.params) : String(error?.message ?? error)),
  };
}
