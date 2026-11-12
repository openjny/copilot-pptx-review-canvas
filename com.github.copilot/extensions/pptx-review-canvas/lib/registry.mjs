import { fileURLToPath } from 'node:url';
import { AnnotationStore } from './annotations.mjs';
import { DeckCache } from './deck-cache.mjs';
import { ViewerError } from './errors.mjs';
import { DEFAULT_LANGUAGE, SUPPORTED_LANGUAGES, createI18n, normalizeLanguage } from './i18n.mjs';
import { createPresentationWithPowerPoint } from './new-deck.mjs';
import { renderWithPowerPoint } from './renderer.mjs';
import { PositionStore, SettingsStore, dataRootFor } from './store.mjs';
import { ViewerInstance } from './viewer-instance.mjs';

const ASSETS_DIR = fileURLToPath(new URL('../assets/', import.meta.url));

/** The open viewers of this extension process, one per canvas instance. They all share one language, kept in `i18n`. */
export class ViewerRegistry {
  constructor({ root = dataRootFor(), renderer = renderWithPowerPoint, creator = createPresentationWithPowerPoint, assetsDir = ASSETS_DIR, log = () => {}, watchIntervalMs, env = process.env } = {}) {
    this.assetsDir = assetsDir;
    this.log = log;
    this.creator = creator;
    this.watchIntervalMs = watchIntervalMs;
    this.env = env;
    this.i18n = createI18n({ language: normalizeLanguage(env.PPTX_CANVAS_LANG) ?? DEFAULT_LANGUAGE });
    this.settings = new SettingsStore(root, { log, i18n: this.i18n });
    this.deckCache = new DeckCache({ root, renderer, log, i18n: this.i18n });
    this.positions = new PositionStore(root, { log, i18n: this.i18n });
    this.annotationStore = new AnnotationStore(root, { log, i18n: this.i18n });
    this._instances = new Map();
    this._creating = new Map();
    this._ready = null;
  }

  /**
   * Settles the language: the one the user saved, else PPTX_CANVAS_LANG, else English. Safe to call repeatedly; `ensure` and
   * `setLanguage` wait for it, and extension.mjs awaits it before it registers the canvas (whose name is in that language).
   */
  ready() {
    this._ready ??= this.settings.load().then(({ language }) => {
      if (language) this.i18n.setLanguage(language);
    });
    return this._ready;
  }

  get language() {
    return this.i18n.language;
  }

  /** Saves and applies a language, and pushes the new texts to every open viewer. Returns the normalized language. */
  async setLanguage(value) {
    const language = normalizeLanguage(value);
    if (!language) throw new ViewerError('invalid_input', 'input.languageInvalid', { languages: SUPPORTED_LANGUAGES });
    await this.ready();
    this.i18n.setLanguage(language);
    const saved = this.settings.save({ language }).catch((error) => this.log(this.i18n.t('log.settingsSaveFailed', { detail: error.message })));
    for (const instance of this._instances.values()) instance.republish();
    await saved;
    return language;
  }

  /** The viewers that are open now. */
  instances() {
    return [...this._instances.values()];
  }

  /** Returns the live viewer for `instanceId`, starting one when there is none. Safe to call repeatedly. */
  async ensure(instanceId, { baseDir } = {}) {
    await this.ready();
    let instance = this._instances.get(instanceId);
    if (!instance) {
      let creating = this._creating.get(instanceId);
      if (!creating) {
        creating = this._create(instanceId, baseDir).finally(() => this._creating.delete(instanceId));
        this._creating.set(instanceId, creating);
      }
      instance = await creating;
    }
    if (baseDir) instance.baseDir = baseDir;
    return instance;
  }

  require(instanceId) {
    const instance = this._instances.get(instanceId);
    if (!instance) throw new ViewerError('instance_not_found', 'error.instanceNotFound');
    return instance;
  }

  async close(instanceId) {
    const instance = this._instances.get(instanceId);
    if (!instance) return;
    this._instances.delete(instanceId);
    await instance.close();
  }

  async closeAll() {
    await Promise.all([...this._instances.keys()].map((instanceId) => this.close(instanceId)));
  }

  async _create(instanceId, baseDir) {
    const instance = new ViewerInstance({
      instanceId,
      deckCache: this.deckCache,
      positions: this.positions,
      annotationStore: this.annotationStore,
      creator: this.creator,
      assetsDir: this.assetsDir,
      baseDir,
      i18n: this.i18n,
      changeLanguage: (language) => this.setLanguage(language),
      log: (message) => this.log(message),
      ...(this.watchIntervalMs === undefined ? {} : { watchIntervalMs: this.watchIntervalMs }),
    });
    try {
      await instance.start();
    } catch (error) {
      throw new ViewerError('server_start_failed', 'error.serverStartFailed', { detail: error.message });
    }
    this._instances.set(instanceId, instance);
    return instance;
  }
}
