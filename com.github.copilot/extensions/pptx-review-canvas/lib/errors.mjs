import { tEn } from './i18n.mjs';

/**
 * An expected failure with a stable code; extension.mjs maps it to a CanvasError. `messageKey` and `params` name its text in the
 * catalog (see i18n.mjs); `message` is that text in English.
 */
export class ViewerError extends Error {
  constructor(code, messageKey, params) {
    super(tEn(messageKey, params));
    this.name = 'ViewerError';
    this.code = code;
    this.messageKey = messageKey;
    this.params = params;
  }
}

export function isAbortError(error) {
  return error instanceof ViewerError && error.code === 'cancelled';
}
