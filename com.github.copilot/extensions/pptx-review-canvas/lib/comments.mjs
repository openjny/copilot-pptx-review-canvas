import { randomUUID } from 'node:crypto';
import { ViewerError } from './errors.mjs';

// Comments the user writes about a page: about the whole page (no anchor), about one point on it, or about an area. The position is
// kept as fractions of the page (0..1), like strokes, so it stays put at any size.

export const MAX_COMMENT_CHARS = 2000;
export const MAX_COMMENTS_PER_PAGE = 60;
const COMMENT_ID = /^c[0-9a-f]{12}$/;
// An area smaller than this (in fractions of the page) is a point.
const POINT_SIZE = 0.004;

const invalid = (messageKey, params) => new ViewerError('invalid_input', messageKey, params);
const round = (value) => Math.round(value * 1e4) / 1e4;
const clamp01 = (value) => Math.min(1, Math.max(0, value));

/** The text of a comment, trimmed and with plain line breaks. */
export function validateCommentText(input) {
  if (typeof input !== 'string') throw invalid('comment.textString');
  const text = input.replace(/\r\n?/g, '\n').trim();
  if (!text) throw invalid('comment.empty');
  if (text.length > MAX_COMMENT_CHARS) throw invalid('comment.tooLong', { max: MAX_COMMENT_CHARS });
  return text;
}

/** `{ x, y, w, h }` in fractions of the page, or null for a comment about the whole page. A point has no size (`w` and `h` are 0). */
export function validateAnchor(input) {
  if (input === null || input === undefined) return null;
  if (typeof input !== 'object' || Array.isArray(input)) throw invalid('comment.anchorNotObject');
  const { x, y, w = 0, h = 0 } = input;
  if (![x, y, w, h].every(Number.isFinite)) throw invalid('comment.anchorNotNumber');
  const left = clamp01(x);
  const top = clamp01(y);
  const width = Math.min(1 - left, Math.max(0, w));
  const height = Math.min(1 - top, Math.max(0, h));
  const point = width < POINT_SIZE && height < POINT_SIZE;
  return { x: round(left), y: round(top), w: point ? 0 : round(width), h: point ? 0 : round(height) };
}

/** Checks a comment read back from a record. */
export function validateComment(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw invalid('comment.notObject');
  if (typeof input.id !== 'string' || !COMMENT_ID.test(input.id)) throw invalid('comment.idInvalid');
  if (!Number.isFinite(input.createdAt) || !Number.isFinite(input.updatedAt)) throw invalid('comment.datesInvalid');
  return { id: input.id, text: validateCommentText(input.text), anchor: validateAnchor(input.anchor), createdAt: input.createdAt, updatedAt: input.updatedAt };
}

export function newComment({ text, anchor }) {
  const now = Date.now();
  return { id: `c${randomUUID().replace(/-/g, '').slice(0, 12)}`, text: validateCommentText(text), anchor: validateAnchor(anchor), createdAt: now, updatedAt: now };
}

/** 'page' (about the whole page), 'point' or 'area'. */
export function commentKind(comment) {
  if (!comment.anchor) return 'page';
  return comment.anchor.w === 0 && comment.anchor.h === 0 ? 'point' : 'area';
}
