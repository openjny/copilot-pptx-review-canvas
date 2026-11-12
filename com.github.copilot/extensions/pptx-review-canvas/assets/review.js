// The "Submit review" button: the user's way of saying "I am done, now make the changes". Pressing it sends the review to the agent at once
// (`api/review/submit`), with nothing in between. This module only shows when that is possible and what the button says about it.
//
// `elements`: `button` (with `buttonIcon` and `buttonLabel` inside it).
// The state it reads is `state.review`: `{ canSubmit, submitted, slides, drafts }`.

import { icon } from './icons.js';
import { t, tn } from './i18n.js';

export function createReview({ elements, getState, post, onError, announce }) {
  const { button, buttonIcon, buttonLabel } = elements;
  let sending = false;

  const reviewOf = () => getState()?.review ?? null;

  function setIcon(name) {
    if (buttonIcon.dataset.icon === name) return;
    buttonIcon.dataset.icon = name;
    buttonIcon.innerHTML = icon(name);
  }

  /** What will be sent, in words: "2 slides with feedback · 1 new slide". It is in the tooltip, since there is no confirmation to read it in. */
  function summaryText(review) {
    const parts = [];
    if (review.slides.length) parts.push(tn('review.summarySlides', review.slides.length));
    if (review.drafts) parts.push(tn('review.summaryDrafts', review.drafts));
    return parts.join(t('review.summaryJoin'));
  }

  /** Redraws the button from the state. Called after every render, and after the language changes. */
  function sync() {
    const review = reviewOf();
    const canSubmit = Boolean(review?.canSubmit);
    const submitted = Boolean(review?.submitted);
    button.disabled = !canSubmit || sending;
    button.classList.toggle('is-done', submitted && !sending);
    setIcon(submitted && !sending ? 'check' : 'send');
    buttonLabel.textContent = sending ? t('review.sending') : submitted ? t('review.submitted') : t('review.submit');
    const summary = review ? summaryText(review) : '';
    button.title = submitted ? t('review.submittedTitle') : canSubmit ? `${t('review.submitTitle')}${summary ? `\n${summary}` : ''}` : t('review.nothingTitle');
  }

  async function submit() {
    if (sending || !reviewOf()?.canSubmit) return;
    sending = true;
    sync();
    try {
      await post('api/review/submit');
      announce(t('review.done'));
    } catch (error) {
      onError(error);
    } finally {
      sending = false;
      sync();
    }
  }

  button.addEventListener('click', submit);

  return { sync };
}
