// The "Submit review" button and the small panel it opens: the user's way of saying "I am done, now make the changes". The server
// sends the agent a message (`api/review/submit`); this module only shows when that is possible and collects the optional note.
//
// `elements`: `button` (with `buttonIcon` and `buttonLabel` inside it), `panel`, `summary`, `note`, `cancel`, `send`.
// The state it reads is `state.review`: `{ canSubmit, submitted, slides, drafts }`.

import { icon } from './icons.js';
import { t, tn } from './i18n.js';

export function createReview({ elements, getState, post, onError, announce, onOpen = () => {} }) {
  const { button, buttonIcon, buttonLabel, panel, summary, note, cancel, send } = elements;
  let open = false;
  let sending = false;

  const reviewOf = () => getState()?.review ?? null;

  function setIcon(name) {
    if (buttonIcon.dataset.icon === name) return;
    buttonIcon.dataset.icon = name;
    buttonIcon.innerHTML = icon(name);
  }

  function summaryText(review) {
    const parts = [];
    if (review.slides.length) parts.push(tn('review.summarySlides', review.slides.length));
    if (review.drafts) parts.push(tn('review.summaryDrafts', review.drafts));
    return parts.join(t('review.summaryJoin'));
  }

  /** Redraws the button and the panel from the state. Called after every render, and after the language changes. */
  function sync() {
    const review = reviewOf();
    const canSubmit = Boolean(review?.canSubmit);
    const submitted = Boolean(review?.submitted);
    button.disabled = !canSubmit;
    button.classList.toggle('is-done', submitted);
    setIcon(submitted ? 'check' : 'send');
    buttonLabel.textContent = submitted ? t('review.submitted') : t('review.submit');
    button.title = submitted ? t('review.submittedTitle') : canSubmit ? t('review.submitTitle') : t('review.nothingTitle');
    summary.textContent = review ? summaryText(review) : '';
    send.disabled = !canSubmit || sending;
    // The marks were read (or removed) while the panel was open, so there is nothing to submit any more.
    if (open && !canSubmit && !sending) setOpen(false);
  }

  function setOpen(next) {
    if (open === next) return;
    open = next;
    panel.hidden = !open;
    button.setAttribute('aria-expanded', String(open));
    if (open) {
      // The bar wraps in a narrow window, so the button is not always in the first row: put the panel right under it.
      panel.style.top = `${Math.round(button.getBoundingClientRect().bottom + 6)}px`;
      onOpen();
      note.focus();
    }
  }

  async function submit() {
    if (sending || !reviewOf()?.canSubmit) return;
    sending = true;
    send.disabled = true;
    send.textContent = t('review.sending');
    try {
      await post('api/review/submit', { note: note.value.trim() });
      note.value = '';
      setOpen(false);
      announce(t('review.done'));
    } catch (error) {
      onError(error);
    } finally {
      sending = false;
      send.textContent = t('review.send');
      sync();
    }
  }

  button.addEventListener('click', () => {
    if (!button.disabled) setOpen(!open);
  });
  cancel.addEventListener('click', () => {
    setOpen(false);
    button.focus();
  });
  send.addEventListener('click', submit);
  panel.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      setOpen(false);
      button.focus();
      event.preventDefault();
    } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && event.target === note) {
      submit();
      event.preventDefault();
    }
  });
  document.addEventListener('click', (event) => {
    if (!open || !(event.target instanceof Node)) return;
    if (!panel.contains(event.target) && !button.contains(event.target)) setOpen(false);
  });

  return { sync, close: () => setOpen(false), isOpen: () => open };
}
