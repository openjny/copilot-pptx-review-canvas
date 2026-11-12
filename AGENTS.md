# Agent notes

How to work in this repository, for agents and people. [README.md](./README.md) says what the project is, [CONTRIBUTING.md](./CONTRIBUTING.md) has the contributor guide and [docs/architecture.md](./docs/architecture.md) describes the modules; this file does not repeat them.

## Commands

- Test: `npm test` (Node 22 or later). It runs `node --test` with the file globs `test/*.test.mjs` and `com.github.copilot/extensions/pptx-review-canvas/test/*.test.mjs`; a bare directory does not work.
- Real PowerPoint check: `npm run smoke`. Unit tests never start PowerPoint.
- Sample deck: `uv run --with python-pptx python scripts/make_sample_deck.py`. Change the script, not `samples/q3-review.pptx`.

## Rules that are easy to break

- **Every visible string is in a catalog.** Server text: `lib/locales/en.mjs` and `ja.mjs`. Page text: `assets/locales/en.js` and `ja.js`. Add the English entry and the Japanese entry in the same change, with the same placeholders. Never put a Japanese literal in code; a test fails if you do. `assets/index.html` has English text as its default and `data-i18n*` attributes.
- **Errors carry keys.** Throw `ViewerError(code, messageKey, params)`; do not format a localized message where the language is unknown.
- **No inline styles** in the page (`style-src 'self'`): use classes or the CSSOM.
- **No dependencies and no build step.** Do not add a package or a bundler.
- **Existing data must keep loading.** Records are `{ format, version, data }`; changing a record shape needs a new version and a migration. Data written by older versions (Japanese draft titles, for example) is still read.
- **Short paths.** Keep every file path in the repository at 72 characters or fewer; see CONTRIBUTING.md for why.
- **Screenshots** in `docs/images/` and `assets/preview.png` show only the sample deck on a neutral path (no user names, no real decks). Retake them when the UI changes: start a viewer with the real renderer on a copy of the sample deck, and drive the page with a browser.

## Commits

Conventional Commits, for example `feat(i18n): add a language switch`. Work on a branch; do not commit to `main` directly.
