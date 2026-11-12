import { CanvasError, createCanvas, joinSession } from '@github/copilot-sdk/extension';
import { ViewerError } from './lib/errors.mjs';
import { ViewerRegistry } from './lib/registry.mjs';
import { CANVAS_ID, normalizePathKey } from './lib/store.mjs';

const RENDER_WAIT_MS = 3 * 60 * 1000;

let session;
const log = (message) => {
  session?.log(`[${CANVAS_ID}] ${message}`, { ephemeral: true }).catch(() => {});
};
const registry = new ViewerRegistry({ log });
// The language (the one the user chose in the viewer, or PPTX_CANVAS_LANG) has to be known before the canvas is registered under its name.
await registry.ready();
const t = (key, params) => registry.i18n.t(key, params);

function toCanvasError(error) {
  if (error instanceof CanvasError) return error;
  if (error instanceof ViewerError) return new CanvasError(error.code, registry.i18n.messageOf(error));
  return new CanvasError('internal_error', t('error.unexpected', { detail: error?.message ?? error }));
}

async function guarded(work) {
  try {
    return await work();
  } catch (error) {
    throw toCanvasError(error);
  }
}

function asObject(input) {
  if (input === undefined || input === null) return {};
  if (typeof input !== 'object' || Array.isArray(input)) throw new CanvasError('invalid_input', t('input.notObject'));
  return input;
}

function optionalPath(value) {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !value.trim()) throw new CanvasError('invalid_input', t('input.pathNonEmpty'));
  return value;
}

function optionalSlide(value) {
  if (value === undefined) return undefined;
  if (!Number.isInteger(value) || value < 1) throw new CanvasError('invalid_input', t('input.slidePositive'));
  return value;
}

function optionalBoolean(value, name) {
  if (value === undefined) return false;
  if (typeof value !== 'boolean') throw new CanvasError('invalid_input', t('input.notBoolean', { name }));
  return value;
}

function requireSlide(value) {
  const slide = optionalSlide(value);
  if (slide === undefined) throw new CanvasError('invalid_input', t('input.slideRequired'));
  return slide;
}

function optionalDraft(value) {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !/^d[0-9a-f]{12}$/.test(value)) throw new CanvasError('invalid_input', t('input.draftId'));
  return value;
}

function describeImage(image) {
  if (image.draft) return `New slide idea "${image.title}" (insert after slide ${image.after}) with the user's sketch on a white page; black numbers match marks[].number, blue numbers (and blue rectangles) match comments[].number`;
  return image.kind === 'clean' ? `Slide ${image.slide} without marks` : `Slide ${image.slide} with the user's marks; black numbers match marks[].number, blue numbers (and blue rectangles) match comments[].number`;
}

/** Waits for the deck to be ready; a failed render becomes an error for the caller. */
async function settled(instance) {
  const result = await instance.whenSettled(RENDER_WAIT_MS);
  if (result.phase === 'error' && result.error) throw new CanvasError(result.error.code, result.error.message);
  return result;
}

/** The viewer for an action that reads the deck, once any render that is still running has finished. */
const readyViewer = (ctx) => registry.require(ctx.instanceId).whenReady(RENDER_WAIT_MS);

const pathProperty = { type: 'string', description: 'Absolute path of a PowerPoint file (.pptx). A relative path resolves against the session working directory.' };
const slideProperty = { type: 'integer', minimum: 1, description: 'Slide number, counted from 1.' };

function boundedInteger(value, name, min, max, fallback) {
  if (value === undefined) return fallback;
  if (!Number.isInteger(value) || value < min || value > max) throw new CanvasError('invalid_input', t('input.integerRange', { name, min, max }));
  return value;
}

const MARK_GUIDE = [
  'Each mark is a group of pen or highlighter strokes that touch each other. marks[].number is the black number printed next to that mark in the image.',
  'comments[] are the notes the user wrote about a page. kind "page" is about the whole page (no position); "point" and "area" are tied to a place: boxPercent is where (x, y, and w, h for an area, in percent of the slide), and the image shows a blue number at it (and a blue rectangle for an area) that matches comments[].number. A comment is what the user wants, in their own words, so follow it like their chat message; if it contradicts their chat message or is unclear, ask before editing.',
  'onShapes lists the shapes under the mark or comment, most relevant first; their names are the paths of the deck build tree, so use them to find the source of the element. A shape with near: true is next to the mark rather than under it.',
  'The marks only say where the user pointed. What to change comes from the user\'s message and comments; if they disagree or are unclear, ask before editing.',
  'draftPages are ideas for NEW slides that exist only in the viewer, not in the file. insertAfterSlide says where each goes (0 = before the first slide); the strokes are a rough sketch of its layout or diagram, and an empty draft page means "put a slide here" with the content in the user\'s message or comments. The numbers in the draft titles are internal: the screen does not show them, so when you talk to the user, name a draft by where it goes (for example "the new slide after slide 3") and never by its number.',
  'Build each draft page as a real slide at that position, then call resolve_drafts with the ids you used so they leave the viewer.',
  'When kind is "sketchbook" there is no .pptx yet: the pages are sketches only. Create the deck from them, save it, open it with open_deck, and resolve the drafts you used.',
  'After you rebuild the deck, the viewer reloads it by itself and keeps the marks and comments on slides whose picture did not change, so the user can review again.',
].join(' ');

/** Hidden context for the agent when the user has drawn on pages it has not looked at yet, or added new slide ideas. */
function unreadMarksContext() {
  const byDeck = new Map();
  for (const instance of registry.instances()) {
    const unread = instance.unreadAnnotations();
    if (unread && !byDeck.has(normalizePathKey(unread.path))) byDeck.set(normalizePathKey(unread.path), { unread, instance });
  }
  if (!byDeck.size) return null;
  const lines = [...byDeck.values()].map(({ unread, instance }) => {
    const parts = [];
    if (unread.slides.length) parts.push(t('unread.slides', { list: unread.slides }));
    if (unread.drafts.length) parts.push(t('unread.drafts', { count: unread.drafts.length }));
    return t('unread.deckLine', { deck: unread.deck, parts: parts.join(' / '), instanceId: instance.instanceId });
  });
  return [t('unread.intro', { canvasId: CANVAS_ID }), ...lines, t('unread.outro')].join('\n');
}

const pptxSlideViewer = createCanvas({
  id: CANVAS_ID,
  displayName: t('app.title'),
  description: 'Shows a PowerPoint (.pptx) deck as slide images with speaker notes, lets the user draw on slides and on new-slide drafts (or on a blank sketchbook with no file), reorder slides by drag and drop, create a new .pptx, and lets the agent read those marks (get_annotations), open decks and move between slides.',
  inputSchema: {
    type: 'object',
    properties: {
      path: pathProperty,
      slide: { ...slideProperty, description: 'Slide to show first, counted from 1. Defaults to the slide the deck was last left on.' },
      sketchbook: { type: 'string', description: 'Name of a sketchbook to start or continue: blank pages to sketch ideas on, with no PowerPoint file. Use instead of path.' },
    },
    additionalProperties: false,
  },
  actions: [
    {
      name: 'open_deck',
      description: 'Open a .pptx file in this viewer and wait until its slides are rendered. Returns the slide count.',
      inputSchema: {
        type: 'object',
        properties: { path: pathProperty, slide: slideProperty },
        required: ['path'],
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const input = asObject(ctx.input);
        const path = optionalPath(input.path);
        if (!path) throw new CanvasError('invalid_input', t('input.pathRequired'));
        const instance = registry.require(ctx.instanceId);
        if (ctx.session?.workingDirectory) instance.baseDir = ctx.session.workingDirectory;
        await instance.openDeck(path, { slide: optionalSlide(input.slide) });
        return settled(instance);
      }),
    },
    {
      name: 'go_to_slide',
      description: 'Show the given slide in the viewer. Returns its title and speaker notes.',
      inputSchema: {
        type: 'object',
        properties: { slide: slideProperty },
        required: ['slide'],
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const slide = requireSlide(asObject(ctx.input).slide);
        return (await readyViewer(ctx)).setCurrent(slide);
      }),
    },
    {
      name: 'get_slide',
      description: 'Read the title and speaker notes of a slide (the one on screen by default), optionally with its image.',
      inputSchema: {
        type: 'object',
        properties: {
          slide: slideProperty,
          includeImage: { type: 'boolean', description: 'Also return the rendered slide image.' },
        },
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const input = asObject(ctx.input);
        const includeImage = optionalBoolean(input.includeImage, 'includeImage');
        const requested = optionalSlide(input.slide);
        const instance = await readyViewer(ctx);
        const info = instance.slideInfo(requested);
        if (!includeImage) return info;
        const png = await instance.slideImage(info.slide);
        return {
          textResultForLlm: JSON.stringify(info),
          binaryResultsForLlm: [{ type: 'image', mimeType: 'image/png', data: png.toString('base64'), description: `Slide ${info.slide}: ${info.title}` }],
          resultType: 'success',
        };
      }),
    },
    {
      name: 'get_outline',
      description: 'List the titles of all slides in the open deck, and which slide is on screen.',
      inputSchema: {
        type: 'object',
        properties: { includeNotes: { type: 'boolean', description: 'Also return the speaker notes of every slide.' } },
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const includeNotes = optionalBoolean(asObject(ctx.input).includeNotes, 'includeNotes');
        return (await readyViewer(ctx)).outline({ includeNotes });
      }),
    },
    {
      name: 'get_annotations',
      description: 'Read what the user drew and wrote: pen marks and comments on the slides and on new-slide drafts (ideas for slides that are not in the file yet); comments are about a whole page, a point or an area. Each page with marks or a point/area comment comes as an image with a number next to every mark (black) and comment (blue). '
        + 'Call this first when the user gives feedback on the deck, refers to their marks or comments, or asks for new slides; then edit or build the deck as they instruct. Reading it clears the "unread" state.',
      inputSchema: {
        type: 'object',
        properties: {
          slide: { ...slideProperty, description: 'Only this slide (no drafts). Defaults to every slide that has marks or comments, plus all draft pages.' },
          draft: { type: 'string', pattern: '^d[0-9a-f]{12}$', description: 'Only this draft page (an id from draftPages).' },
          includeImages: { type: 'boolean', description: 'Return the annotated page images (default true).' },
          includeClean: { type: 'boolean', description: 'Also return each slide without the marks, to see what is underneath (default false).' },
          maxImages: { type: 'integer', minimum: 1, maximum: 20, description: 'Upper limit of annotated images (default 8).' },
        },
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const input = asObject(ctx.input);
        const instance = await readyViewer(ctx);
        const includeImages = input.includeImages === undefined ? true : optionalBoolean(input.includeImages, 'includeImages');
        const { info, images } = await instance.annotationReport({
          slide: optionalSlide(input.slide),
          draft: optionalDraft(input.draft),
          includeImages,
          includeClean: optionalBoolean(input.includeClean, 'includeClean'),
          maxImages: boundedInteger(input.maxImages, 'maxImages', 1, 20, 8),
        });
        if (info.annotatedSlides === 0 && info.draftPages.length === 0) return { ...info, note: t('report.nothing') };
        return {
          textResultForLlm: JSON.stringify({ ...info, guide: MARK_GUIDE }),
          binaryResultsForLlm: images.map((image) => ({
            type: 'image',
            mimeType: 'image/png',
            data: image.data.toString('base64'),
            description: describeImage(image),
          })),
          resultType: 'success',
        };
      }),
    },
    {
      name: 'resolve_drafts',
      description: 'Tell the viewer that you turned these new-slide drafts into real slides: they leave the viewer (and are set aside in the extension data, not deleted). Pass the ids from get_annotations draftPages.',
      inputSchema: {
        type: 'object',
        properties: { ids: { type: 'array', minItems: 1, maxItems: 40, items: { type: 'string', pattern: '^d[0-9a-f]{12}$' }, description: 'Draft page ids.' } },
        required: ['ids'],
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const { ids } = asObject(ctx.input);
        if (!Array.isArray(ids) || ids.length === 0 || ids.some((id) => typeof id !== 'string')) throw new CanvasError('invalid_input', t('input.idsArray'));
        return (await readyViewer(ctx)).resolveDrafts(ids);
      }),
    },
    {
      name: 'new_sketchbook',
      description: 'Start a sketchbook in this viewer (or continue the one with the same name): blank pages where the user sketches ideas for a deck that does not exist yet. Read the sketches with get_annotations, then build the deck.',
      inputSchema: {
        type: 'object',
        properties: { name: { type: 'string', maxLength: 80, description: 'Name of the sketchbook. Defaults to one made from the current time.' } },
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const name = asObject(ctx.input).name;
        if (name !== undefined && typeof name !== 'string') throw new CanvasError('invalid_input', t('input.nameString'));
        const instance = registry.require(ctx.instanceId);
        await instance.newSketchbook(name);
        const result = await settled(instance);
        return { ...result, sketchbookPath: instance.snapshot().deck?.path };
      }),
    },
    {
      name: 'move_slide',
      description: 'Move a slide to another position in the .pptx file itself (only the slide order is rewritten; the previous file is kept as a backup in the extension data). '
        + 'Marks and draft pages go with their slides, and the viewer shows the result at once. Not available for a sketchbook or a legacy .ppt file, or while another app has the file open.',
      inputSchema: {
        type: 'object',
        properties: {
          from: { ...slideProperty, description: 'Current number of the slide to move, counted from 1.' },
          to: { ...slideProperty, description: 'Number the slide should have afterwards, counted from 1.' },
        },
        required: ['from', 'to'],
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const input = asObject(ctx.input);
        const instance = await readyViewer(ctx);
        const result = await instance.moveSlide(requireSlide(input.from), requireSlide(input.to));
        return { ...result, outline: instance.outline() };
      }),
    },
    {
      name: 'new_deck',
      description: 'Create a new .pptx with one slide through PowerPoint and open it in this viewer. Without template it starts from PowerPoint\'s blank presentation; with a template (.potx/.pptx) it starts from a copy of it. '
        + 'Never overwrites: it fails when the file exists. For anything beyond one empty slide, build the deck yourself and open it with open_deck.',
      inputSchema: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Where to save it (.pptx; the extension is added when missing). A relative path resolves against the session working directory. The folder must exist.' },
          template: { type: 'string', description: 'Optional PowerPoint template or deck to start from (.potx, .pptx, ...). It is not changed.' },
        },
        required: ['path'],
        additionalProperties: false,
      },
      handler: (ctx) => guarded(async () => {
        const input = asObject(ctx.input);
        const path = optionalPath(input.path);
        if (!path) throw new CanvasError('invalid_input', t('input.pathRequired'));
        if (input.template !== undefined && (typeof input.template !== 'string' || !input.template.trim())) throw new CanvasError('invalid_input', t('input.templateNonEmpty'));
        const instance = registry.require(ctx.instanceId);
        if (ctx.session?.workingDirectory) instance.baseDir = ctx.session.workingDirectory;
        await instance.newDeck(path, { template: input.template });
        const result = await settled(instance);
        return { ...result, path: instance.snapshot().deck?.path };
      }),
    },
    {
      name: 'refresh',
      description: 'Render the open deck again from the file on disk, for example after the file was rebuilt, and wait until it is ready.',
      handler: (ctx) => guarded(async () => {
        const instance = registry.require(ctx.instanceId);
        await instance.refresh();
        return settled(instance);
      }),
    },
  ],
  open: (ctx) => guarded(async () => {
    const input = asObject(ctx.input);
    const path = optionalPath(input.path);
    const slide = optionalSlide(input.slide);
    if (input.sketchbook !== undefined && typeof input.sketchbook !== 'string') throw new CanvasError('invalid_input', t('input.sketchbookString'));
    if (path && input.sketchbook !== undefined) throw new CanvasError('invalid_input', t('input.pathAndSketchbook'));
    const instance = await registry.ensure(ctx.instanceId, { baseDir: ctx.session?.workingDirectory });
    if (path) await instance.openDeck(path, { slide });
    else if (input.sketchbook !== undefined) await instance.newSketchbook(input.sketchbook);
    else if (slide && instance.hasDeck) instance.setCurrent(Math.min(slide, instance.summary().slideCount));
    return { url: instance.url, title: instance.title, status: instance.statusText };
  }),
  onClose: (ctx) => registry.close(ctx.instanceId).catch((error) => log(t('log.closeFailed', { detail: error.message }))),
});

session = await joinSession({
  canvases: [pptxSlideViewer],
  hooks: {
    onUserPromptSubmitted: async () => {
      try {
        const additionalContext = unreadMarksContext();
        return additionalContext ? { additionalContext } : undefined;
      } catch (error) {
        log(t('log.unreadCheckFailed', { detail: error.message }));
        return undefined;
      }
    },
  },
});
log(t('log.registered'));

process.on('unhandledRejection', (reason) => log(t('log.unhandledRejection', { detail: reason?.stack ?? reason })));
for (const signal of ['SIGTERM', 'SIGINT']) {
  process.once(signal, () => {
    registry.closeAll().finally(() => process.exit(0));
  });
}
