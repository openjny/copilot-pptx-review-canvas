import { randomUUID } from 'node:crypto';

// The actions of one viewer that can be undone and redone: deleting a slide from the file, and deleting a new slide. The list is
// linear, like the undo of a text editor: an undone action waits to be redone until a new action is recorded, which drops it.

const DEFAULT_LIMIT = 10;

export class UndoJournal {
  /**
   * `release(entry)` is called once for every entry that leaves the journal for good, so it can free what the entry holds.
   * `describe(entry)` gives the label the page shows for an entry, in the language in use when `view` is called.
   */
  constructor({ limit = DEFAULT_LIMIT, release = () => {}, describe = (entry) => entry.label } = {}) {
    this._limit = limit;
    this._release = release;
    this._describe = describe;
    this._entries = [];
    this._cursor = 0;
  }

  /** Adds a new action (it can be undone now) and drops the actions that were waiting to be redone. */
  record(entry) {
    for (const dropped of this._entries.splice(this._cursor)) this._release(dropped);
    const added = { ...entry, id: randomUUID().slice(0, 8), at: Date.now() };
    this._entries.push(added);
    while (this._entries.length > this._limit) this._release(this._entries.shift());
    this._cursor = this._entries.length;
    return added;
  }

  /** The action `kind` ('undo' or 'redo') would act on, or null. */
  peek(kind) {
    return (kind === 'undo' ? this._entries[this._cursor - 1] : this._entries[this._cursor]) ?? null;
  }

  /** Moves past the action `kind` just acted on. */
  commit(kind) {
    const entry = this.peek(kind);
    if (!entry) return;
    this._cursor += kind === 'undo' ? -1 : 1;
    entry.at = Date.now();
  }

  clear() {
    const all = this._entries.splice(0);
    this._cursor = 0;
    for (const entry of all) this._release(entry);
  }

  get size() {
    return this._entries.length;
  }

  /** What the page needs: for each of undo and redo, the action it would act on. `at` orders it against the page's own undo of strokes. */
  view() {
    const pick = (entry) => entry && { id: entry.id, kind: entry.type, at: entry.at, label: this._describe(entry) };
    return { undo: pick(this.peek('undo')), redo: pick(this.peek('redo')) };
  }
}
