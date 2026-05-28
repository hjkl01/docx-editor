/**
 * Suggestion Mode Plugin
 *
 * When active, intercepts all text insertions and deletions,
 * wrapping them in tracked change marks (insertion/deletion)
 * instead of modifying the document directly.
 *
 * - Typed text is marked as insertion (green underline)
 * - Deleted text is NOT removed — it's marked as deletion (red strikethrough)
 * - Text already marked as insertion by the current author is deleted normally
 *   (retracting your own suggestion)
 */

import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Node as PMNode, MarkType } from 'prosemirror-model';

import { mintRevisionId } from './revisionIds';
import { applyPostSplitInheritance } from '../extensions/features/BaseKeymapExtension';

const STYLE_MARK_NAMES = new Set(['fontFamily', 'fontSize', 'textColor']);

export const suggestionModeKey = new PluginKey<SuggestionModeState>('suggestionMode');
const SUGGESTION_META = 'suggestionModeApplied';
/** Set by accept/reject commands to bypass suggesting-mode interception. */
export const SUGGESTION_BYPASS_META = 'suggestionModeBypass';

interface SuggestionModeState {
  active: boolean;
  author: string;
}

interface MarkAttrs {
  revisionId: number;
  author: string;
  date: string;
}

function makeMarkAttrs(pluginState: SuggestionModeState): MarkAttrs {
  return {
    revisionId: mintRevisionId(),
    author: pluginState.author,
    date: new Date().toISOString(),
  };
}

/** Reserve a revision triple without applying it — used by Enter/Backspace handlers. */
function makeRevisionInfo(pluginState: SuggestionModeState): MarkAttrs {
  return makeMarkAttrs(pluginState);
}

/**
 * Find an adjacent mark of the same type by the same author.
 * Reuses its revisionId so consecutive edits group into one change.
 */
function findAdjacentRevision(
  doc: PMNode,
  pos: number,
  markTypeName: string,
  author: string
): MarkAttrs | null {
  try {
    const $pos = doc.resolve(pos);
    for (const node of [$pos.nodeBefore, $pos.nodeAfter]) {
      if (node?.isText) {
        const mark = node.marks.find(
          (m) => m.type.name === markTypeName && m.attrs.author === author
        );
        if (mark) return mark.attrs as MarkAttrs;
      }
    }
  } catch {
    /* position out of range */
  }
  return null;
}

/**
 * Find an adjacent revision at either edge of a range.
 * This keeps consecutive backspaces grouped even though the cursor moves left.
 */
function findAdjacentRevisionForRange(
  doc: PMNode,
  from: number,
  to: number,
  markTypeName: string,
  author: string
): MarkAttrs | null {
  return (
    findAdjacentRevision(doc, from, markTypeName, author) ??
    findAdjacentRevision(doc, to, markTypeName, author)
  );
}

/**
 * Find a `pPrIns`/`pPrDel` revision on a paragraph adjacent to `paraStart`
 * carried by the same author. Used to coalesce consecutive Enter / Backspace
 * presses into one tracked change so the sidebar shows a single card and
 * a single Accept resolves the whole run (matches Word's grouping).
 *
 * `attr` selects which paragraph-mark attr to look for; the lookup checks
 * BOTH the previous and next paragraph since a new pPrIns may sit either
 * side of an existing run depending on cursor position.
 */
function findAdjacentParagraphMark(
  doc: PMNode,
  paraStart: number,
  attr: 'pPrIns' | 'pPrDel',
  author: string
): MarkAttrs | null {
  try {
    const $pos = doc.resolve(paraStart);
    const candidates: Array<PMNode | null | undefined> = [$pos.nodeBefore, $pos.nodeAfter];
    for (const node of candidates) {
      if (node?.type.name !== 'paragraph') continue;
      const existing = node.attrs[attr] as MarkAttrs | null;
      if (existing && existing.author === author) return existing;
    }
  } catch {
    /* paragraph at start/end of doc */
  }
  return null;
}

/**
 * Walk a text range and either mark as deletion or retract own insertions.
 * Processes in reverse order to maintain position validity.
 */
function markRangeAsDeleted(
  tr: Transaction,
  doc: PMNode,
  from: number,
  to: number,
  insertionType: MarkType,
  deletionType: MarkType,
  pluginState: SuggestionModeState
): void {
  const ranges: { from: number; to: number; isOwnInsert: boolean }[] = [];

  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return;
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start >= end) return;
    const isOwnInsert = node.marks.some(
      (m) => m.type === insertionType && m.attrs.author === pluginState.author
    );
    ranges.push({ from: start, to: end, isOwnInsert });
  });

  if (ranges.length === 0) return;

  const delAttrs =
    findAdjacentRevisionForRange(doc, from, to, 'deletion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  for (let i = ranges.length - 1; i >= 0; i--) {
    const range = ranges[i];
    if (range.isOwnInsert) {
      tr.delete(range.from, range.to);
    } else {
      tr.addMark(range.from, range.to, deletionType.create(delAttrs));
    }
  }
}

/**
 * Insert text as a tracked insertion, optionally marking replaced selection as deletion.
 */
function applySuggestionInsert(
  view: EditorView,
  from: number,
  to: number,
  text: string,
  pluginState: SuggestionModeState
): boolean {
  const insertionType = view.state.schema.marks.insertion;
  if (!insertionType) return false;

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, 'insertion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  if (from !== to) {
    const deletionType = view.state.schema.marks.deletion;
    if (deletionType) {
      markRangeAsDeleted(tr, view.state.doc, from, to, insertionType, deletionType, pluginState);
    }
  }

  const insertAt = tr.mapping.map(to);
  tr.insertText(text, insertAt, insertAt);

  // Strip inherited deletion marks — new text must never be marked as deleted.
  const deletionType = view.state.schema.marks.deletion;
  if (deletionType) {
    tr.removeMark(insertAt, insertAt + text.length, deletionType);
  }

  // Apply the correct insertion mark. If the cursor was inside an existing
  // insertion by the same author, insertText already inherited that mark and
  // insertAttrs will match — addMark is effectively a no-op that preserves
  // the continuous mark span. We intentionally do NOT removeMark(insertionType)
  // first, because that fragments the mark span and creates a nested change.
  tr.addMark(insertAt, insertAt + text.length, insertionType.create(insertAttrs));

  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Handle delete (forward or backward) in suggestion mode.
 */
function handleSuggestionDelete(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  direction: 'backward' | 'forward'
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) return false;

  const { $from, $to, empty } = state.selection;
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType || !deletionType) return false;

  if (!dispatch) return true;

  const tr = state.tr;
  tr.setMeta(SUGGESTION_META, true);

  // --- Selection delete ---
  if (!empty) {
    markRangeAsDeleted(tr, state.doc, $from.pos, $to.pos, insertionType, deletionType, pluginState);
    // Collapse cursor to after the marked/retracted content
    const cursorPos = tr.mapping.map($to.pos);
    tr.setSelection(TextSelection.near(tr.doc.resolve(cursorPos)));
    dispatch(tr.scrollIntoView());
    return true;
  }

  // --- Single character delete ---
  const isBackward = direction === 'backward';
  const deletePos = isBackward ? $from.pos - 1 : $from.pos;
  const deleteEnd = isBackward ? $from.pos : $from.pos + 1;

  if (deletePos < 0 || deleteEnd > state.doc.content.size) return true;

  const $deletePos = state.doc.resolve(deletePos);
  const nodeAfter = $deletePos.nodeAfter;

  // At block boundary — let default behavior handle (e.g. join paragraphs)
  if (!nodeAfter?.isText) return false;

  const hasOwnInsertion = nodeAfter.marks.some(
    (m) => m.type === insertionType && m.attrs.author === pluginState.author
  );
  const hasDeletion = nodeAfter.marks.some((m) => m.type === deletionType);

  if (hasDeletion) {
    // Already deleted — skip cursor past it
    const newPos = isBackward ? deletePos : deleteEnd;
    tr.setSelection(TextSelection.near(tr.doc.resolve(newPos)));
  } else if (hasOwnInsertion) {
    // Retract own insertion — actually delete the character
    tr.delete(deletePos, deleteEnd);
  } else {
    // Mark as deletion instead of removing
    const delAttrs =
      findAdjacentRevisionForRange(
        state.doc,
        deletePos,
        deleteEnd,
        'deletion',
        pluginState.author
      ) || makeMarkAttrs(pluginState);
    tr.addMark(deletePos, deleteEnd, deletionType.create(delAttrs));
    // Move cursor past the deletion mark
    const newPos = isBackward ? deletePos : deleteEnd;
    tr.setSelection(TextSelection.near(tr.doc.resolve(newPos)));
  }

  dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Suggesting-mode Enter handler. Splits the paragraph (via the existing
 * BaseKeymapExtension `splitBlockClearBorders` behavior, re-implemented
 * inline so we can capture the resulting transaction and add a `pPrIns`
 * attr on the *first* paragraph in the same PM transaction).
 *
 * Per ECMA-376 §17.13.5, the paragraph mark of the FIRST paragraph after
 * a split is the one that was newly introduced. Reject of `pPrIns` joins
 * the first paragraph back with the next; accept just clears the marker.
 */
function handleSuggestionEnter(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) return false;

  // Selection must be inside a paragraph (other block types fall through).
  const { $from, $to } = state.selection;
  if ($from.parent.type.name !== 'paragraph') return false;
  if ($to.parent.type.name !== 'paragraph') return false;

  if (!dispatch) return true;

  // If the selection covers content, mark it as deletion first (existing
  // suggesting-mode behavior) so the split happens at the selection start.
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType || !deletionType) return false;

  // Capture source paragraph + active style marks BEFORE any tr work so
  // `applyPostSplitInheritance` can match `splitBlockClearBorders` behavior:
  // typed text after the split inherits font / size / color via setStoredMarks.
  const sourcePara = $from.parent;
  const preMarks = state.storedMarks ?? $from.marks();
  const styleMarks = preMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));

  const tr = state.tr;
  tr.setMeta(SUGGESTION_META, true);

  if (!state.selection.empty) {
    markRangeAsDeleted(tr, state.doc, $from.pos, $to.pos, insertionType, deletionType, pluginState);
    // Collapse cursor to the deletion start before splitting.
    const collapsePos = tr.mapping.map($from.pos);
    tr.setSelection(TextSelection.near(tr.doc.resolve(collapsePos)));
  }

  // The first paragraph is the one whose mark just got introduced. We need
  // its absolute position BEFORE the split to find it again after.
  const $cursor = tr.selection.$from;
  const firstParaStart = $cursor.before($cursor.depth);

  // Split the paragraph at the cursor. After tr.split, the cursor (mapped)
  // lands at the start of the NEW paragraph, which is what
  // applyPostSplitInheritance expects.
  tr.split(tr.selection.from, 1);

  // Set pPrIns on the FIRST paragraph (the one before the split). Coalesce
  // with an adjacent same-author pPrIns so consecutive Enters in one editing
  // session show as a single tracked change in the sidebar.
  const firstPara = tr.doc.nodeAt(firstParaStart);
  if (firstPara && firstPara.type.name === 'paragraph') {
    const info =
      findAdjacentParagraphMark(tr.doc, firstParaStart, 'pPrIns', pluginState.author) ??
      makeRevisionInfo(pluginState);
    tr.setNodeMarkup(firstParaStart, undefined, {
      ...firstPara.attrs,
      pPrIns: info,
    });
  }

  // Shared with plain Enter: inherits style attrs, clears borders, and
  // (for an empty new paragraph) sets stored marks so typed text picks up
  // the source paragraph's font / size / color.
  applyPostSplitInheritance(tr, sourcePara, styleMarks, state.schema);

  dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Suggesting-mode Backspace at the start of a non-first paragraph: set
 * `pPrDel` on the PREVIOUS paragraph (its terminating mark is the one
 * being eaten). Caret lands at the end of the previous paragraph.
 *
 * Returns false at the very start of the document (nothing to mark), so
 * the base keymap can chain through (which itself is a no-op there).
 */
function handleSuggestionBackspaceAtStart(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) return false;
  const { $from, empty } = state.selection;
  if (!empty) return false;
  if ($from.parentOffset !== 0) return false;
  if ($from.parent.type.name !== 'paragraph') return false;

  const paraStart = $from.before($from.depth);
  if (paraStart <= 0) return false; // first paragraph in the document
  // `paraStart` is the position immediately before the current paragraph's
  // open tag. `nodeBefore` at that position returns the previous sibling.
  const prevPara = state.doc.resolve(paraStart).nodeBefore;
  if (!prevPara || prevPara.type.name !== 'paragraph') return false;
  // Already marked as deleted by the same author — second Backspace is a no-op.
  if (prevPara.attrs.pPrDel) {
    if (dispatch) {
      const prevParaEnd = paraStart - 1;
      const tr = state.tr.setSelection(TextSelection.near(state.doc.resolve(prevParaEnd)));
      dispatch(tr);
    }
    return true;
  }

  if (!dispatch) return true;

  const prevParaStart = paraStart - prevPara.nodeSize;
  // Coalesce with adjacent same-author pPrDel so a run of Backspaces shows
  // as one tracked change.
  const info =
    findAdjacentParagraphMark(state.doc, prevParaStart, 'pPrDel', pluginState.author) ??
    makeRevisionInfo(pluginState);
  const tr = state.tr.setNodeMarkup(prevParaStart, undefined, {
    ...prevPara.attrs,
    pPrDel: info,
  });
  tr.setMeta(SUGGESTION_META, true);
  // Caret to end of previous paragraph.
  const prevParaEnd = paraStart - 1;
  tr.setSelection(TextSelection.near(tr.doc.resolve(prevParaEnd)));
  dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Suggesting-mode Delete at end of a non-last paragraph: set `pPrDel` on
 * the CURRENT paragraph.
 */
function handleSuggestionDeleteAtEnd(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined
): boolean {
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState?.active) return false;
  const { $from, empty } = state.selection;
  if (!empty) return false;
  if ($from.parent.type.name !== 'paragraph') return false;
  if ($from.parentOffset !== $from.parent.content.size) return false;

  const para = $from.parent;
  const paraStart = $from.before($from.depth);
  const paraEnd = paraStart + para.nodeSize;
  if (paraEnd >= state.doc.content.size) return false; // last paragraph
  const $afterPara = state.doc.resolve(paraEnd);
  const nextPara = $afterPara.nodeAfter;
  if (!nextPara || nextPara.type.name !== 'paragraph') return false;
  if (para.attrs.pPrDel) {
    return true; // already marked
  }

  if (!dispatch) return true;

  // Coalesce with adjacent same-author pPrDel.
  const info =
    findAdjacentParagraphMark(state.doc, paraStart, 'pPrDel', pluginState.author) ??
    makeRevisionInfo(pluginState);
  const tr = state.tr.setNodeMarkup(paraStart, undefined, {
    ...para.attrs,
    pPrDel: info,
  });
  tr.setMeta(SUGGESTION_META, true);
  dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Create the suggestion mode plugin.
 * When active, text edits become tracked changes.
 */
export function createSuggestionModePlugin(initialActive = false, author = 'User'): Plugin {
  return new Plugin({
    key: suggestionModeKey,

    state: {
      init(): SuggestionModeState {
        return { active: initialActive, author };
      },
      apply(tr, state): SuggestionModeState {
        const meta = tr.getMeta(suggestionModeKey);
        if (meta) {
          return { ...state, ...meta };
        }
        return state;
      },
    },

    props: {
      handleDOMEvents: {
        // Intercept text input at the DOM level. ProseMirror's handleTextInput
        // is NOT reliably called when the hidden PM has complex mark structures
        // (it requires the change to span exactly one text node). By handling
        // beforeinput directly, we ensure suggestion mode always processes input.
        beforeinput(view: EditorView, event: InputEvent) {
          const pluginState = suggestionModeKey.getState(view.state);
          if (!pluginState?.active) return false;

          if (event.inputType === 'insertText' && event.data) {
            event.preventDefault();
            const { from, to } = view.state.selection;
            return applySuggestionInsert(view, from, to, event.data, pluginState);
          }

          return false;
        },
      },
      // Intercept Backspace and Delete to mark as deletion.
      // Enter splits the paragraph and marks the FIRST paragraph's pPrIns.
      handleKeyDown(view: EditorView, event: KeyboardEvent): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) return false;

        if (event.key === 'Enter' && !event.shiftKey) {
          return handleSuggestionEnter(view.state, view.dispatch);
        }
        if (event.key === 'Backspace') {
          // At paragraph start (non-first paragraph), track the pilcrow
          // deletion instead of joining or deleting a character.
          if (handleSuggestionBackspaceAtStart(view.state, view.dispatch)) return true;
          return handleSuggestionDelete(view.state, view.dispatch, 'backward');
        }
        if (event.key === 'Delete') {
          if (handleSuggestionDeleteAtEnd(view.state, view.dispatch)) return true;
          return handleSuggestionDelete(view.state, view.dispatch, 'forward');
        }
        return false;
      },

      // Backup: also handle via PM's handleTextInput for simple cases
      handleTextInput(view: EditorView, from: number, to: number, text: string): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) return false;
        return applySuggestionInsert(view, from, to, text, pluginState);
      },
    },

    // Catch-all: mark any unhandled new content (e.g. paste) as insertion
    appendTransaction(transactions, _oldState, newState) {
      const pluginState = suggestionModeKey.getState(newState);
      if (!pluginState?.active) return null;

      // Skip the catch-all mark-as-insertion path for both:
      //   - transactions we've already authored (`SUGGESTION_META`)
      //   - accept/reject command transactions (`SUGGESTION_BYPASS_META`)
      // The bypass meta is set by `resolveById` so structural-revision joins
      // (e.g. `pPrIns` reject → `tr.split` + `tr.setNodeMarkup`) aren't
      // re-wrapped as user insertions.
      const userTr = transactions.find(
        (tr) => tr.docChanged && !tr.getMeta(SUGGESTION_META) && !tr.getMeta(SUGGESTION_BYPASS_META)
      );
      if (!userTr) return null;

      const insertionType = newState.schema.marks.insertion;
      if (!insertionType) return null;

      const markAttrs = makeMarkAttrs(pluginState);

      const tr = newState.tr;
      tr.setMeta(SUGGESTION_META, true);

      const deletionType = newState.schema.marks.deletion;
      userTr.steps.forEach((step) => {
        const stepMap = step.getMap();
        stepMap.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
          if (newTo > newFrom) {
            // Only mark text nodes that don't already have tracked change marks.
            // Marking the entire range would overwrite existing marks from other authors.
            newState.doc.nodesBetween(newFrom, newTo, (node, pos) => {
              if (!node.isText) return;
              const hasTrackedMark = node.marks.some(
                (m) => m.type === insertionType || (deletionType && m.type === deletionType)
              );
              if (!hasTrackedMark) {
                const nodeStart = Math.max(pos, newFrom);
                const nodeEnd = Math.min(pos + node.nodeSize, newTo);
                tr.addMark(nodeStart, nodeEnd, insertionType.create(markAttrs));
              }
            });
          }
        });
      });

      return tr.steps.length > 0 ? tr : null;
    },
  });
}

/**
 * Toggle suggestion mode on/off.
 */
export function toggleSuggestionMode(
  state: EditorState,
  dispatch?: (tr: Transaction) => void
): boolean {
  const current = suggestionModeKey.getState(state);
  if (!current) return false;

  if (dispatch) {
    const tr = state.tr.setMeta(suggestionModeKey, {
      active: !current.active,
    });
    dispatch(tr);
  }
  return true;
}

/**
 * Set suggestion mode active state and author.
 */
export function setSuggestionMode(
  active: boolean,
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  author?: string
): boolean {
  if (dispatch) {
    const meta: Partial<SuggestionModeState> = { active };
    if (author !== undefined) meta.author = author;
    const tr = state.tr.setMeta(suggestionModeKey, meta);
    dispatch(tr);
  }
  return true;
}

/**
 * Check if suggestion mode is currently active.
 */
export function isSuggestionModeActive(state: EditorState): boolean {
  return suggestionModeKey.getState(state)?.active ?? false;
}
