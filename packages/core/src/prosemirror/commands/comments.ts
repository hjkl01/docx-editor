/**
 * Comment and Track Changes Commands
 *
 * PM commands for adding/removing comments and accepting/rejecting tracked changes.
 */

import type { Command, Transaction } from 'prosemirror-state';
import type { EditorState } from 'prosemirror-state';
import { SUGGESTION_BYPASS_META } from '../plugins/suggestionMode';

/**
 * Add a comment mark to the current selection.
 */
export function addCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;

    const commentType = state.schema.marks.comment;
    if (!commentType) return false;

    if (dispatch) {
      const tr = state.tr.addMark(from, to, commentType.create({ commentId }));
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Remove a comment mark by ID from the entire document.
 */
export function removeCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const commentType = state.schema.marks.comment;
    if (!commentType) return false;

    if (dispatch) {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        if (node.isText) {
          for (const mark of node.marks) {
            if (mark.type === commentType && mark.attrs.commentId === commentId) {
              tr.removeMark(pos, pos + node.nodeSize, mark);
            }
          }
        }
      });
      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

/**
 * Resolve a tracked change: accept or reject.
 * - Accept: keep insertions (remove mark), delete deletions (remove text)
 * - Reject: keep deletions (remove mark), delete insertions (remove text)
 */
function resolveChange(from: number, to: number, mode: 'accept' | 'reject'): Command {
  return (state, dispatch) => {
    const insertionType = state.schema.marks.insertion;
    const deletionType = state.schema.marks.deletion;
    if (!insertionType && !deletionType) return false;

    // "keep" mark type: remove the mark but keep the text
    // "remove" mark type: remove both the mark and the text
    const keepType = mode === 'accept' ? insertionType : deletionType;
    const removeType = mode === 'accept' ? deletionType : insertionType;

    if (dispatch) {
      const tr = state.tr;
      const deleteRanges: Array<{ from: number; to: number }> = [];

      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const nodeEnd = pos + node.nodeSize;
        const rangeFrom = Math.max(from, pos);
        const rangeTo = Math.min(to, nodeEnd);

        if (removeType && node.marks.some((m) => m.type === removeType)) {
          deleteRanges.push({ from: rangeFrom, to: rangeTo });
        }

        if (keepType && node.marks.some((m) => m.type === keepType)) {
          tr.removeMark(rangeFrom, rangeTo, keepType);
        }
      });

      for (const range of deleteRanges.reverse()) {
        tr.delete(range.from, range.to);
      }

      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

/**
 * Accept a tracked change at the given range.
 * - Insertion: remove mark, keep text
 * - Deletion: remove mark AND text
 */
export function acceptChange(from: number, to: number): Command {
  return resolveChange(from, to, 'accept');
}

/**
 * Reject a tracked change at the given range.
 * - Insertion: remove mark AND text
 * - Deletion: remove mark, keep text
 */
export function rejectChange(from: number, to: number): Command {
  return resolveChange(from, to, 'reject');
}

/**
 * Accept all tracked changes in the document.
 */
export function acceptAllChanges(): Command {
  return (state, dispatch) => {
    return acceptChange(0, state.doc.content.size)(state, dispatch);
  };
}

/**
 * Reject all tracked changes in the document.
 */
export function rejectAllChanges(): Command {
  return (state, dispatch) => {
    return rejectChange(0, state.doc.content.size)(state, dispatch);
  };
}

interface ChangeRange {
  from: number;
  to: number;
  type: 'insertion' | 'deletion';
}

/**
 * Find the next tracked change after the given position.
 */
export function findNextChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType && !deletionType) return null;

  let result: ChangeRange | null = null;

  state.doc.descendants((node, pos) => {
    if (result) return false;
    if (!node.isText) return;
    if (pos + node.nodeSize <= startPos) return;

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        result = {
          from: Math.max(pos, startPos),
          to: pos + node.nodeSize,
          type: mark.type === insertionType ? 'insertion' : 'deletion',
        };
        return false;
      }
    }
  });

  // Wrap around (only once)
  if (!result && startPos > 0) {
    return findNextChange(state, 0);
  }

  return result;
}

// ============================================================================
// REVISION-ID-ADDRESSABLE COMMANDS (structural revisions on node attrs)
// ============================================================================

interface ParagraphMarkSite {
  /** Position immediately before the paragraph's open tag. */
  pos: number;
  /** The paragraph node carrying the revision attr. */
  // Kept as `any` here to avoid the Node import cycle; callers handle typing.
  node: import('prosemirror-model').Node;
  kind: 'pPrIns' | 'pPrDel';
}

interface ParagraphPropertyChangeSite {
  pos: number;
  node: import('prosemirror-model').Node;
  /** Index into the paragraph's `pPrChange` array (since multiple authors can stack). */
  entryIndex: number;
  /** The prior `ParagraphFormatting` snapshot from the matching entry. */
  prior: import('../../types/document').ParagraphFormatting | undefined;
}

/**
 * Walk the document and collect every paragraph that carries a
 * `pPrIns` or `pPrDel` attr with the given revision id.
 */
function findParagraphMarkSites(state: EditorState, revisionId: number): ParagraphMarkSite[] {
  const sites: ParagraphMarkSite[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return;
    const ins = node.attrs.pPrIns as { revisionId: number } | null;
    const del = node.attrs.pPrDel as { revisionId: number } | null;
    if (ins && ins.revisionId === revisionId) {
      sites.push({ pos, node, kind: 'pPrIns' });
    }
    if (del && del.revisionId === revisionId) {
      sites.push({ pos, node, kind: 'pPrDel' });
    }
  });
  return sites;
}

/**
 * Walk the document for paragraph nodes whose `pPrChange` array has an
 * entry with `info.id === revisionId`. Returns one site per matching entry
 * with the entry index for later mutation.
 */
function findParagraphPropertyChangeSites(
  state: EditorState,
  revisionId: number
): ParagraphPropertyChangeSite[] {
  const sites: ParagraphPropertyChangeSite[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return;
    const changes = node.attrs.pPrChange as Array<{
      info: { id: number };
      previousFormatting?: unknown;
    }> | null;
    if (!Array.isArray(changes)) return;
    changes.forEach((entry, idx) => {
      if (entry.info.id === revisionId) {
        sites.push({
          pos,
          node,
          entryIndex: idx,
          prior: entry.previousFormatting as
            | import('../../types/document').ParagraphFormatting
            | undefined,
        });
      }
    });
  });
  return sites;
}

/** Find every inline `insertion`/`deletion` mark range with the given id. */
function findInlineMarkSites(
  state: EditorState,
  revisionId: number
): Array<{ from: number; to: number; markName: 'insertion' | 'deletion' }> {
  const sites: Array<{ from: number; to: number; markName: 'insertion' | 'deletion' }> = [];
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  state.doc.descendants((node, pos) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (
        (insertionType && mark.type === insertionType) ||
        (deletionType && mark.type === deletionType)
      ) {
        if (mark.attrs.revisionId === revisionId) {
          const markName: 'insertion' | 'deletion' =
            mark.type === insertionType ? 'insertion' : 'deletion';
          // Coalesce contiguous siblings sharing the same id.
          const last = sites[sites.length - 1];
          if (last && last.markName === markName && last.to === pos) {
            last.to = pos + node.nodeSize;
          } else {
            sites.push({ from: pos, to: pos + node.nodeSize, markName });
          }
        }
      }
    }
  });
  return sites;
}

/**
 * Join paragraph at position `paraStart` (start-of-open-tag) with the
 * following sibling paragraph. The joined paragraph inherits the
 * SECOND paragraph's pPr (matches Word: the surviving mark wins). Both
 * paragraphs must exist; caller checks.
 */
function joinParagraphWithNext(
  tr: Transaction,
  paraStart: number,
  options: { inheritFromSecond: boolean }
): void {
  const para = tr.doc.nodeAt(paraStart);
  if (!para) return;
  const nextParaStart = paraStart + para.nodeSize;
  const nextPara = tr.doc.nodeAt(nextParaStart);
  if (!nextPara || nextPara.type.name !== 'paragraph') return;
  // Per-OOXML: rejecting a paragraph-mark insertion (or accepting a deletion)
  // collapses the boundary; the resulting paragraph's properties come from
  // the SECOND paragraph (the one whose mark survives the join).
  if (options.inheritFromSecond) {
    // Replace para's attrs with nextPara's attrs first, then join.
    tr.setNodeMarkup(paraStart, undefined, { ...nextPara.attrs, pPrIns: null, pPrDel: null });
  }
  // `tr.join(pos)` joins the block ending immediately before `pos` with
  // the block starting at `pos`. `nextParaStart` is between the two paragraphs.
  tr.join(nextParaStart);
}

/** Clear pPrIns/pPrDel attrs on the paragraph at `paraStart`. */
function clearParagraphMarkRevision(
  tr: Transaction,
  paraStart: number,
  kind: 'pPrIns' | 'pPrDel'
): void {
  const para = tr.doc.nodeAt(paraStart);
  if (!para) return;
  const newAttrs = { ...para.attrs };
  newAttrs[kind] = null;
  tr.setNodeMarkup(paraStart, undefined, newAttrs);
}

/**
 * Remove a `pPrChange` entry by array index from the paragraph at `paraStart`.
 * If the array becomes empty, the attr is set to `null` so PM treats it as
 * absent on save.
 */
function clearParagraphPropertyChangeEntry(
  tr: Transaction,
  paraStart: number,
  entryIndex: number
): void {
  const para = tr.doc.nodeAt(paraStart);
  if (!para) return;
  const existing = para.attrs.pPrChange as Array<unknown> | null;
  if (!Array.isArray(existing) || entryIndex < 0 || entryIndex >= existing.length) return;
  const next = existing.slice();
  next.splice(entryIndex, 1);
  tr.setNodeMarkup(paraStart, undefined, {
    ...para.attrs,
    pPrChange: next.length > 0 ? next : null,
  });
}

/**
 * Restore fields from a prior `ParagraphFormatting` snapshot onto the
 * paragraph's PM attrs. Only the user-visible fields are copied — anything
 * not in `prior` is left untouched.
 */
function applyPriorParagraphFormattingToAttrs(
  attrs: Record<string, unknown>,
  prior: import('../../types/document').ParagraphFormatting
): Record<string, unknown> {
  const next: Record<string, unknown> = { ...attrs };
  const fields: Array<keyof import('../../types/document').ParagraphFormatting> = [
    'alignment',
    'spaceBefore',
    'spaceAfter',
    'lineSpacing',
    'lineSpacingRule',
    'indentLeft',
    'indentRight',
    'indentFirstLine',
    'hangingIndent',
    'styleId',
    'borders',
    'shading',
    'tabs',
    'pageBreakBefore',
    'keepNext',
    'keepLines',
    'contextualSpacing',
    'bidi',
    'outlineLevel',
    'numPr',
  ];
  for (const f of fields) {
    if (Object.prototype.hasOwnProperty.call(prior, f)) {
      next[f as string] = prior[f] ?? null;
    }
  }
  return next;
}

/**
 * Resolve every site sharing a revision id in one PM transaction. Bypass
 * the suggesting-mode keymap (we're applying, not authoring).
 *
 * Per-marker semantics (see openspec/changes/tracked-structural-changes):
 *   accept pPrIns → clear marker, keep split.
 *   reject pPrIns → join with following paragraph; result inherits second's pPr.
 *   accept pPrDel → join with following paragraph; result inherits second's pPr.
 *   reject pPrDel → clear marker, keep split.
 *   accept insertion mark → keep text, drop mark.
 *   reject insertion mark → remove text and mark.
 *   accept deletion mark → remove text and mark.
 *   reject deletion mark → keep text, drop mark.
 */
function resolveById(revisionId: number, mode: 'accept' | 'reject'): Command {
  return (state, dispatch) => {
    const paragraphMarkSites = findParagraphMarkSites(state, revisionId);
    const inlineSites = findInlineMarkSites(state, revisionId);
    const propChangeSites = findParagraphPropertyChangeSites(state, revisionId);
    if (
      paragraphMarkSites.length === 0 &&
      inlineSites.length === 0 &&
      propChangeSites.length === 0
    ) {
      return false;
    }

    if (!dispatch) return true;

    const tr = state.tr;
    tr.setMeta(SUGGESTION_BYPASS_META, true);

    // Process inline marks FIRST (positions still valid in original doc), in
    // reverse order so deletions don't shift earlier positions.
    const insertionType = state.schema.marks.insertion;
    const deletionType = state.schema.marks.deletion;
    const sortedInline = [...inlineSites].sort((a, b) => b.from - a.from);
    for (const site of sortedInline) {
      const isInsertion = site.markName === 'insertion';
      const removeText = (mode === 'accept' && !isInsertion) || (mode === 'reject' && isInsertion);
      if (removeText) {
        tr.delete(site.from, site.to);
      } else {
        const markType = isInsertion ? insertionType : deletionType;
        if (markType) tr.removeMark(site.from, site.to, markType);
      }
    }

    // Then process paragraph-mark revisions. Process in reverse position order
    // so earlier-positioned joins don't shift later sites. Track resolved
    // positions through tr.mapping.
    const sortedPara = [...paragraphMarkSites].sort((a, b) => b.pos - a.pos);
    for (const site of sortedPara) {
      const mappedPos = tr.mapping.map(site.pos);
      const liveNode = tr.doc.nodeAt(mappedPos);
      if (!liveNode || liveNode.type.name !== 'paragraph') continue;
      // Re-confirm the revision is still on the live node (inline-mark
      // deletions above may have removed text but won't have removed the
      // paragraph attrs).
      const stillHasIns =
        site.kind === 'pPrIns' &&
        (liveNode.attrs.pPrIns as { revisionId: number } | null)?.revisionId === revisionId;
      const stillHasDel =
        site.kind === 'pPrDel' &&
        (liveNode.attrs.pPrDel as { revisionId: number } | null)?.revisionId === revisionId;
      if (!stillHasIns && !stillHasDel) continue;

      const shouldJoin =
        (mode === 'reject' && site.kind === 'pPrIns') ||
        (mode === 'accept' && site.kind === 'pPrDel');

      if (shouldJoin) {
        // No following paragraph → just clear the marker (last-paragraph edge case).
        const liveParaEnd = mappedPos + liveNode.nodeSize;
        const after = tr.doc.nodeAt(liveParaEnd);
        if (!after || after.type.name !== 'paragraph') {
          clearParagraphMarkRevision(tr, mappedPos, site.kind);
        } else {
          // Clear our marker first, then perform the join inheriting the
          // second paragraph's pPr.
          clearParagraphMarkRevision(tr, mappedPos, site.kind);
          joinParagraphWithNext(tr, mappedPos, { inheritFromSecond: true });
        }
      } else {
        clearParagraphMarkRevision(tr, mappedPos, site.kind);
      }
    }

    // Finally, paragraph-property changes. Accept clears the matching entry
    // (current props win). Reject restores the entry's `prior` fields onto
    // the paragraph's attrs and clears the entry.
    const sortedPropChanges = [...propChangeSites].sort((a, b) => b.pos - a.pos);
    for (const site of sortedPropChanges) {
      const mappedPos = tr.mapping.map(site.pos);
      const liveNode = tr.doc.nodeAt(mappedPos);
      if (!liveNode || liveNode.type.name !== 'paragraph') continue;
      const liveChanges = liveNode.attrs.pPrChange as Array<{ info: { id: number } }> | null;
      if (!Array.isArray(liveChanges)) continue;
      const liveIndex = liveChanges.findIndex((e) => e.info.id === revisionId);
      if (liveIndex < 0) continue;

      if (mode === 'reject' && site.prior) {
        // Restore prior fields BEFORE clearing the entry so we don't lose
        // the snapshot in the intermediate state.
        const restored = applyPriorParagraphFormattingToAttrs(liveNode.attrs, site.prior);
        const nextChanges = liveChanges.slice();
        nextChanges.splice(liveIndex, 1);
        tr.setNodeMarkup(mappedPos, undefined, {
          ...restored,
          pPrChange: nextChanges.length > 0 ? nextChanges : null,
        });
      } else {
        clearParagraphPropertyChangeEntry(tr, mappedPos, liveIndex);
      }
    }

    if (tr.steps.length === 0) return false;
    dispatch(tr);
    return true;
  };
}

/**
 * Accept any revision in the document by its `w:id`. Resolves every site
 * sharing the id (paragraph-mark attrs and inline marks) in a single PM
 * transaction. Returns false (no-op) if the id is not present.
 */
export function acceptChangeById(revisionId: number): Command {
  return resolveById(revisionId, 'accept');
}

/**
 * Reject any revision in the document by its `w:id`. Inverse of accept
 * per the Word semantics table in the spec.
 */
export function rejectChangeById(revisionId: number): Command {
  return resolveById(revisionId, 'reject');
}

/**
 * Find the previous tracked change before the given position.
 */
export function findPreviousChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType && !deletionType) return null;

  let result: ChangeRange | null = null;

  state.doc.descendants((node, pos) => {
    if (!node.isText) return;
    if (pos >= startPos) return false;

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        result = {
          from: pos,
          to: pos + node.nodeSize,
          type: mark.type === insertionType ? 'insertion' : 'deletion',
        };
      }
    }
  });

  // Wrap around (only once — guard prevents infinite recursion)
  if (!result && startPos < state.doc.content.size) {
    return findPreviousChange(state, state.doc.content.size);
  }

  return result;
}
