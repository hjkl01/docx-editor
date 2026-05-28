/**
 * Walk the PM doc once and derive (a) the tracked-change list and (b) a
 * comment→revision overlap map for threading. Adjacent entries from the
 * same revision are merged; deletion+insertion pairs from the same
 * author/date become a single `replacement` entry (matches Word's UX
 * for replace ops).
 *
 * Pure function — no React, no Vue, no side effects. Single O(N) walk
 * over text nodes. Lifted from packages/react/src/hooks/useTrackedChanges.ts
 * so both adapters can call it directly.
 *
 * @remarks
 * Tagged `@internal` post-1.0 cut. Both adapters re-export this through
 * their own composables (`useTrackedChanges`); consumers should prefer
 * those. The subpath stays in `package.json` `exports` for back-compat;
 * expect it to move behind a public surface in a future major.
 *
 * @packageDocumentation
 * @internal
 */
import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import type { TrackedChangeEntry } from '../../utils/comments';

/** @internal */
export interface TrackedChangesResult {
  /** Tracked-change entries, sorted by document position, with adjacent same-revision entries merged. */
  entries: TrackedChangeEntry[];
  /**
   * Map of `commentId -> revisionId` for comments whose range overlaps a tracked-change mark.
   * Consumers (DocxEditor's threading effect) use this to thread comments under their tracked change.
   */
  commentToRevision: Map<number, number>;
}

const EMPTY_RESULT: TrackedChangesResult = {
  entries: [],
  commentToRevision: new Map(),
};

/** @internal */
export function extractTrackedChanges(state: EditorState | null): TrackedChangesResult {
  if (!state) return EMPTY_RESULT;
  const { doc, schema } = state;
  const insertionType = schema.marks.insertion;
  const deletionType = schema.marks.deletion;
  const commentType = schema.marks.comment;
  if (!insertionType && !deletionType) return EMPTY_RESULT;

  const raw: TrackedChangeEntry[] = [];
  const commentToRevision = new Map<number, number>();
  doc.descendants((node, pos) => {
    // Structural revisions on the paragraph mark itself
    // (`<w:pPr><w:rPr><w:ins/>` / `<w:del/>`). Surface as their own entry
    // types so the sidebar can label and dispatch them correctly.
    if (node.type.name === 'paragraph') {
      const ins = node.attrs.pPrIns as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      const del = node.attrs.pPrDel as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      if (ins) {
        raw.push({
          type: 'paragraphMarkInsertion',
          text: node.textContent || '',
          author: ins.author || '',
          date: ins.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: ins.revisionId,
        });
      }
      if (del) {
        raw.push({
          type: 'paragraphMarkDeletion',
          text: node.textContent || '',
          author: del.author || '',
          date: del.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: del.revisionId,
        });
      }
      // Paragraph-property changes — one entry per (id, author, date) entry
      // in the pPrChange array. Reject restores prior values; accept clears.
      const pPrChange = node.attrs.pPrChange as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(pPrChange)) {
        for (const entry of pPrChange) {
          raw.push({
            type: 'paragraphPropertiesChanged',
            text: node.textContent || '',
            author: entry.info.author || '',
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
      // Descend into paragraph content; do not return here.
    }

    // Table-row revisions (`<w:trPr><w:ins/>` / `<w:del/>` / `<w:trPrChange>`).
    if (node.type.name === 'tableRow') {
      const trIns = node.attrs.trIns as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      const trDel = node.attrs.trDel as {
        revisionId: number;
        author: string;
        date: string | null;
      } | null;
      if (trIns) {
        raw.push({
          type: 'rowInserted',
          text: node.textContent || '',
          author: trIns.author || '',
          date: trIns.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: trIns.revisionId,
        });
      }
      if (trDel) {
        raw.push({
          type: 'rowDeleted',
          text: node.textContent || '',
          author: trDel.author || '',
          date: trDel.date ?? undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: trDel.revisionId,
        });
      }
      const trPrChange = node.attrs.trPrChange as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(trPrChange)) {
        for (const entry of trPrChange) {
          if (!entry?.info || typeof entry.info.id !== 'number') continue;
          raw.push({
            type: 'rowPropertiesChanged',
            text: node.textContent || '',
            author: entry.info.author || '',
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
      // Descend into cells.
    }

    // Table-cell revisions (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>`,
    // `<w:tcPrChange>`).
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      const cellMarker = node.attrs.cellMarker as {
        kind: 'ins' | 'del' | 'merge';
        info: { revisionId: number; author: string; date: string | null };
      } | null;
      if (cellMarker?.info && typeof cellMarker.info.revisionId === 'number') {
        const kindToType = {
          ins: 'cellInserted' as const,
          del: 'cellDeleted' as const,
          merge: 'cellMerged' as const,
        };
        const resolvedType = kindToType[cellMarker.kind];
        if (resolvedType) {
          raw.push({
            type: resolvedType,
            text: node.textContent || '',
            author: cellMarker.info.author || '',
            date: cellMarker.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: cellMarker.info.revisionId,
          });
        }
      }
      const tcPrChange = node.attrs.tcPrChange as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(tcPrChange)) {
        for (const entry of tcPrChange) {
          if (!entry?.info || typeof entry.info.id !== 'number') continue;
          raw.push({
            type: 'cellPropertiesChanged',
            text: node.textContent || '',
            author: entry.info.author || '',
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
    }

    // Table-level property change (`<w:tblPrChange>`).
    if (node.type.name === 'table') {
      const tblPrChange = node.attrs.tblPrChange as Array<{
        info: { id: number; author: string; date?: string };
      }> | null;
      if (Array.isArray(tblPrChange)) {
        for (const entry of tblPrChange) {
          if (!entry?.info || typeof entry.info.id !== 'number') continue;
          raw.push({
            type: 'tablePropertiesChanged',
            text: '',
            author: entry.info.author || '',
            date: entry.info.date ?? undefined,
            from: pos,
            to: pos + node.nodeSize,
            revisionId: entry.info.id,
          });
        }
      }
    }

    if (!node.isText) return;
    let tcMark: Mark | null = null;
    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        raw.push({
          type: mark.type === insertionType ? 'insertion' : 'deletion',
          text: node.text || '',
          author: (mark.attrs.author as string) || '',
          date: mark.attrs.date as string | undefined,
          from: pos,
          to: pos + node.nodeSize,
          revisionId: mark.attrs.revisionId as number,
        });
        tcMark = mark;
      }
    }
    if (commentType && tcMark) {
      const commentMark = node.marks.find((m) => m.type === commentType);
      if (commentMark) {
        const cid = commentMark.attrs.commentId as number;
        const rid = tcMark.attrs.revisionId as number;
        if (!commentToRevision.has(cid)) commentToRevision.set(cid, rid);
      }
    }
  });

  // Coalesce structural-revision entries that share a `(id, author, date)`
  // triple across nested nodes. A row-insertion typically produces one
  // `rowInserted` entry on the `<tr>` PLUS one `cellInserted` entry per
  // cell, all sharing the triple. The spec says these should render as a
  // single sidebar row (per `tracked-structural-tables/spec.md` "Sidebar
  // groups co-revision-id entries as one"). Prefer the broader entry:
  // priority is `table > row > cell > paragraph-mark`.
  //
  // Inline insertion/deletion entries are NOT coalesced here — the
  // adjacent-merge pass below handles them.
  //
  // Single-pass: track each triple's slot index in `ordered` so an
  // in-place replacement is O(1) (vs `ordered.indexOf(existing)` which
  // would be O(n) inside an O(n) loop).
  const STRUCTURAL_PRIORITY: Record<string, number> = {
    tablePropertiesChanged: 5,
    rowInserted: 4,
    rowDeleted: 4,
    rowPropertiesChanged: 4,
    cellInserted: 3,
    cellDeleted: 3,
    cellMerged: 3,
    cellPropertiesChanged: 3,
    paragraphMarkInsertion: 2,
    paragraphMarkDeletion: 2,
    paragraphPropertiesChanged: 2,
  };
  const isStructuralType = (t: TrackedChangeEntry['type']) => t in STRUCTURAL_PRIORITY;
  const slotByKey = new Map<string, number>();
  const ordered: TrackedChangeEntry[] = [];
  for (const entry of raw) {
    if (!isStructuralType(entry.type)) {
      ordered.push(entry);
      continue;
    }
    const key = `${entry.revisionId}|${entry.author}|${entry.date ?? ''}`;
    const slot = slotByKey.get(key);
    if (slot === undefined) {
      slotByKey.set(key, ordered.push(entry) - 1);
      continue;
    }
    // Keep the entry with HIGHER structural priority (broader scope wins).
    if ((STRUCTURAL_PRIORITY[entry.type] ?? 0) > (STRUCTURAL_PRIORITY[ordered[slot]!.type] ?? 0)) {
      ordered[slot] = entry;
    }
    // Otherwise drop the new entry — broader sibling already represents it.
  }

  // Merge adjacent entries with the same revisionId and type into one.
  // Restricted to INLINE types (insertion / deletion) — paragraph-mark entries
  // use the whole-paragraph range, and two consecutive paragraphs sharing a
  // revision triple (legal in OOXML for a multi-paragraph insertion applied
  // under one id) must stay as distinct sidebar rows.
  const merged: TrackedChangeEntry[] = [];
  for (const entry of ordered) {
    const last = merged[merged.length - 1];
    const isInlineType = entry.type === 'insertion' || entry.type === 'deletion';
    if (
      isInlineType &&
      last &&
      last.revisionId === entry.revisionId &&
      last.type === entry.type &&
      last.to === entry.from
    ) {
      last.text += entry.text;
      last.to = entry.to;
    } else {
      merged.push({ ...entry });
    }
  }

  // Detect replacement pairs: adjacent deletion + insertion from the
  // same author/date. Word assigns different w:id values but same
  // author+date for a single replace.
  const final: TrackedChangeEntry[] = [];
  for (let i = 0; i < merged.length; i++) {
    const curr = merged[i]!;
    const next = merged[i + 1];
    if (
      curr.type === 'deletion' &&
      next &&
      next.type === 'insertion' &&
      curr.author === next.author &&
      curr.date === next.date &&
      curr.to === next.from
    ) {
      final.push({
        type: 'replacement',
        text: next.text,
        deletedText: curr.text,
        author: curr.author,
        date: curr.date,
        from: curr.from,
        to: next.to,
        revisionId: curr.revisionId,
        insertionRevisionId: next.revisionId,
      });
      i++;
    } else {
      final.push(curr);
    }
  }
  return { entries: final, commentToRevision };
}
