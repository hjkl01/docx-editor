/**
 * Regression for the move-pair misclassification bug discovered during
 * the issue #614 coalescing work.
 *
 * A tracked replace produces an adjacent `deletion` mark + `insertion`
 * mark in PM. `fromProseDoc/paragraph.ts:340` reads document-wide
 * insertion/deletion counts and, when the SAME `revisionId` appears on
 * both an insertion and a deletion, emits the run as `moveFrom`/`moveTo`
 * (Word's tracked-move shape) instead of `insertion`/`deletion`.
 *
 * Our suggestion-mode plugin must NEVER mint shared `w:id` for the two
 * halves of a replace — otherwise Word renders the edit as a tracked
 * MOVE on save, silently corrupting the user's intent.
 *
 * This test exercises the conversion directly: build a PM doc with del +
 * ins of the same text, distinct revisionIds, and assert the resulting
 * Document model carries `type: 'insertion'` / `type: 'deletion'` runs —
 * NOT `type: 'moveFrom'` / `type: 'moveTo'`.
 */

import { describe, test, expect } from 'bun:test';
import { schema } from '../../schema';
import { fromProseDoc } from '../fromProseDoc';
import type { Paragraph } from '../../../types/document';

function runTypesIn(paragraph: Paragraph): string[] {
  return paragraph.content
    .map((c) => (c as { type?: string }).type)
    .filter((t): t is string => typeof t === 'string');
}

describe('fromProseDoc: replace produces ins/del, not move (regression for #614)', () => {
  test('adjacent deletion + insertion with DIFFERENT revisionIds emit as insertion / deletion', () => {
    const insertionMark = schema.marks.insertion.create({
      revisionId: 100,
      author: 'Jane',
      date: '2026-05-28T00:00:00Z',
    });
    const deletionMark = schema.marks.deletion.create({
      revisionId: 101, // DIFFERENT id — sidebar groups by (author, date)
      author: 'Jane',
      date: '2026-05-28T00:00:00Z',
    });

    const deletedText = schema.text('old', [deletionMark]);
    const insertedText = schema.text('new', [insertionMark]);
    const paragraph = schema.nodes.paragraph.create({}, [deletedText, insertedText]);
    const doc = schema.nodes.doc.create({}, [paragraph]);

    const result = fromProseDoc(doc);
    const para = result.package?.document?.content?.[0] as Paragraph | undefined;
    expect(para).toBeTruthy();
    expect(para?.type).toBe('paragraph');

    const types = runTypesIn(para as Paragraph);
    expect(types).toContain('insertion');
    expect(types).toContain('deletion');
    expect(types).not.toContain('moveFrom');
    expect(types).not.toContain('moveTo');
  });

  test('adjacent deletion + insertion with SAME revisionId still misclassify as move (proves regression case)', () => {
    // This test documents the existing behavior: if the suggestion-mode
    // plugin ever shares ids again, this test will pass — meaning the
    // OOXML output silently becomes a tracked move. The fix lives in
    // suggestionMode.ts (shareDate, not shareAttrs).
    const sharedId = 200;
    const insertionMark = schema.marks.insertion.create({
      revisionId: sharedId,
      author: 'Jane',
      date: '2026-05-28T00:00:00Z',
    });
    const deletionMark = schema.marks.deletion.create({
      revisionId: sharedId,
      author: 'Jane',
      date: '2026-05-28T00:00:00Z',
    });

    const deletedText = schema.text('old', [deletionMark]);
    const insertedText = schema.text('new', [insertionMark]);
    const paragraph = schema.nodes.paragraph.create({}, [deletedText, insertedText]);
    const doc = schema.nodes.doc.create({}, [paragraph]);

    const result = fromProseDoc(doc);
    const para = result.package?.document?.content?.[0] as Paragraph | undefined;
    const types = runTypesIn(para as Paragraph);
    // Shared id triggers the move-pair branch — proving why the
    // suggestion-mode plugin must keep the ids distinct for replaces.
    expect(types).toContain('moveFrom');
    expect(types).toContain('moveTo');
  });
});
