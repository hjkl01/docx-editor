/**
 * Tracked structural changes — paragraph-mark insertion/deletion (issue #614).
 *
 * Verifies that pressing Enter in suggesting mode produces a tracked
 * paragraph-mark insertion (`pPrIns`) on the FIRST of the two resulting
 * paragraphs, and that Backspace at the start of a paragraph produces a
 * tracked paragraph-mark deletion (`pPrDel`) on the previous paragraph —
 * not an untracked structural edit. Then exercises accept/reject by id.
 *
 * Background: https://github.com/eigenpal/docx-editor/issues/614
 */

import { test, expect } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

type ParagraphRevision = {
  pPrIns: { revisionId: number; author: string; date: string | null } | null;
  pPrDel: { revisionId: number; author: string; date: string | null } | null;
};

async function getParaRevision(
  page: import('@playwright/test').Page,
  index: number
): Promise<ParagraphRevision | null> {
  return page.evaluate((i) => {
    const hook = window.__DOCX_EDITOR_E2E__;
    return hook?.getParagraphRevisionAt(i) ?? null;
  }, index);
}

async function setSuggestionMode(
  page: import('@playwright/test').Page,
  active: boolean,
  author?: string
) {
  const ok = await page.evaluate(
    ({ a, u }) => window.__DOCX_EDITOR_E2E__?.setSuggestionMode(a, u) ?? false,
    { a: active, u: author }
  );
  // Re-focus the editor: the meta dispatch can trigger a React re-render
  // path that briefly loses contentEditable focus, which makes subsequent
  // page.keyboard.press('Backspace') target the wrong element.
  await page.locator('.ProseMirror').first().focus();
  return ok;
}

async function acceptById(page: import('@playwright/test').Page, id: number) {
  return page.evaluate((rid) => window.__DOCX_EDITOR_E2E__?.acceptChangeById(rid) ?? false, id);
}

async function rejectById(page: import('@playwright/test').Page, id: number) {
  return page.evaluate((rid) => window.__DOCX_EDITOR_E2E__?.rejectChangeById(rid) ?? false, id);
}

test.describe('Tracked paragraph-mark revisions (issue #614)', () => {
  let editor: EditorPage;

  test.beforeEach(async ({ page }) => {
    editor = new EditorPage(page);
    await editor.gotoEmpty();
    await editor.waitForReady();
    await editor.focus();
  });

  test('Enter in suggesting mode sets pPrIns on the first paragraph', async ({ page }) => {
    await editor.typeText('Hello world');
    // Move caret to between "Hello" and " world".
    await editor.selectRange(0, 0, 5);
    await page.keyboard.press('ArrowRight'); // collapse to end of selection
    expect(await setSuggestionMode(page, true, 'Jane')).toBe(true);

    await editor.pressEnter();

    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);

    expect(first).not.toBeNull();
    expect(second).not.toBeNull();

    // First paragraph (the new pilcrow) carries pPrIns.
    expect(first?.pPrIns).not.toBeNull();
    expect(first?.pPrIns?.author).toBe('Jane');
    expect(first?.pPrDel).toBeNull();
    // Second paragraph carries no new revision attr.
    expect(second?.pPrIns).toBeNull();
    expect(second?.pPrDel).toBeNull();
  });

  test('Backspace at paragraph start in suggesting mode sets pPrDel on previous paragraph', async ({
    page,
  }) => {
    await editor.typeText('Hello');
    await editor.pressEnter();
    await editor.typeText('world');
    // Place caret at start of "world".
    await page.keyboard.press('Home');

    expect(await setSuggestionMode(page, true, 'Jane')).toBe(true);
    await page.keyboard.press('Backspace');
    // Wait for React/PM to flush the Backspace transaction before reading attrs.
    await page.waitForTimeout(150);

    // Paragraphs still split — the join is DEFERRED until accept.
    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);
    expect(first).not.toBeNull();
    expect(second).not.toBeNull();

    expect(first?.pPrDel).not.toBeNull();
    expect(first?.pPrDel?.author).toBe('Jane');
    expect(first?.pPrIns).toBeNull();
    expect(second?.pPrDel).toBeNull();
  });

  test('Backspace at the very start of the document is a no-op', async ({ page }) => {
    await editor.typeText('only paragraph');
    await page.keyboard.press('Home');

    expect(await setSuggestionMode(page, true)).toBe(true);
    await page.keyboard.press('Backspace');

    const first = await getParaRevision(page, 0);
    expect(first?.pPrIns).toBeNull();
    expect(first?.pPrDel).toBeNull();
  });

  test('Accept pPrIns clears the marker; paragraphs stay split', async ({ page }) => {
    await editor.typeText('Hello world');
    await editor.selectRange(0, 0, 5);
    await page.keyboard.press('ArrowRight');
    await setSuggestionMode(page, true, 'Jane');
    await editor.pressEnter();

    const before = await getParaRevision(page, 0);
    const revId = before?.pPrIns?.revisionId;
    expect(typeof revId).toBe('number');

    expect(await acceptById(page, revId as number)).toBe(true);

    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);
    expect(first?.pPrIns).toBeNull();
    expect(second).not.toBeNull(); // still split
  });

  test('Reject pPrIns joins the two paragraphs back together', async ({ page }) => {
    await editor.typeText('Hello world');
    await editor.selectRange(0, 0, 5);
    await page.keyboard.press('ArrowRight');
    await setSuggestionMode(page, true, 'Jane');
    await editor.pressEnter();

    const before = await getParaRevision(page, 0);
    const revId = before?.pPrIns?.revisionId as number;

    expect(await rejectById(page, revId)).toBe(true);

    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);
    expect(first?.pPrIns).toBeNull();
    expect(second).toBeNull(); // second paragraph gone — join happened
  });

  test('Accept pPrDel joins the paragraphs (matches Word)', async ({ page }) => {
    await editor.typeText('Hello');
    await editor.pressEnter();
    await editor.typeText('world');
    await page.keyboard.press('Home');

    await setSuggestionMode(page, true, 'Jane');
    await page.keyboard.press('Backspace');
    // Wait for React to flush the Backspace transaction. Without this the
    // `getParaRevision` read can race against the dispatch and see stale attrs.
    await page.waitForTimeout(150);

    const before = await getParaRevision(page, 0);
    expect(before?.pPrDel, 'Backspace must set pPrDel on previous paragraph').not.toBeNull();
    const revId = before?.pPrDel?.revisionId as number;

    expect(await acceptById(page, revId)).toBe(true);

    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);
    expect(first?.pPrDel).toBeNull();
    expect(second).toBeNull();
  });

  test('Reject pPrDel clears the marker; paragraphs stay split', async ({ page }) => {
    await editor.typeText('Hello');
    await editor.pressEnter();
    await editor.typeText('world');
    await page.keyboard.press('Home');

    await setSuggestionMode(page, true, 'Jane');
    await page.keyboard.press('Backspace');
    await page.waitForTimeout(150);

    const before = await getParaRevision(page, 0);
    expect(before?.pPrDel, 'Backspace must set pPrDel on previous paragraph').not.toBeNull();
    const revId = before?.pPrDel?.revisionId as number;

    expect(await rejectById(page, revId)).toBe(true);

    const first = await getParaRevision(page, 0);
    const second = await getParaRevision(page, 1);
    expect(first?.pPrDel).toBeNull();
    expect(second).not.toBeNull();
  });

  test('Painted paragraph fragments carry data-revision-id and the pilcrow class', async ({
    page,
  }) => {
    await editor.typeText('Hello world');
    await editor.selectRange(0, 0, 5);
    await page.keyboard.press('ArrowRight');
    await setSuggestionMode(page, true, 'Jane');
    await editor.pressEnter();

    // The painter renders into .layout-paragraph fragments. The first
    // fragment of the inserted paragraph should carry the revision attrs
    // and the layout-revision-ins class.
    const insMark = page
      .locator('.layout-paragraph.layout-revision-pmark.layout-revision-ins')
      .first();
    await expect(insMark).toBeVisible();
    const revId = await insMark.getAttribute('data-revision-id');
    expect(revId).toMatch(/^\d+$/);
    expect(await insMark.getAttribute('data-revision-author')).toBe('Jane');
  });

  test('acceptChangeById on an unknown revisionId is a no-op (returns false)', async ({ page }) => {
    await editor.typeText('untouched');
    expect(await acceptById(page, 999999)).toBe(false);
  });

  test('pPrChange round-trips and reject restores prior alignment', async ({ page }) => {
    // Manually plant a pPrChange entry via the PM view, then verify
    // acceptChangeById clears it and rejectChangeById restores prior fields.
    await editor.typeText('Hello');
    await page.evaluate(() => {
      const w = window as unknown as {
        __DOCX_EDITOR_E2E__?: {
          plantParagraphPropertyChange?: (revisionId: number, prior: unknown) => boolean;
        };
      };
      w.__DOCX_EDITOR_E2E__?.plantParagraphPropertyChange?.(99, {
        alignment: 'left',
        indentLeft: 0,
      });
    });
    // (test-only helper — see App.tsx; falls through if not present)
    // Sanity: the pPrChange attr is present.
    const before = await page.evaluate(() => {
      const w = window as unknown as {
        __DOCX_EDITOR_E2E__?: { getParagraphAttrs?: (i: number) => Record<string, unknown> | null };
      };
      return w.__DOCX_EDITOR_E2E__?.getParagraphAttrs?.(0) ?? null;
    });
    // The helper IS wired in examples/vite/src/App.tsx; assert presence rather
    // than skipping, so a future regression of the helper surfaces here.
    expect(before, 'plantParagraphPropertyChange helper must populate pPrChange').toBeTruthy();
    expect((before as Record<string, unknown>).pPrChange).toBeTruthy();
    // Reject restores `alignment: 'left'` via applyPriorParagraphFormattingToAttrs.
    const ok = await rejectById(page, 99);
    expect(ok).toBe(true);
    const after = await page.evaluate(() => {
      const w = window as unknown as {
        __DOCX_EDITOR_E2E__?: { getParagraphAttrs?: (i: number) => Record<string, unknown> | null };
      };
      return w.__DOCX_EDITOR_E2E__?.getParagraphAttrs?.(0) ?? null;
    });
    expect(after).not.toBeNull();
    expect((after as Record<string, unknown>).pPrChange).toBeNull();
    expect((after as Record<string, unknown>).alignment).toBe('left');
  });

  test('trIns round-trips and acceptChangeById clears the marker on the row', async ({ page }) => {
    // Plant a 1×1 table at the cursor via PM dispatch, then plant trIns on
    // the row and verify acceptChangeById clears it.
    await page.evaluate(() => window.__DOCX_EDITOR_E2E__?.plantSimpleTable?.());
    const planted = await page.evaluate(
      () => window.__DOCX_EDITOR_E2E__?.plantTableRowInsertion?.(77) ?? false
    );
    expect(planted, 'plantTableRowInsertion must populate trIns on the first row').toBe(true);

    const rowAttrs = await page.evaluate(
      () => window.__DOCX_EDITOR_E2E__?.getFirstTableRowAttrs?.() ?? null
    );
    expect(rowAttrs?.trIns).toBeTruthy();
    expect((rowAttrs?.trIns as { revisionId: number }).revisionId).toBe(77);

    // Accept clears the marker (Phase 2 round-trip semantic; full row-remove
    // semantics come with suggesting-aware commands).
    expect(await acceptById(page, 77)).toBe(true);
    const afterAttrs = await page.evaluate(
      () => window.__DOCX_EDITOR_E2E__?.getFirstTableRowAttrs?.() ?? null
    );
    expect(afterAttrs?.trIns).toBeNull();
  });

  test('Sidebar surfaces a paragraph-mark revision card with accept/reject', async ({ page }) => {
    await editor.typeText('Hello world');
    await editor.selectRange(0, 0, 5);
    await page.keyboard.press('ArrowRight');
    await setSuggestionMode(page, true, 'Jane');
    await editor.pressEnter();

    // Open the unified sidebar so revision cards render.
    const toggle = page.locator('[aria-label="Toggle comments sidebar"]');
    if ((await toggle.getAttribute('aria-pressed')) !== 'true') {
      await toggle.click();
      await page.waitForTimeout(150);
    }

    // The new TrackedChangeCard uses the existing `.docx-tracked-change-card`
    // class — confirm an entry shows up for the paragraph-mark revision.
    const card = page.locator('.docx-unified-sidebar .docx-tracked-change-card').first();
    await expect(card).toBeVisible();
    await expect(card).toContainText('Jane');
    await expect(card).toContainText('Inserted paragraph break');
  });
});
