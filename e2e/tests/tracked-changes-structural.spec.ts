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
  return page.evaluate(({ a, u }) => window.__DOCX_EDITOR_E2E__?.setSuggestionMode(a, u) ?? false, {
    a: active,
    u: author,
  });
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

    const before = await getParaRevision(page, 0);
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

    const before = await getParaRevision(page, 0);
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
