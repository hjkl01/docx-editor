/**
 * Render a single paragraph as a block of markdown.
 *
 * Three cases:
 *   1. Heading style → `#`/`##`/.../`######`.
 *   2. List item → indented marker + inline content (Word's exact marker
 *      preserved verbatim).
 *   3. Plain prose → escaped inline content.
 */

import type { DocxPackage, Paragraph } from '../types/document';
import { parseHeadingLevel, isHeadingStyle } from '../agent/text-utils';
import { renderParagraphInline } from './renderRuns';
import type { RenderContext } from './types';

/**
 * Render a paragraph and return the block text. No surrounding blank line:
 * the caller joins blocks.
 */
export function renderParagraph(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  para: Paragraph
): string {
  const inline = renderParagraphInline(ctx, pkg, para.content, para.paraId);
  const styleId = para.formatting?.styleId;

  // Heading?
  if (isHeadingStyle(styleId)) {
    const level = parseHeadingLevel(styleId) ?? 1;
    const hashes = '#'.repeat(Math.max(1, Math.min(6, level)));
    return `${hashes} ${inline}`.trimEnd();
  }

  // List item?
  if (para.listRendering) {
    return renderListItem(para, inline);
  }

  // Code block (no native code-block detection; if the entire paragraph uses
  // a monospace run, our run renderer already wrapped it in backticks).

  // Quote block? Word uses a "Quote" / "IntenseQuote" style.
  if (styleId && /quote/i.test(styleId)) {
    return inline
      .split('\n')
      .map((line) => `> ${line}`)
      .join('\n');
  }

  return inline;
}

function renderListItem(para: Paragraph, inline: string): string {
  const list = para.listRendering!;
  const indent = '  '.repeat(list.level);
  if (list.isBullet) {
    return `${indent}- ${inline}`.trimEnd();
  }
  // Preserve Word's exact marker (e.g. "1.", "a)", "i."). Strip trailing
  // whitespace from the marker but keep its punctuation intact.
  const marker = list.marker.trim();
  return `${indent}${marker} ${inline}`.trimEnd();
}
