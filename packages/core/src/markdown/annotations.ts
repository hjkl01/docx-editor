/**
 * Encoders for constructs markdown can't express natively: tracked changes,
 * comments, headers, footers. Three modes: `html` (default), `pandoc`,
 * `strip`. Renderers call the wrap functions with the already-rendered
 * inner text; the wrappers decide whether to emit an `<ins>`/`<del>`/
 * `<comment>` HTML tag, the Pandoc-flavored bracketed span equivalent, or
 * to drop the wrapper entirely.
 */

import type { TrackedChangeInfo } from '../types/document';
import type { RenderContext } from './types';

type Mode = 'tracked' | 'comment';

function escapeAttr(value: string): string {
  return value.replace(/"/g, '&quot;');
}

function attrString(info: { author?: string; date?: string; id?: number }): string {
  const parts: string[] = [];
  if (info.author) parts.push(`author="${escapeAttr(info.author)}"`);
  if (info.date) parts.push(`date="${escapeAttr(info.date)}"`);
  if (typeof info.id === 'number') parts.push(`id="${info.id}"`);
  return parts.length ? ' ' + parts.join(' ') : '';
}

/**
 * Wrap `inner` with the active annotations mode's encoding.
 *
 * @param keepWhenStripped - what to emit when `annotations === 'strip'`.
 *   Insertions and comments keep the visible text; deletions drop entirely.
 */
function wrap(
  ctx: RenderContext,
  inner: string,
  htmlTag: string,
  pandocClass: string,
  info: { author?: string; date?: string; id?: number },
  extraHtmlAttrs: string,
  mode: Mode,
  keepWhenStripped: boolean
): string {
  if (ctx.opts.annotations === 'strip') return keepWhenStripped ? inner : '';
  if (ctx.opts.annotations === 'pandoc') {
    const author = info.author ? ` author="${escapeAttr(info.author)}"` : '';
    const id = mode === 'comment' && typeof info.id === 'number' ? ` id="${info.id}"` : '';
    return `[${inner}]{.${pandocClass}${id}${author}}`;
  }
  return `<${htmlTag}${attrString(info)}${extraHtmlAttrs}>${inner}</${htmlTag}>`;
}

export function wrapInsertion(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  return wrap(ctx, inner, 'ins', 'ins', info, '', 'tracked', true);
}

export function wrapDeletion(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  return wrap(ctx, inner, 'del', 'del', info, '', 'tracked', false);
}

export function wrapMoveFrom(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  return wrap(ctx, inner, 'del', 'move-from', info, ' data-move="from"', 'tracked', false);
}

export function wrapMoveTo(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  return wrap(ctx, inner, 'ins', 'move-to', info, ' data-move="to"', 'tracked', true);
}

export function wrapComment(
  ctx: RenderContext,
  meta: { id: number; author?: string },
  inner: string
): string {
  return wrap(ctx, inner, 'comment', 'comment', meta, '', 'comment', true);
}

/**
 * Block wrapper for headers/footers.
 *
 * Blank lines between the wrapper tag and inner content make each side
 * parse as its own HTML block (CommonMark type-7), so markdown inside
 * (bold, links, lists, images, tables) is parsed normally by GFM renderers
 * while the `<header>` / `<footer>` tags pass through as raw HTML.
 */
export function wrapHeaderFooter(
  ctx: RenderContext,
  kind: 'header' | 'footer',
  inner: string
): string {
  if (ctx.opts.annotations === 'strip') return '';
  if (ctx.opts.annotations === 'pandoc') return `:::${kind}\n\n${inner}\n\n:::`;
  return `<${kind}>\n\n${inner}\n\n</${kind}>`;
}
