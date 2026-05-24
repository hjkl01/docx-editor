/**
 * Shared trailers: footnote definitions, hyperlink reference list, comments
 * sidecar. Same shape for paged and non-paged output, so both entry points
 * call into here instead of duplicating the emission logic.
 */

import type { Comment, Document, Footnote } from '../types/document';
import { renderBlocks } from './renderBlock';
import type { RenderContext } from './types';

/**
 * Append footnote definitions / hyperlink list / comments sidecar to the body.
 * Sections are emitted only when their relevant context has entries.
 */
export function appendTrailers(ctx: RenderContext, doc: Document, body: string): string {
  const sections: string[] = body.trim() ? [body] : [];

  if (ctx.footnoteRefs.length) {
    const refs = ctx.footnoteRefs.map(({ refId, markerNumber }) => {
      const note = doc.package.footnotes?.find((f) => f.id === refId);
      return `[^${markerNumber}]: ${note ? footnoteText(ctx, doc, note) : ''}`;
    });
    sections.push(refs.join('\n'));
  }

  if (ctx.opts.hyperlinks === 'reference' && ctx.hyperlinkRefs.length) {
    sections.push(
      ctx.hyperlinkRefs.map(({ href, refNumber }) => `[${refNumber}]: ${href}`).join('\n')
    );
  }

  if (ctx.opts.comments === 'sidecar' && ctx.commentRefs.length) {
    const lines: string[] = ['## Comments'];
    for (const { commentId, markerNumber } of ctx.commentRefs) {
      const c = doc.package.document.comments?.find((cm) => cm.id === commentId);
      const author = c?.author ? `${c.author}: ` : '';
      const text = c ? commentText(c) : '';
      lines.push(`[^c${markerNumber}]: ${author}${text}`);
    }
    sections.push(lines.join('\n'));
  }

  return sections.join('\n\n');
}

function footnoteText(ctx: RenderContext, doc: Document, note: Footnote): string {
  return renderBlocks(ctx, doc.package, note.content).replace(/\n+/g, ' ').trim();
}

function commentText(comment: Comment): string {
  return comment.content
    .map((p) =>
      p.content
        .map((c) =>
          c.type === 'run' ? c.content.map((x) => (x.type === 'text' ? x.text : '')).join('') : ''
        )
        .join('')
    )
    .join(' ')
    .trim();
}
