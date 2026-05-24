/**
 * Build a fresh `RenderContext` from a caller's options. Shared by the
 * paged and non-paged entry points so default resolution lives in one place.
 */

import type { MarkdownOptionsBase, PagedMarkdownOptions, RenderContext } from './types';

export function newContext(
  opts: MarkdownOptionsBase & { headerFooter?: 'strip' | 'first-page' | 'all' } = {}
): RenderContext {
  return {
    opts: {
      annotations: opts.annotations ?? 'html',
      trackedChanges: opts.trackedChanges ?? 'annotate',
      comments: opts.comments ?? 'inline',
      hyperlinks: opts.hyperlinks ?? 'inline',
      footnotes: 'end',
      headerFooter: opts.headerFooter ?? 'strip',
      imagePath: opts.imagePath,
    },
    images: new Map(),
    imagesByPath: new Map(),
    warnings: [],
    footnoteRefs: [],
    commentRefs: [],
    hyperlinkRefs: [],
    imageCounter: 0,
  };
}

// Re-export so paged.ts can take a tighter input type while sharing impl.
export type { PagedMarkdownOptions };

/**
 * Push a warning into the context, deduplicating against existing entries.
 * Recurring messages (e.g. merged-cell warnings on a 200-cell table) appear
 * at most once, matching the contract in `MarkdownResult.warnings`.
 */
export function pushWarning(ctx: RenderContext, message: string): void {
  if (!ctx.warnings.includes(message)) ctx.warnings.push(message);
}
