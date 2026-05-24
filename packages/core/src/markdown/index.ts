/**
 * DOCX to Markdown converter.
 *
 * Public entry points:
 *
 * - `toMarkdown` produces a single markdown string.
 * - `toMarkdownPaged` produces one markdown string per Word page plus a
 *   `combined` string with `<!-- page N -->` separators.
 * - `toMarkdownAsync` and `toMarkdownPagedAsync` wrap the above with an
 *   `imageHandler` callback that the converter awaits per image to substitute
 *   the default `![alt](./images/...)` reference with something else (an LLM
 *   description, an uploaded URL, a data URL, an empty string to drop, ...).
 *
 * Each function accepts a parsed `Document` (sync) or raw DOCX bytes
 * (`Buffer` / `Uint8Array` / `ArrayBuffer`, async because parsing is async).
 *
 * See {@link MarkdownOptions} and {@link PagedMarkdownOptions} for the option
 * matrix and {@link MarkdownResult} for the return shape.
 *
 * @example Parse a buffer and print markdown
 * ```ts
 * import { toMarkdown } from '@eigenpal/docx-editor-core/markdown';
 * import { readFile } from 'node:fs/promises';
 *
 * const { markdown, images, warnings } = await toMarkdown(await readFile('doc.docx'));
 * console.log(markdown);
 * ```
 *
 * @packageDocumentation
 */

import type { Comment, Document, Footnote } from '../types/document';
import { parseDocx } from '../docx/parser';
import { renderBlocks } from './renderBlock';
import type { MarkdownOptions, MarkdownResult, RenderContext } from './types';

export type {
  ImageMeta,
  ImageRef,
  ImageHandler,
  MarkdownOptions,
  MarkdownOptionsBase,
  MarkdownResult,
  PagedMarkdownOptions,
  PagedMarkdownResult,
} from './types';

export { toMarkdownPaged } from './paged';
export { toMarkdownAsync, toMarkdownPagedAsync } from './async';

type ByteInput = Uint8Array | ArrayBuffer;

/**
 * Narrow an `unknown` input to a parsed `Document`. Checks that the input is
 * a non-null object with a `.package.document` slot, the minimum shape the
 * renderer reads from.
 */
export function isDocument(input: unknown): input is Document {
  if (typeof input !== 'object' || input === null) return false;
  const pkg = (input as { package?: unknown }).package;
  if (typeof pkg !== 'object' || pkg === null) return false;
  return typeof (pkg as { document?: unknown }).document === 'object';
}

/**
 * Build a descriptive `Error` for inputs that aren't a `Document` or a known
 * byte type. The message names the function and hints at common mistakes
 * (passing a `File` or a `Blob` without awaiting `.arrayBuffer()`).
 */
export function badInputError(fnName: string, input: unknown): Error {
  const got =
    input === null
      ? 'null'
      : typeof input === 'object'
        ? Object.prototype.toString.call(input)
        : typeof input;
  const hint =
    got === '[object File]' || got === '[object Blob]' ? ' (await file.arrayBuffer() first)' : '';
  return new Error(
    `${fnName} expected Buffer, Uint8Array, ArrayBuffer, or Document. Received: ${got}${hint}.`
  );
}

/**
 * Convert a parsed `Document` (or raw DOCX bytes) to markdown.
 *
 * With a `Document`, the call is synchronous. With raw bytes, the function
 * parses the DOCX first and returns a `Promise<MarkdownResult>`.
 *
 * @example From a buffer (Node)
 * ```ts
 * import { toMarkdown } from '@eigenpal/docx-editor-core/markdown';
 * import { readFile } from 'node:fs/promises';
 *
 * const { markdown } = await toMarkdown(await readFile('contract.docx'));
 * ```
 *
 * @example From a pre-parsed Document
 * ```ts
 * import { parseDocx, toMarkdown } from '@eigenpal/docx-editor-core/markdown';
 *
 * const doc = await parseDocx(buffer);
 * const result = toMarkdown(doc, { trackedChanges: 'clean' });
 * ```
 *
 * @public
 */
export function toMarkdown(doc: Document, opts?: MarkdownOptions): MarkdownResult;
export function toMarkdown(buffer: ByteInput, opts?: MarkdownOptions): Promise<MarkdownResult>;
export function toMarkdown(
  input: Document | ByteInput,
  opts?: MarkdownOptions
): MarkdownResult | Promise<MarkdownResult> {
  if (isDocument(input)) return renderDocumentSync(input, opts);
  if (input instanceof Uint8Array || input instanceof ArrayBuffer) {
    return parseDocx(input).then((doc) => renderDocumentSync(doc, opts));
  }
  throw badInputError('toMarkdown', input);
}

function renderDocumentSync(doc: Document, opts: MarkdownOptions = {}): MarkdownResult {
  const ctx = newContext(opts);
  const body = renderBlocks(ctx, doc.package, doc.package.document.content);
  const markdown = assembleDocument(ctx, doc, body);
  if (doc.warnings) ctx.warnings.unshift(...doc.warnings);
  if (!markdown.trim()) ctx.warnings.push('document has no content');
  return { markdown, images: ctx.images, warnings: ctx.warnings };
}

function assembleDocument(ctx: RenderContext, doc: Document, body: string): string {
  const sections: string[] = [];
  if (body.trim()) sections.push(body);

  if (ctx.opts.footnotes !== 'inline' && ctx.footnoteRefs.length) {
    const refs = ctx.footnoteRefs.map(({ refId, markerNumber }) => {
      const note = doc.package.footnotes?.find((f) => f.id === refId);
      return `[^${markerNumber}]: ${note ? extractFootnoteText(ctx, doc, note) : ''}`;
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
      const text = c ? extractCommentText(c) : '';
      lines.push(`[^c${markerNumber}]: ${author}${text}`);
    }
    sections.push(lines.join('\n'));
  }

  return sections.join('\n\n');
}

function extractFootnoteText(ctx: RenderContext, doc: Document, note: Footnote): string {
  return renderBlocks(ctx, doc.package, note.content).replace(/\n+/g, ' ').trim();
}

function extractCommentText(comment: Comment): string {
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

function newContext(opts: MarkdownOptions): RenderContext {
  return {
    opts: {
      annotations: opts.annotations ?? 'html',
      trackedChanges: opts.trackedChanges ?? 'annotate',
      comments: opts.comments ?? 'inline',
      hyperlinks: opts.hyperlinks ?? 'inline',
      footnotes: opts.footnotes ?? 'end',
      headerFooter: 'strip',
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
