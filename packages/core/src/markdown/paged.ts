/**
 * Paged DOCX-to-Markdown.
 *
 * Page boundaries are inferred from Word's pre-baked pagination hints:
 * `paragraph.renderedPageBreakBefore` flags, explicit `<w:br w:type="page"/>`
 * inside runs, and section breaks of type `nextPage`/`evenPage`/`oddPage`.
 * No canvas or measurement is required. Documents that Word has rendered at
 * least once carry these hints; programmatically generated DOCX files often
 * do not, in which case the whole document renders as a single page.
 *
 * Each page's blocks are rendered by the same `renderBlock` pipeline used by
 * `toMarkdown`, so a paragraph that fits on one page is byte-identical
 * between the paged and continuous outputs.
 *
 * @packageDocumentation
 */

import type {
  BlockContent,
  DocxPackage,
  Document,
  HeaderFooter,
  Paragraph,
  Run,
} from '../types/document';
import { renderBlocks } from './renderBlock';
import { wrapHeaderFooter } from './annotations';
import { badInputError, isDocument } from './index';
import type { PagedMarkdownOptions, PagedMarkdownResult, RenderContext } from './types';
import { parseDocx } from '../docx/parser';

type ByteInput = Uint8Array | ArrayBuffer;

/**
 * Convert a parsed `Document` (or raw DOCX bytes) to markdown, one entry per
 * page plus a `combined` string with `<!-- page N -->` separators.
 *
 * With a `Document`, the call is synchronous. With raw bytes, parsing runs
 * first and the result is wrapped in a `Promise`.
 *
 * @example Parse and split a buffer into pages
 * ```ts
 * import { toMarkdownPaged } from '@eigenpal/docx-editor-core/markdown';
 * import { readFile } from 'node:fs/promises';
 *
 * const buf = await readFile('contract.docx');
 * const { pages, combined } = await toMarkdownPaged(buf);
 * for (const p of pages) {
 *   console.log(`--- page ${p.pageNumber} ---\n${p.markdown}`);
 * }
 * ```
 *
 * @public
 */
export function toMarkdownPaged(doc: Document, opts?: PagedMarkdownOptions): PagedMarkdownResult;
export function toMarkdownPaged(
  buffer: ByteInput,
  opts?: PagedMarkdownOptions
): Promise<PagedMarkdownResult>;
export function toMarkdownPaged(
  input: Document | ByteInput,
  opts?: PagedMarkdownOptions
): PagedMarkdownResult | Promise<PagedMarkdownResult> {
  if (isDocument(input)) return renderPagedSync(input, opts);
  if (input instanceof Uint8Array || input instanceof ArrayBuffer) {
    return parseDocx(input).then((doc) => renderPagedSync(doc, opts));
  }
  throw badInputError('toMarkdownPaged', input);
}

function renderPagedSync(
  doc: Document,
  opts: PagedMarkdownOptions | undefined
): PagedMarkdownResult {
  const ctx = newPagedContext(opts ?? {});
  const blocks = doc.package.document.content;
  if (!blocks.length) {
    if (doc.warnings) ctx.warnings.unshift(...doc.warnings);
    ctx.warnings.push('document has no content');
    return { pages: [], combined: '', images: ctx.images, warnings: ctx.warnings };
  }
  const pages = splitIntoPages(blocks);
  const renderedPages: Array<{ pageNumber: number; markdown: string }> = [];

  pages.forEach((pageBlocks, idx) => {
    const pageNumber = idx + 1;
    ctx.pageNumber = pageNumber;
    const sections: string[] = [];

    // Header/footer for this page
    if (ctx.opts.headerFooter !== 'strip') {
      const hf = resolveHeaderFooter(doc.package, pageNumber);
      if (hf.header && shouldEmitHeaderFooter(ctx, pageNumber)) {
        const inner = renderBlocks(ctx, doc.package, hf.header.content);
        if (inner) sections.push(wrapHeaderFooter(ctx, 'header', inner));
      }
    }

    sections.push(renderBlocks(ctx, doc.package, pageBlocks));

    if (ctx.opts.headerFooter !== 'strip') {
      const hf = resolveHeaderFooter(doc.package, pageNumber);
      if (hf.footer && shouldEmitHeaderFooter(ctx, pageNumber)) {
        const inner = renderBlocks(ctx, doc.package, hf.footer.content);
        if (inner) sections.push(wrapHeaderFooter(ctx, 'footer', inner));
      }
    }

    renderedPages.push({
      pageNumber,
      markdown: sections.filter((s) => s.trim()).join('\n\n'),
    });
  });

  if (ctx.opts.footnotes === 'end' && ctx.footnoteRefs.length) {
    const refs = ctx.footnoteRefs
      .map(({ refId, markerNumber }) => {
        const note = doc.package.footnotes?.find((f) => f.id === refId);
        const text = note
          ? renderBlocks(ctx, doc.package, note.content).replace(/\n+/g, ' ').trim()
          : '';
        return `[^${markerNumber}]: ${text}`;
      })
      .join('\n');
    if (renderedPages.length) {
      renderedPages[renderedPages.length - 1].markdown += '\n\n' + refs;
    }
  }

  // Comments sidecar appended to last page
  if (ctx.opts.comments === 'sidecar' && ctx.commentRefs.length) {
    const lines = ['## Comments'];
    for (const { commentId, markerNumber } of ctx.commentRefs) {
      const c = doc.package.document.comments?.find((cm) => cm.id === commentId);
      const author = c?.author ? `${c.author}: ` : '';
      const text = c
        ? c.content
            .map((p) =>
              p.content
                .map((cc) =>
                  cc.type === 'run'
                    ? cc.content.map((x) => (x.type === 'text' ? x.text : '')).join('')
                    : ''
                )
                .join('')
            )
            .join(' ')
            .trim()
        : '';
      lines.push(`[^c${markerNumber}]: ${author}${text}`);
    }
    if (renderedPages.length) {
      renderedPages[renderedPages.length - 1].markdown += '\n\n' + lines.join('\n');
    }
  }

  // Hyperlink ref list appended to last page
  if (ctx.opts.hyperlinks === 'reference' && ctx.hyperlinkRefs.length) {
    const refs = ctx.hyperlinkRefs
      .map(({ href, refNumber }) => `[${refNumber}]: ${href}`)
      .join('\n');
    if (renderedPages.length) {
      renderedPages[renderedPages.length - 1].markdown += '\n\n' + refs;
    }
  }

  const combined = renderedPages
    .map((p, i) => (i === 0 ? p.markdown : `<!-- page ${p.pageNumber} -->\n\n${p.markdown}`))
    .join('\n\n');

  return {
    pages: renderedPages,
    combined,
    images: ctx.images,
    warnings: ctx.warnings,
  };
}

/** Heuristic page splitter: walks blocks, starts a new page on each break signal. */
function splitIntoPages(blocks: BlockContent[]): BlockContent[][] {
  if (!blocks.length) return [[]];
  const pages: BlockContent[][] = [[]];
  let pendingBreakAfter = false;

  const startNewPage = () => {
    if (pages[pages.length - 1].length) pages.push([]);
  };

  for (const block of blocks) {
    if (pendingBreakAfter) {
      startNewPage();
      pendingBreakAfter = false;
    }
    if (block.type === 'paragraph') {
      if (startsNewPage(block)) startNewPage();
      pages[pages.length - 1].push(block);
      // A paragraph containing an explicit mid-paragraph page break splits
      // the document at the *next* paragraph (we keep the source paragraph
      // whole rather than splitting its inline runs. Line-accurate splits
      // are reserved for a layout-driven mode.
      if (containsExplicitPageBreak(block)) pendingBreakAfter = true;
    } else {
      pages[pages.length - 1].push(block);
    }
  }
  return pages;
}

function startsNewPage(para: Paragraph): boolean {
  if (para.renderedPageBreakBefore) return true;
  if (para.formatting?.pageBreakBefore) return true;
  const sectionStart = para.sectionProperties?.sectionStart;
  return sectionStart === 'nextPage' || sectionStart === 'evenPage' || sectionStart === 'oddPage';
}

function containsExplicitPageBreak(para: Paragraph): boolean {
  return para.content.some(
    (c) =>
      c.type === 'run' &&
      (c as Run).content.some((r) => r.type === 'break' && r.breakType === 'page')
  );
}

function resolveHeaderFooter(
  pkg: DocxPackage,
  pageNumber: number
): { header?: HeaderFooter; footer?: HeaderFooter } {
  const section = pkg.document.sections?.[0];
  if (!section) return {};
  const isFirstPage = pageNumber === 1;
  const { headers, footers } = section;
  return {
    header: (isFirstPage && headers?.get('first')) || headers?.get('default'),
    footer: (isFirstPage && footers?.get('first')) || footers?.get('default'),
  };
}

function shouldEmitHeaderFooter(ctx: RenderContext, pageNumber: number): boolean {
  if (ctx.opts.headerFooter === 'first-page') return pageNumber === 1;
  return ctx.opts.headerFooter === 'all';
}

function newPagedContext(opts: PagedMarkdownOptions): RenderContext {
  return {
    opts: {
      annotations: opts.annotations ?? 'html',
      trackedChanges: opts.trackedChanges ?? 'annotate',
      comments: opts.comments ?? 'inline',
      hyperlinks: opts.hyperlinks ?? 'inline',
      footnotes: opts.footnotes ?? 'end',
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
    pageNumber: undefined,
  };
}
