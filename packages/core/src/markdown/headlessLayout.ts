/**
 * Layout-engine fallback for paged markdown. Wires the existing core
 * pagination pipeline (Document → ProseDoc → FlowBlock → measureBlocks →
 * layoutDocument) so paged output works even for DOCX files that don't
 * carry Word's pre-baked pagination cache.
 *
 * Lazy-loads `@napi-rs/canvas` as an optional peer dep — caller installs
 * it only when they need this path. In the browser, `document.createElement`
 * already provides a canvas so no extra install is needed.
 *
 * Returns null when canvas can't be obtained, letting the caller fall back
 * to the heuristic splitter.
 */

import type { BlockContent, Document } from '../types/document';
import type {
  FlowBlock,
  Layout,
  Measure,
  ParagraphBlock,
  TableBlock,
} from '../layout-engine/types';
import { assertExhaustiveFlowBlock } from '../layout-engine/types';
import { layoutDocument } from '../layout-engine';
import { toProseDoc } from '../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../layout-bridge/toFlowBlocks';
import { measureParagraph, setCanvasContext } from '../layout-bridge/measuring';
import { measureTableBlock } from '../layout-bridge/measureTable';
import { registerOfficeSubstitutes } from './officeFonts';

let canvasReady: Promise<boolean> | undefined;

/**
 * Bring up a Canvas2D context on the current runtime. Memoized.
 *
 * - In a browser DOM the cached `document.createElement('canvas')` path
 *   inside `measureContainer.ts` already works; no action needed.
 * - In Node / Bun, dynamically import `@napi-rs/canvas` and inject its 2D
 *   context. The peer dep is optional; if it fails to import we return
 *   false and the caller falls back.
 */
async function ensureCanvas(): Promise<boolean> {
  if (canvasReady) return canvasReady;
  canvasReady = (async () => {
    if (typeof document !== 'undefined') return true;
    try {
      const mod = await import('@napi-rs/canvas');
      // Register Office-font substitutes (Carlito, Caladea, Arimo, ...) so
      // the CSS cascade in `buildFontString` resolves to known metrics. Without
      // these, "Calibri" falls through to whatever Skia picks as default and
      // pagination diverges from what the browser produces.
      await registerOfficeSubstitutes(mod);
      const c = mod.createCanvas(2000, 2000);
      const ctx = c.getContext('2d');
      if (!ctx) return false;
      setCanvasContext(ctx as unknown as CanvasRenderingContext2D);
      return true;
    } catch {
      return false;
    }
  })();
  return canvasReady;
}

/**
 * `FlowBlock` measurement dispatcher. Floating-image exclusion zones are
 * omitted: we measure for pagination only, not painting. The exhaustiveness
 * guard catches missing variants at typecheck time.
 */
function measureBlockForLayout(block: FlowBlock, contentWidth: number): Measure {
  switch (block.kind) {
    case 'paragraph':
      return measureParagraph(block as ParagraphBlock, contentWidth);
    case 'table':
      return measureTableBlock(block as TableBlock, contentWidth, measureBlockForLayout);
    case 'image':
      return { kind: 'image', width: block.width ?? 100, height: block.height ?? 100 };
    case 'textBox': {
      const innerMeasures = block.content.map((p) => measureParagraph(p, contentWidth));
      const totalHeight = innerMeasures.reduce((sum, m) => sum + m.totalHeight, 0);
      return {
        kind: 'textBox',
        width: block.width,
        height: block.height ?? totalHeight,
        innerMeasures,
      };
    }
    case 'pageBreak':
      return { kind: 'pageBreak' };
    case 'columnBreak':
      return { kind: 'columnBreak' };
    case 'sectionBreak':
      return { kind: 'sectionBreak' };
    default:
      assertExhaustiveFlowBlock(block, 'markdown headlessLayout');
  }
}

/**
 * Run the layout engine on a parsed document and return a mapping of which
 * source body blocks land on which page.
 *
 * Returns null when canvas isn't available or layout fails — the caller
 * falls back to the heuristic.
 */
export async function computePagedGroups(doc: Document): Promise<BlockContent[][] | null> {
  if (!(await ensureCanvas())) return null;

  let layout: Layout;
  let blocks: FlowBlock[];
  try {
    const pmDoc = toProseDoc(doc);
    blocks = toFlowBlocks(pmDoc);
    const sectPr = doc.package.document.finalSectionProperties;
    const pageWidth = sectPr?.pageWidth ?? 12240; // 8.5in
    const pageHeight = sectPr?.pageHeight ?? 15840; // 11in
    const twip2px = (twips: number): number => (twips / 1440) * 96;
    const pageSize = { w: twip2px(pageWidth), h: twip2px(pageHeight) };
    const margins = {
      top: twip2px(sectPr?.marginTop ?? 1440),
      right: twip2px(sectPr?.marginRight ?? 1440),
      bottom: twip2px(sectPr?.marginBottom ?? 1440),
      left: twip2px(sectPr?.marginLeft ?? 1440),
    };
    const contentWidth = pageSize.w - margins.left - margins.right;
    const measures = blocks.map((b) => measureBlockForLayout(b, contentWidth));
    layout = layoutDocument(blocks, measures, { pageSize, margins });
  } catch {
    return null;
  }

  // Walk source blocks and FlowBlocks in parallel. The order matches because
  // `toProseDoc` and `toFlowBlocks` preserve document order. Synthetic
  // FlowBlocks (sectionBreak / pageBreak / columnBreak) have no source
  // counterpart, so we skip them in the source walk.
  const flowBlockIndexByBlockId = new Map<string | number, number>();
  blocks.forEach((b, i) => {
    flowBlockIndexByBlockId.set(b.id, i);
  });

  const sourceBlocks = doc.package.document.content;
  const sourceIndexByFlowIndex = new Map<number, number>();
  let sourceIdx = 0;
  for (let i = 0; i < blocks.length && sourceIdx < sourceBlocks.length; i++) {
    const fb = blocks[i];
    if (fb.kind === 'sectionBreak' || fb.kind === 'pageBreak' || fb.kind === 'columnBreak')
      continue;
    // Skip the matching source block too if it's metadata-only (very rare).
    sourceIndexByFlowIndex.set(i, sourceIdx);
    sourceIdx += 1;
  }

  // For each page, collect the lowest source index that appears. That tells
  // us where this page's content starts in the source body. Sort pages by
  // their starting source index and slice the source array accordingly.
  const pageStarts: number[] = [];
  for (const page of layout.pages) {
    let pageMinSrc = Number.POSITIVE_INFINITY;
    for (const frag of page.fragments) {
      const fbIdx = flowBlockIndexByBlockId.get(frag.blockId);
      if (fbIdx === undefined) continue;
      const srcIdx = sourceIndexByFlowIndex.get(fbIdx);
      if (srcIdx === undefined) continue;
      if (srcIdx < pageMinSrc) pageMinSrc = srcIdx;
    }
    if (pageMinSrc !== Number.POSITIVE_INFINITY) pageStarts.push(pageMinSrc);
  }
  if (!pageStarts.length) return null;

  // Slice source blocks into page-aligned groups using the starts.
  const groups: BlockContent[][] = [];
  for (let i = 0; i < pageStarts.length; i++) {
    const from = pageStarts[i];
    const to = i + 1 < pageStarts.length ? pageStarts[i + 1] : sourceBlocks.length;
    groups.push(sourceBlocks.slice(from, to));
  }
  return groups;
}
