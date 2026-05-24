/**
 * Render a Word table as a GFM markdown table.
 *
 * Limitations (each emits a single deduped warning):
 *   - Merged cells (`gridSpan` / `vMerge`). Markdown has no merged cells.
 *     Content goes in the first cell of the merge; remaining cells render
 *     as empty placeholders.
 *   - Nested tables. Flattened to `<br>`-joined text inside the parent cell.
 *   - Multi-paragraph cells. Joined with `<br>` so the row stays on one line.
 */

import type { DocxPackage, Table, TableCell, TableRow } from '../types/document';
import { renderParagraph } from './renderParagraph';
import { escapeTableCell } from './escape';
import type { RenderContext } from './types';

interface RenderTableOptions {
  /** Rows to slice. Defaults to all rows in the table. */
  rowRange?: { from: number; to: number };
  /** When true, the first row is treated as the header row. Default true. */
  firstRowIsHeader?: boolean;
}

/**
 * Render a table. Caller picks the row range; paged mode uses this to slice
 * a `TableFragment`.
 */
export function renderTable(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  table: Table,
  options: RenderTableOptions = {}
): string {
  if (!table.rows.length) return '';
  const from = options.rowRange?.from ?? 0;
  const to = options.rowRange?.to ?? table.rows.length;
  const rows = table.rows.slice(from, to);
  if (!rows.length) return '';
  const firstRowIsHeader = options.firstRowIsHeader !== false;

  const cellTexts = rows.map((row) => renderRow(ctx, pkg, row));
  const maxCols = cellTexts.reduce((m, c) => Math.max(m, c.length), 0);
  if (!maxCols) return '';

  // Pad each row to maxCols.
  const padded = cellTexts.map((row) => {
    const out = row.slice();
    while (out.length < maxCols) out.push('');
    return out;
  });

  const lines: string[] = [];
  if (firstRowIsHeader) {
    lines.push(toRowLine(padded[0]));
    lines.push(`| ${new Array(maxCols).fill('---').join(' | ')} |`);
    for (let i = 1; i < padded.length; i++) lines.push(toRowLine(padded[i]));
  } else {
    // No header row → synthesize an empty header so the table still parses as GFM.
    lines.push(`| ${new Array(maxCols).fill('').join(' | ')} |`);
    lines.push(`| ${new Array(maxCols).fill('---').join(' | ')} |`);
    for (const row of padded) lines.push(toRowLine(row));
  }

  return lines.join('\n');
}

function toRowLine(cells: string[]): string {
  return `| ${cells.join(' | ')} |`;
}

function warnOnce(ctx: RenderContext, message: string): void {
  if (!ctx.warnings.includes(message)) ctx.warnings.push(message);
}

function renderRow(ctx: RenderContext, pkg: DocxPackage | undefined, row: TableRow): string[] {
  const out: string[] = [];
  for (const cell of row.cells) {
    out.push(renderCell(ctx, pkg, cell));
    const span = cell.formatting?.gridSpan ?? 1;
    if (span > 1) {
      warnOnce(ctx, 'merged cells (gridSpan) not representable in GFM');
      for (let i = 1; i < span; i++) out.push('');
    }
  }
  return out;
}

function renderCell(ctx: RenderContext, pkg: DocxPackage | undefined, cell: TableCell): string {
  if (cell.formatting?.vMerge === 'continue') {
    warnOnce(ctx, 'merged cells (vMerge) not representable in GFM');
    return '';
  }
  const blocks: string[] = [];
  for (const item of cell.content) {
    if (item.type === 'paragraph') {
      const md = renderParagraph(ctx, pkg, item);
      if (md.trim()) blocks.push(md);
    } else if (item.type === 'table') {
      warnOnce(ctx, 'nested table flattened to text');
      blocks.push(renderTable(ctx, pkg, item));
    }
  }
  return escapeTableCell(blocks.join('\n'));
}
