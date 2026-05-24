/**
 * Async variants of `toMarkdown` / `toMarkdownPaged` that route every image
 * reference through an `imageHandler` callback and substitute the returned
 * string into the markdown.
 *
 * The handler receives the full `ImageRef` (bytes, base64, data URL, MIME
 * type, page number, alt text) plus the virtual path the markdown currently
 * uses. It returns the markdown chunk to substitute. Typical handlers return
 * `![description](url)` after uploading or describing the image, or `""` to
 * drop the image entirely.
 *
 * Per-image handler errors are caught, recorded in `warnings`, and the
 * default reference is left in place. The overall call only rejects when the
 * underlying `toMarkdown` / `toMarkdownPaged` parse fails.
 */

import type { Document } from '../types/document';
import { toMarkdown } from './index';
import { toMarkdownPaged } from './paged';
import type {
  ImageHandler,
  ImageRef,
  MarkdownOptions,
  MarkdownResult,
  PagedMarkdownOptions,
  PagedMarkdownResult,
} from './types';

type ByteInput = Uint8Array | ArrayBuffer;

/**
 * Same as {@link toMarkdown} but applies `opts.imageHandler` to each image.
 * The handler's return value replaces the default `![alt](virtualPath)`
 * reference in the output.
 *
 * @example Describe each image with a vision model
 * ```ts
 * import { toMarkdownAsync } from '@eigenpal/docx-editor-core/markdown';
 *
 * const { markdown } = await toMarkdownAsync(buffer, {
 *   imageHandler: async (ref) => {
 *     const text = await describe(ref.base64, ref.mimeType);
 *     return `![${text}]()`;
 *   },
 * });
 * ```
 *
 * @public
 */
export async function toMarkdownAsync(
  input: Document | ByteInput,
  opts: MarkdownOptions & { imageHandler: ImageHandler }
): Promise<MarkdownResult> {
  const base =
    input instanceof Uint8Array || input instanceof ArrayBuffer
      ? await toMarkdown(input, opts)
      : toMarkdown(input, opts);
  const markdown = await substituteImages(
    base.markdown,
    base.images,
    opts.imageHandler,
    base.warnings
  );
  return { ...base, markdown };
}

/**
 * Same as {@link toMarkdownPaged} but applies `opts.imageHandler` to each
 * image. Substitution runs per page so the handler sees the right
 * `ref.pageNumber` for images that appear inside a page's body.
 *
 * @public
 */
export async function toMarkdownPagedAsync(
  input: Document | ByteInput,
  opts: PagedMarkdownOptions & { imageHandler: ImageHandler }
): Promise<PagedMarkdownResult> {
  const base =
    input instanceof Uint8Array || input instanceof ArrayBuffer
      ? await toMarkdownPaged(input, opts)
      : toMarkdownPaged(input, opts);
  const pages = await Promise.all(
    base.pages.map(async (p) => ({
      pageNumber: p.pageNumber,
      markdown: await substituteImages(
        p.markdown,
        base.images,
        opts.imageHandler,
        base.warnings,
        p.pageNumber
      ),
    }))
  );
  const combined = pages
    .map((p, i) => (i === 0 ? p.markdown : `<!-- page ${p.pageNumber} -->\n\n${p.markdown}`))
    .join('\n\n');
  return { ...base, pages, combined };
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Resolve every registered image reference through the handler and rewrite
 * the markdown accordingly. The image map's keys are used as literal needles
 * so custom `imagePath` callbacks work too (the default `./images/...`
 * scheme is not assumed).
 */
async function substituteImages(
  markdown: string,
  images: Map<string, ImageRef>,
  handler: ImageHandler,
  warnings: string[],
  pageNumber?: number
): Promise<string> {
  if (!images.size || !markdown) return markdown;

  const entries = [...images.entries()].filter(([virtualPath]) => markdown.includes(virtualPath));
  if (!entries.length) return markdown;

  const replacements = await Promise.all(
    entries.map(async ([virtualPath, ref]) => {
      try {
        const result = await handler(ref, {
          virtualPath,
          pageNumber: ref.pageNumber ?? pageNumber,
        });
        return [virtualPath, ref, result] as const;
      } catch (err) {
        const cause = err instanceof Error ? err.message : String(err);
        warnings.push(`imageHandler failed for ${virtualPath}: ${cause}`);
        return [virtualPath, ref, null] as const;
      }
    })
  );

  let out = markdown;
  for (const [virtualPath, ref, result] of replacements) {
    if (result === null) continue;
    // Replace `![<alt>](<virtualPath>)` exactly. Alt is `[^\]]*` per CommonMark.
    const pattern = new RegExp(`!\\[[^\\]]*\\]\\(${escapeRegExp(virtualPath)}\\)`, 'g');
    out = out.replace(pattern, () => result);
    void ref;
  }
  return out;
}
