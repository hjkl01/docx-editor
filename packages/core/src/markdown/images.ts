/**
 * Image handling: virtual-path generation and ImageRef construction from
 * MediaFile entries.
 *
 * Default virtual path: `./images/{paraId}-img{n}.{ext}` when paraId is known,
 * otherwise `./images/img{n}.{ext}`. Callers can override via `opts.imagePath`.
 *
 * A media file referenced multiple times by the same document is only
 * registered once: the first `registerImage` call computes base64 and stores
 * the `ImageRef`; subsequent calls return the same ref.
 */

import type { MediaFile } from '../types/styles';
import type { Image } from '../types/document';
import type { ImageMeta, ImageRef, RenderContext } from './types';

const MIME_TO_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/svg+xml': 'svg',
  'image/webp': 'webp',
  'image/tiff': 'tiff',
  'image/bmp': 'bmp',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
};

function extFor(mimeType: string, fallback: string): string {
  return MIME_TO_EXT[mimeType] ?? fallback;
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  // btoa is browser/web; falls back to a manual encoder if absent.
  if (typeof btoa !== 'undefined') return btoa(binary);
  return manualBase64Encode(bytes);
}

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
function manualBase64Encode(bytes: Uint8Array): string {
  let out = '';
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
    out += BASE64_ALPHABET[(n >> 18) & 63];
    out += BASE64_ALPHABET[(n >> 12) & 63];
    out += BASE64_ALPHABET[(n >> 6) & 63];
    out += BASE64_ALPHABET[n & 63];
  }
  if (i < bytes.length) {
    const remaining = bytes.length - i;
    const n = remaining === 2 ? (bytes[i] << 16) | (bytes[i + 1] << 8) : bytes[i] << 16;
    out += BASE64_ALPHABET[(n >> 18) & 63];
    out += BASE64_ALPHABET[(n >> 12) & 63];
    out += remaining === 2 ? BASE64_ALPHABET[(n >> 6) & 63] : '=';
    out += '=';
  }
  return out;
}

function toUint8(data: ArrayBuffer | Uint8Array): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
}

/**
 * Build an `ImageRef` from a `MediaFile` and store it in the context's image
 * map. If the same `media.path` has already been registered for this render,
 * the existing ref is returned without re-encoding base64 or assigning a new
 * virtual path.
 *
 * @param ctx - active render context (mutated)
 * @param media - resolved media entry from `Document.package.media`
 * @param image - the OOXML drawing that referenced this media (alt text source)
 * @param paraId - the paraId of the containing paragraph, used to namespace
 *   the default virtual path; pass `undefined` when not known.
 */
export function registerImage(
  ctx: RenderContext,
  media: MediaFile,
  image: Image | undefined,
  paraId: string | undefined
): ImageRef {
  const existing = ctx.imagesByPath.get(media.path);
  if (existing) return existing;

  ctx.imageCounter += 1;
  const ext = extFor(media.mimeType, 'png');
  const meta: ImageMeta = {
    paraId,
    index: ctx.imageCounter,
    originalPath: media.path,
    mimeType: media.mimeType,
    alt: image?.alt ?? image?.title ?? image?.filename,
    pageNumber: ctx.pageNumber,
  };
  const virtualPath = ctx.opts.imagePath ? ctx.opts.imagePath(meta) : defaultVirtualPath(meta, ext);
  const data = toUint8(media.data);
  const base64 = media.base64 ?? bytesToBase64(data);
  const dataUrl = media.dataUrl ?? `data:${media.mimeType};base64,${base64}`;
  const ref: ImageRef = { ...meta, data, base64, dataUrl, virtualPath };
  ctx.images.set(virtualPath, ref);
  ctx.imagesByPath.set(media.path, ref);
  return ref;
}

function defaultVirtualPath(meta: ImageMeta, ext: string): string {
  if (meta.paraId) return `./images/${meta.paraId}-img${meta.index}.${ext}`;
  return `./images/img${meta.index}.${ext}`;
}
