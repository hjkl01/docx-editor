/**
 * Shared input-handling helpers for the markdown converter.
 *
 * Lives in its own module to break the `index.ts ↔ paged.ts` cycle that
 * existed when both entry points needed the same `Document` typeguard and
 * `Error` builder. Internal only — not re-exported from `index.ts`.
 */

import type { Document } from '../types/document';

/**
 * Narrow an `unknown` input to a parsed `Document`. Checks that the input is
 * a non-null object with `package.document` populated, the minimum shape the
 * renderer reads from.
 */
export function isDocument(input: unknown): input is Document {
  if (typeof input !== 'object' || input === null) return false;
  const pkg = (input as { package?: unknown }).package;
  if (typeof pkg !== 'object' || pkg === null) return false;
  const body = (pkg as { document?: unknown }).document;
  return typeof body === 'object' && body !== null;
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
