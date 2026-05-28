/**
 * Single source of tracked-revision ids across the package.
 *
 * Why a shared module-level counter:
 *
 *   - Each `revisionId` is the OOXML `<w:ins w:id="…"/>` attribute. Two
 *     unrelated revisions emitted by the same author with the same id
 *     would silently collapse in the sidebar (grouped by id+author+date)
 *     and on accept (resolved by id), so collisions are observable.
 *   - Pre-refactor, three call sites each kept their own
 *     `Date.now() + offset` counter (suggestionMode.ts, table commands
 *     delete.ts, table commands insert.ts). Offsets made first-load
 *     collisions rare but not impossible — counters drift independently,
 *     and parallel-browser Playwright workers can start from identical
 *     `Date.now()` seeds.
 *
 * Routing every mint through this module guarantees within-realm
 * uniqueness even when callers interleave across plugins, commands, and
 * test harnesses.
 *
 * @internal
 * @packageDocumentation
 */

let counter = Date.now();

/** Mint the next tracked-revision id (`w:id`). Strictly monotonic per realm. */
export function mintRevisionId(): number {
  return counter++;
}
