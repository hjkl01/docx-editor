---
'@eigenpal/docx-editor-core': minor
---

Add DOCX to Markdown converter at `@eigenpal/docx-editor-core/markdown`.

Four entry functions, each accepting either a parsed `Document` (sync) or raw bytes (`Buffer`, `Uint8Array`, `ArrayBuffer`, async via `parseDocx`):

- `toMarkdown` for continuous output.
- `toMarkdownPaged` for one markdown string per Word page, plus a `combined` string with `<!-- page N -->` separators. Page boundaries come from Word's pre-baked pagination hints (`renderedPageBreakBefore`, explicit page breaks, section breaks of type `nextPage`/`evenPage`/`oddPage`) so the call is sync and dep-free.
- `toMarkdownAsync` and `toMarkdownPagedAsync` wrap the above with an `imageHandler` callback that the converter awaits per image to substitute the default `![alt](./images/...)` reference (LLM-described alt text, uploaded URLs, data URLs, empty string to drop, etc.).

Options: `annotations: 'html' | 'pandoc' | 'strip'`, `trackedChanges: 'clean' | 'annotate'`, `comments: 'strip' | 'inline' | 'sidecar'`, `hyperlinks: 'inline' | 'reference'`, `footnotes: 'inline' | 'end'`, `headerFooter: 'strip' | 'first-page' | 'all'` (paged only), plus a custom `imagePath(info)` callback. Images come back in `result.images` as a `Map<virtualPath, ImageRef>` with raw bytes, base64, data URL, and metadata so callers decide what to do with them.

A live playground lives at `examples/markdown-playground` (`bun run --filter './examples/markdown-playground' dev`).
