# DOCX to Markdown playground

Drop a `.docx` file. The rendered Word document appears on the left, the converted markdown on the right. Every option drives the conversion live.

```bash
bun run --filter './examples/markdown-playground' dev
# http://localhost:5180
```

What it exercises:

- `toMarkdown` (continuous) and `toMarkdownPaged` (per-page) from `@eigenpal/docx-editor-core/markdown`
- All option dimensions: `annotations`, `trackedChanges`, `comments`, `hyperlinks`, `footnotes`, `headerFooter`
- Image registration (count shown in the top-right stats)
- Warnings panel for things markdown can't represent (merged cells, multi-column, etc.)

Useful sample files in the repo:

- `examples/shared/sample.docx`: minimal, clean output
- `screenshots/demo.docx`: rich formatting, footnotes, images, tables
- `screenshots/pr320/multi-section.docx`: multi-section, exercises paged mode
