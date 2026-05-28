---
'@eigenpal/docx-editor-core': minor
'@eigenpal/docx-editor-react': minor
'@eigenpal/docx-editor-vue': minor
'@eigenpal/docx-editor-i18n': minor
'@eigenpal/docx-editor-agents': minor
'@eigenpal/nuxt-docx-editor': minor
---

Track paragraph splits and joins as OOXML revisions in suggesting mode (fixes #614).

Pressing Enter in suggesting mode now marks the new paragraph break as tracked (`<w:pPr><w:rPr><w:ins/>`), and Backspace at the start of a paragraph marks the prior paragraph break as deleted (`<w:del/>`) without actually joining until accepted. These markers round-trip losslessly through DOCX.

New commands `acceptChangeById(revisionId)` and `rejectChangeById(revisionId)` resolve any revision by its `w:id`, including paragraph-mark revisions and inline insertion/deletion marks, in a single transaction. The painter shows a margin change bar and a pilcrow at the end of affected paragraphs.

Adds the `revisions.*` i18n namespace with 15 keys for upcoming review-sidebar entries (Phase 1 stubs; further phases land table-row, cell, and property revisions).
