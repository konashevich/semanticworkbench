# Office MCP — Essential Plan (MVP)

Date: 2025-09-17

Goal: Provide just the essential steps to add safe file lifecycle primitives for Word (create + save-as) so AI agents can reliably create and export documents.

Essential Steps (MVP):
- Implement `create_word_document(path, content, content_format='markdown', overwrite=False, make_active=True, visible=False)`
- Implement `save_word_document_as(target_path, source_path='', format='docx', overwrite=False, include_comments=True, close_after=False)`
- Add path validation and basic sandbox check (env: `MCP_OFFICE_SANDBOX_ROOT`) to prevent directory traversal
- Use per-document locks (existing `canonical_doc_key`) and atomic save patterns (undo-record or temp file + rename fallback)
- Export formats: minimal support `docx`, `pdf`, `md` (markdown export by using `get_markdown_representation`)
- Unit tests for path/overwrite logic and markdown export; mark COM-dependent tests as Windows-only integration tests

Quick Acceptance Criteria:
- Creating a file where one already exists without `overwrite` returns an error
- Creating with `content_format='markdown'` writes correctly formatted content
- `save_word_document_as(..., format='pdf')` produces a non-empty PDF file at `target_path`

Next Small Steps (after MVP):
- Add `list_word_documents()` utility
- Add `save_word_document(path)` explicit save helper
- Document the new tools in `README.md` under "File Operations"

That's the minimal, high-value plan to ship file creation and save-as safely. Tell me if you want me to implement the MVP tools now (I can scaffold `mcp_server/tools/word_file.py`, register tools, and add tests). 
## Summary
This plan introduces disciplined, minimal file lifecycle primitives (create + save_as) while establishing a scalable organization and naming approach. Further expansion (listing, sandbox enforcement, exports, snapshots, diffs) slots cleanly without large future refactors.
