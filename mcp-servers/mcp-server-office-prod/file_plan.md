# Office MCP - File Operations & Expansion Plan

Date: 2025-09-16
Status: Draft (New File + Save As prioritized)

## Objectives
Provide a minimal, coherent, and extensible file operation surface for AI agents interacting with Office documents (Word first; later PowerPoint/Excel) that:
- Is deterministic, idempotent where appropriate, and safe (no silent overwrites)
- Separates concerns: file lifecycle vs document content vs formatting vs review
- Enables future expansion (templating, export, conversion, comparison, snapshots)
- Maintains internal organization: grouping by domain (word_, ppt_, excel_) and capability family (file, content, format, analysis)

## Current Gaps
Existing tools emphasize content/formatting actions but lack explicit primitives for:
- Creating a new document at a path from structured input
- Saving current doc to a new format / location with validation and collision policy
- Enumerating open docs / verifying active context
- Explicit persistence (auto-save pattern vs explicit 'save' semantics)

## Naming & Namespacing Principles
- Prefix strongly by application when behavior is app-specific: `word.*`, `powerpoint.*`, `excel.*`
- Use capability families as second component when grouping grows: `word.file.new`, `word.file.save_as` (internally can remain flat tool names but aliasing pattern documented)
- Keep outward tool names concise: `create_word_document`, `save_word_document_as` for initial release to match existing style (`open_word_document`, `close_word`)
- Internally stage future transition to hierarchical alias mapping if FastMCP adds support.

## Proposed Initial File Operation Tools
1. create_word_document
2. save_word_document_as
3. list_word_documents (optional phase 2 – inspection utility)
4. save_word_document (explicit save of active or by path, optional phase 2)
5. export_word_document (phase 2 – format-specific advanced conversion wrapper)

---
## 1. Tool: create_word_document
Creates a new .docx at the specified path (or returns an error) optionally seeding with markdown or plain text or a template profile.

Signature (proposed):
```
create_word_document(
  path: str,
  content: str = "",
  content_format: str = "markdown",  # markdown|plain|none
  template: str = "default",          # default|blank|minimal_heading|research|custom:<name>
  overwrite: bool = False,
  make_active: bool = True,
  visible: bool = False
) -> { "path": str, "created": bool, "activated": bool, "size_bytes": int, "notes": str } | { "error": str }
```

Behavior:
- Validates extension (.docx enforced at MVP)
- Fails if file exists unless overwrite=True
- Creates parent directories if missing (within allowed root / sandbox)
- Opens in Word (hidden unless visible=True) and seeds content
- If content_format == markdown uses existing `write_markdown_to_document`
- If template != default and content empty, seed from template factory
- Calls save explicitly and returns metadata

Validation / Errors:
- Path outside sandbox (future) → error
- Unsupported extension → error
- Directory creation failure → error
- Word COM startup failure → error
- Template not found (custom) → error

Templates (phaseable):
- blank: no content
- minimal_heading: "# Title\n\n" stub
- research: heading sections stub (Abstract, Introduction, Methods, Results, Discussion)
- custom:<name>: loads from `data/templates/<name>.md` (phase 2)

Atomicity:
- If any failure after file creation but before save, delete the partial file
- Use try/finally cleanup pattern

Concurrency:
- Acquire per-document lock via canonical path key
- No global lock unless make_active=True and ambiguous context risk identified

---
## 2. Tool: save_word_document_as
Saves an open document (by path or the sole active doc) to a new path/format.

Signature:
```
save_word_document_as(
  target_path: str,
  source_path: str = "",     # optional; if empty and only one document open use that
  format: str = "docx",        # docx|pdf|rtf|txt|html|md (md = server-side markdown export)
  overwrite: bool = False,
  include_comments: bool = True,   # if pdf/html export supports it
  optimize_for: str = "default",  # default|print|web (affects html/pdf nuances later)
  close_after: bool = False,
  reveal: bool = False
) -> { "saved": bool, "target_path": str, "bytes": int, "format": str, "notes": str } | { "error": str }
```

Format Handling:
- docx: native SaveAs
- pdf: Word constant `wdFormatPDF` (ensure extension corrected)
- rtf: `wdFormatRTF`
- txt: `wdFormatText` (UTF-8 enforced)
- html: `wdFormatFilteredHTML` (default) or full HTML later
- md: not native – export via `get_markdown_representation` and write out `.md` file directly (no round-trip back into Word)

Behavior Flow:
1. Resolve both paths (sandbox enforce later)
2. Identify source document: by canonical path match or fallback if exactly one open
3. Validate source not read-only for formats requiring SaveAs
4. If target exists and not overwrite → error
5. Map format string to Word format constant (except md)
6. Perform SaveAs (COM) or write markdown file
7. Optionally close the new target (if close_after True) leaving original open; if format is conversion (pdf) keep original loaded
8. Return metadata

Edge Cases:
- Saving over same path with different format → allowed if overwrite True
- Attempt markdown export on unsaved new document (no original path) → require explicit source_path or error
- Hidden Word instance: reveal=True brings to foreground after save

Error Conditions:
- Multiple open docs & no source_path
- Source not found or not open
- Unsupported format
- SaveAs COM failure with underlying HRESULT captured

---
## 3. Future File/Session Tools (Backlog Extract)
- list_word_documents() → enumerate `{ path, read_only, active, modified, window_visible }`
- save_word_document(path?="", all: bool=False)
- close_word_document(path, save=False)
- reopen_word_document(path) (refresh + markdown snapshot)
- duplicate_word_document(source_path, target_path, overwrite=False)
- snapshot_word_document(path, label) → version manifest entry
- diff_word_documents(path_a, path_b, mode="text|format|structure") (later, uses styled JSON)

---
## Internal Module Organization Changes
Current `server.py` is large ( >2000 LOC ). Introduce modular separation for clarity and testability:

Proposed structure:
```
mcp_server/
  server.py                # composition & registration only (thin)
  tools/
    __init__.py
    word_file.py           # create/save/save_as/list
    word_content.py        # open/get/edit/highlight/search
    word_format.py         # batch_format/font/paragraph/list/style/clear
    word_review.py         # comments/analyze
    ppt_file.py            # future
    ppt_content.py
    excel_file.py
    excel_content.py
```

Refactor Principles:
- Each tool module exposes `register(mcp)` returning None
- Shared helpers remain in `app_interaction/` or new `word_utils.py`
- Keep existing public tool names stable; optionally add aliases later
- Introduce a small `errors.py` with typed error dataclasses for clarity (optional)

Incremental Refactor Strategy:
1. Add new modules for file ops only (no churn to existing code) – Phase A
2. Migrate formatting tools next (batch_format + dependents) – Phase B
3. Migrate content/review tools – Phase C
4. Slim `server.py` to orchestrator – Phase D

Risk Mitigation:
- Maintain import compatibility (from mcp_server.server import create_mcp_server)
- Run existing tests after each phase
- Add new tests for file ops (creation, overwrite protection, format conversions using small fixture docs)

---
## Error & Response Conventions
Standard success envelope for structured tools returning more than a string:
```
{ "ok": true, "data": { ... }, "warnings": [ ...? ] }
```
Standard failure envelope:
```
{ "ok": false, "error": { "code": "validation|not_found|conflict|io|com_error|unsupported|state", "message": str, "detail"?: any } }
```
For backward compatibility initial implementation may directly return either dict-with-error or success dict; envelope upgrade is optional Phase 2.

---
## Security & Safety (File Ops Scope)
- Enforce root sandbox using env var `MCP_OFFICE_SANDBOX_ROOT` (Phase 2)
- Reject path traversal outside root (use Path.resolve() prefix check)
- Enforce allowed extensions for create (`.docx`) / export whitelist
- Size guard for markdown export (>5MB plain text) unless `force=True` (future param)

---
## Testing Strategy
- Unit tests: path validation, overwrite logic, format mapping, markdown export content
- Integration tests (Windows only, marked with skip if not win32) for COM SaveAs flows
- Use tiny synthetic document fixture generated on the fly (avoid large binaries in repo)
- Mock markdown export writing for md format path length edge cases

---
## Implementation Roadmap (Focused Extract)
Phase A (File Ops MVP)
1. Implement `word_file.py` with create & save_as tools
2. Register tools from `server.py` (no removals)
3. Add tests for create/save_as (non-COM heavy parts) + skip markers for Windows-only integration
4. Docs: Update README new section "File Operations"

Phase B (Refactor & Additional Utilities)
5. Introduce `list_word_documents` & `save_word_document`
6. Add sandbox enforcement + config doc
7. Start migrating batch_format into `word_format.py`

Phase C (Expansion)
8. Export enhancements (html options, optimize_for)
9. Markdown template directory & custom: scheme
10. Snapshot & diff scaffolding (blocked on styled JSON phases)

---
## Prioritized Backlog (Top 10)
1. create_word_document (MVP)
2. save_word_document_as (MVP)
3. list_word_documents
4. sandbox root enforcement
5. markdown export path (md format) + template directory
6. save_word_document (explicit persistence)
7. export_word_document (pdf/html advanced switches)
8. duplicate_word_document
9. snapshot_word_document
10. diff_word_documents (structure mode leveraging future styled JSON)

---
## Data Points to Collect (Telemetry Later)
- Count of create vs open vs save_as calls
- Average document size on creation
- Format distribution for save_as
- Error type frequencies (validation vs com_error)

---
## Risks & Mitigations
| Risk | Impact | Mitigation |
|------|--------|-----------|
| COM SaveAs silent failures | Data loss / confusion | Capture HRESULT & wrap in structured error |
| Path traversal exploit | Unauthorized filesystem write | Sandbox + normalization + prefix check |
| Large doc markdown export | Memory / token blow-up | Size guard + require force flag |
| Race on simultaneous create same path | Corruption | Per-doc lock + existence re-check after lock |
| Over-refactor causing merge conflicts | Slow velocity | Incremental module adoption Phase A only for new features |

---
## Minimal Pseudocode Sketch (create)
```python
def create_word_document(path, content="", content_format="markdown", template="default", overwrite=False, make_active=True, visible=False):
    p = resolve_user_path(path)
    if p.exists() and not overwrite: return {"error":"exists"}
    ensure_parent_dirs(p)
    word = get_word_app()
    doc = word.Documents.Add()  # start new
    if not make_active: word.ActiveWindow.Visible = False
    seed = derive_seed_content(content, content_format, template)
    if seed:
        write_markdown_to_document(doc, seed) if content_format=="markdown" else replace_document_content(doc, seed)
    doc.SaveAs(str(p))
    if not make_active and not visible: word.Visible=False
    size = p.stat().st_size
    return {"path": str(p), "created": True, "activated": True, "size_bytes": size, "notes": "created from template '..."}
```

---
## Open Questions
- Should create auto-return markdown snapshot? (Optional param `return_content=true` later)
- Should save_as optionally return converted content for md/txt? (Probably yes Phase 2) 
- Introduce uniform response envelope now vs later? (Lean: later) 

---
## Acceptance Criteria (MVP)
- Creating a new doc that exists without overwrite returns error message (no mutation)
- Creating with markdown content writes correctly formatted headings & lists
- Saving to pdf produces a file at path with size > 0
- Markdown export writes file whose content hash equals `get_markdown_representation` output
- No global lock contention added beyond per-document locks

---
## Summary
This plan introduces disciplined, minimal file lifecycle primitives (create + save_as) while establishing a scalable organization and naming approach. Further expansion (listing, sandbox enforcement, exports, snapshots, diffs) slots cleanly without large future refactors.
