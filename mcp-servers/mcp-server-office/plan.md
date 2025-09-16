# Office MCP Document Representation & Interaction Plan

Purpose: Enable AI agents to flexibly access, navigate, and modify Word document content with progressively richer formatting fidelity while minimizing noise, token overhead, and synchronization errors.

## Guiding Principles
1. Progressive disclosure: lightweight first, richer only on demand.
2. Stable identities for safe edits (optimistic concurrency).
3. High information density per token (avoid raw XML/HTML spam by default).
4. Opt-in fallbacks (raw OOXML) for edge diagnostics.
5. Stateless-friendly, but enable caching & incremental refresh.
6. Clear separation: content vs formatting vs structure vs analytics.

## Representation Tiers (Baseline)
- `markdown` (current) – minimal structural/text view.
- `styled_json` (new) – whole-document normalized semantic formatting (loss-minimized, noise reduced).
- `ooxml_raw` (optional) – exact WordprocessingML snapshot (never default).

---
## Phase 0 – Foundations & Decisions
**Goals**: Confirm scope, naming, extension points.
**Deliverables**:
- This plan (`plan.md`).
- Enumeration of tool names & parameter conventions.
**Acceptance**: Plan reviewed & frozen for Phase 1.

---
## Phase 1 – Block Identity & Metadata Layer
**Goals**: Provide stable identifiers for paragraphs, tables, headers/footers.
**Additions**:
- Extraction adds for each block:
  - `block_id` (synthetic: e.g., `b::<scope>::<section>::<seq>`)
  - `para_id` (if available from OOXML `w14:paraId`)
  - `seq_index` (0-based order inside scope)
  - `char_range` (start,end snapshot offsets in concatenated scope text)
  - `content_hash` (SHA1 of canonicalized text + style_name)
- Document version object: `{file_path, last_modified, size_bytes, version_hash}`.
**Tools**:
- `get_document_overview()` -> counts & version.
**Acceptance**:
- IDs persist across read-only operations.
- Hash mismatch detectable after edit.
**Risks**: paraId absence; mitigated with multi-key fallback.

---
## Phase 2 – Headings & Structural Index
**Goals**: Fast navigation of outline.
**Tools**:
- `list_headings()` -> ordered heading array `{block_id, level, text, path}`.
- `get_heading_tree()` (optional) -> nested structure.
**Acceptance**:
- Correct ordering vs visual doc order.
- Excludes TOC field-generated paragraphs unless `include_toc=true`.
**Risks**: Misclassification of custom styles; mitigation: configurable heading style map.

---
## Phase 3 – Styled JSON Full Export
**Goals**: Whole document in normalized JSON.
**Normalization Rules**:
- Merge adjacent runs with identical effective formatting.
- Resolve style inheritance (paragraph + character + defaults).
- Inline effective attributes only if non-default: bold, italic, underline, strike, superscript, subscript, font_family, font_size_pt, color_rgb, highlight_rgb, alignment, list marker, heading_level.
- Represent tables: hierarchy `table -> rows -> cells -> blocks` with cell spans.
- Represent headers & footers with `section_index`, `type` (primary|first|even).
**API**:
- `open_word_document(format=styled_json, include_headers_footers=true|false, include_tables=true|false, max_blocks?, cursor?)`.
**Acceptance**:
- Output size < 25% of raw OOXML token estimate (target metric baseline example doc).
- No duplicate adjacent formatting entries.
**Risks**: Performance on large docs; mitigation: pagination (cursor + block window).

---
## Phase 4 – Raw OOXML Opt-In
**Goals**: Diagnostic fallback.
**API**:
- `open_word_document(format=ooxml_raw, scope=body|all)`.
**Protections**:
- Size guard: reject if > threshold unless `force=true`.
**Acceptance**:
- Exact WordOpenXML match (hash compare) to internal retrieval.
**Risks**: Token blow-up; mitigated by gate + warning banner in response.

---
## Phase 5 – Partial Fetch / Drill-Down Tools
**Goals**: Avoid full export when only subsets needed.
**Tools**:
- `get_blocks(block_ids[])` -> styled JSON subset.
- `get_paragraph(block_id)` (alias convenience).
- `get_table(table_block_id)` -> table subtree.
- `get_header_footer(section_index, type)`.
**Acceptance**:
- Latency improvement vs full export for small requests.
**Risks**: Inconsistent pagination; unify through shared resolver.

---
## Phase 6 – Editing & Optimistic Concurrency
**Goals**: Safe formatting/content edits via IDs.
**Tools**:
- `apply_paragraph_format(block_id, expected_hash, changes)`.
- `replace_text(block_id, expected_hash, new_text)`.
- `bulk_edit(edits[])` batch variant.
**Behavior**:
- Reject with `conflict` if hash mismatch or block not found.
- Return new hash + minimal diff summary.
**Acceptance**:
- Edit does not corrupt adjacent formatting; verified by hash shift only for target block.
**Risks**: Word COM race; mitigation: lock per document during mutation.

---
## Phase 7 – Semantic Search & Lightweight RAG
**Goals**: Retrieval without sending all text.
**Pipeline**:
- Chunk = paragraph or merged paragraphs < token limit.
- Embed on first structured export or explicit `build_index()`.
- In-memory vector store (e.g., NumPy + cosine or optional FAISS) keyed by `block_id`.
**Tools**:
- `semantic_search(query, top_k=5, filters?)` returns `{block_id, score, snippet}`.
- `reindex_document()` to rebuild after edits (or incremental update per edit).
**Acceptance**:
- Queries return relevant headings/content in test cases.
**Risks**: Model drift; allow embedding model name in index metadata.

---
## Phase 8 – Annotated Markdown Mode (Optional)
**Goals**: Human + model-friendly middle representation.
**Format**:
```
[H1 size=16 color=#2F5496]Executive Summary[/]
1. Objectives
```
- Only non-default attributes included.
**API**: `open_word_document(format=annotated_markdown, ...)`.
**Acceptance**: Round-trip readability; size << styled_json.
**Risks**: Parser ambiguity; define strict tag grammar.

---
## Phase 9 – Analytics & Summaries
**Goals**: Derived insights to reduce raw data transfer.
**Tools**:
- `style_usage()` -> counts by style.
- `heading_outline_summary()` -> tree + word counts.
- `list_statistics()` -> max depth, item counts.
**Acceptance**: Correct counts vs ground truth sample docs.
**Risks**: Drift after edits; refresh on demand.

---
## Phase 10 – Performance & Pagination Enhancements
**Goals**: Scale to large documents (> 500 pages).
**Features**:
- Streaming cursors for `styled_json` (windowed `block_start`, `limit`).
- Background prefetch / cache warming.
- Token size estimator (`estimate_tokens(format, options)` tool).
**Acceptance**: Large doc operations under target latency thresholds.
**Risks**: Memory growth; implement LRU eviction.

---
## Phase 11 – Advanced Features (Deferred)
- Revisions mode (include / exclude / diff summarization).
- Field code parsing (TOC, REF, PAGE) annotation.
- Inline comment extraction & resolution tools.
- Layout metrics (page number, column position) if needed.

---
## Phase 12 – Telemetry & Quality
**Metrics**:
- Avg blocks per export, styled_json size vs markdown ratio.
- Semantic search latency.
- Edit conflict rate.
- Cache hit ratio (identity + embeddings).
**Logging**: Structured (operation, duration_ms, doc_version).
**Acceptance**: Dash or log queries produce actionable stats.

---
## Phase 13 – Security & Safety
- Path allowlist / sandbox for opening docs.
- Size limits & guardrails for OOXML and raw exports.
- Redaction hook (optional) to mask sensitive patterns before returning.
- Avoid base64 docx to model unless explicitly requested.

---
## Phase 14 – Rollout Strategy
1. Ship Phase 1–2 (read-only nav) behind feature flag.
2. Gather usage metrics (which tools called most).
3. Release Phase 3 export + Phase 5 partial fetch.
4. Add edits (Phase 6) after stability burn-in.
5. Introduce semantic search (Phase 7) once content flows proven stable.
6. Iterate optional modes (Annotated, Analytics) based on demand.

---
## Data Structures (Sketch)
Block (common):
```
{
  "block_id": "b::body::1::37",
  "para_id": "00AF12BC",        // optional
  "seq_index": 37,
  "type": "paragraph" | "list_item" | "table" | "table_row" | "table_cell" | "header" | "footer",
  "text": "...",                 // paragraphs/list items
  "heading_level": 1,             // if heading
  "list": {"kind":"number","level":0,"marker":"1."},
  "format": {"bold":true,"font_size_pt":16,"color_rgb":"#2F5496"},
  "content_hash": "sha1:..."
}
```

Document version:
```
{"file_path":"C:/.../doc.docx","last_modified":"2025-09-11T10:20:30Z","size_bytes":123456,"version_hash":"sha1:..."}
```

---
## Conflict Handling Flow (Example)
1. Client fetches block with `content_hash=H1`.
2. Client sends edit referencing `block_id`, `expected_hash=H1`.
3. Server recomputes hash → H1? apply & return new hash H2; else conflict.
4. On conflict: client re-fetches block, re-applies intent.

---
## Minimal Initial Tool List (after Phase 3)
- `open_word_document` (formats: markdown|styled_json|ooxml_raw)
- `get_document_overview`
- `list_headings`
- `get_blocks`
- (Phase 6) `replace_text`, `apply_paragraph_format`
- (Phase 7) `semantic_search`

---
## Open Questions (Track & Resolve)
1. Heading style map customization source (config file? tool param?).
2. Embedding model selection & pluggability.
3. Field code inclusion policy.
4. Revisions default mode (exclude vs collapsed summary?).
5. Memory limit thresholds for cache eviction.

---
## Success Metrics
- 80% of model formatting queries answerable with `styled_json` subset or derived tools without OOXML dump.
- Styled JSON median size < 30% of OOXML tokens across sample corpus.
- <5% edit requests result in conflict retries under normal collaborative usage.
- Semantic search p95 latency < 150 ms (medium docs ~2k blocks, local embeddings cached).

---
## Deferred / Nice-to-Have Later
- Page layout extraction (coordinates, page numbers).
- Image OCR integration.
- Style diffing between document versions.
- Cross-document semantic linking.

---
## Implementation Order Summary
0. Plan
1. Identity + overview
2. Headings index
3. Styled JSON export + pagination
4. OOXML opt-in
5. Partial fetch tools
6. Editing (text + formatting)
7. Semantic search index
8. Annotated markdown (optional)
9. Analytics summaries
10. Performance/pagination tuning
11. Advanced (revisions, fields, comments)
12. Telemetry & security hardening
13. Rollout & iteration

---
## Notes
- Keep backward compatibility: default format remains markdown until stable.
- Avoid premature micro-optimizations; measure after Phase 3.
- Provide clear user-facing warnings when returning raw OOXML.
# Office MCP - Formatting and Tool Surface Plan - IMPLEMENTED

✅ **IMPLEMENTATION COMPLETE** - This plan has been successfully implemented with enhancements.

## Implementation Summary

### ✅ 1) Enhanced `word.batch_format` Tool - COMPLETED

**Implemented Features:**
- ✅ Atomic behavior with rollback on failure using temporary document saves
- ✅ Performance limits: Maximum 50 operations per call, 10,000 matches processed
- ✅ Comprehensive parameter validation:
  - Font sizes: 1-1638 points
  - Paragraph spacing: 0-1584 points  
  - Alignment: left, center, right, justify
  - List types: bullet, numbered, none
  - Clear formatting targets: all, font, paragraph
- ✅ Enhanced color format validation with detailed error messages
- ✅ Superscript/subscript mutual exclusion with clear error messages
- ✅ Selection scope validation (errors when no selection exists)
- ✅ Style existence validation before execution

**Contract Implemented:**
```json
{
  "scope": "document|selection|matches",
  "operations": [
    {
      "font": { "name?", "size?", "color?", "bold?", "italic?", "underline?", "strikethrough?", "superscript?", "subscript?" },
      "paragraph": { "alignment?", "line_spacing_rule?", "space_before?", "space_after?" },
      "list": { "type": "bullet|numbered|none" },
      "style": { "name" },
      "clear_formatting": { "target?": "all|font|paragraph" },
      "document_base_font": { "name?", "size?", "color?", "bold?", "italic?", "underline?" }
    }
  ],
  "find": { "text", "case_sensitive?", "whole_word?", "max_matches?" },
  "path?": "string"
}
```

**Output:**
```json
{
  "changed": boolean,
  "summary": { "operations": number, "ranges_affected": number },
  "details": [{ "op_index", "op_type", "ranges_affected" }]
}
```

### ✅ 2) Enhanced `word.highlight` Tool - COMPLETED  

**Implemented Features:**
- ✅ Unified API replacing separate highlight_text/highlight_all tools
- ✅ Enhanced color validation with comprehensive error messages
- ✅ Performance limits: Maximum 10,000 matches for performance reasons
- ✅ Detailed documentation of supported highlight colors
- ✅ Enhanced parameter validation with specific error messages

**Contract Implemented:**
```json
{
  "text": "string (required)",
  "scope": "first|all (default: first)",
  "color": "yellow|red|blue|green|pink|turquoise|bright_green|gray|...|0-16",
  "case_sensitive": "boolean",
  "whole_word": "boolean", 
  "max_matches": "number (max 10,000)",
  "path": "string (optional)"
}
```

**Output:**
```json
{
  "matches": number,
  "color": number,
  "scope": "first|all"
}
```

### ✅ 3) Comprehensive Testing Suite - COMPLETED

**Implemented Tests:**
- ✅ Operation limit validation (50 operations max)
- ✅ Font size validation (1-1638 points)  
- ✅ Paragraph spacing validation (0-1584 points)
- ✅ Alignment validation (left, center, right, justify)
- ✅ List type validation (bullet, numbered, none)
- ✅ Superscript/subscript mutual exclusion
- ✅ Performance limits (10,000 matches max)
- ✅ Scope validation for all tools
- ✅ Color validation for highlight tool
- ✅ Empty operations handling
- ✅ Find parameter format validation

### ✅ 4) Enhanced Error Handling - COMPLETED

**Implemented Error Categories:**
- ✅ **Validation Errors**: Parameter validation with specific ranges and allowed values
- ✅ **Performance Errors**: Clear limits with explanatory messages
- ✅ **Execution Errors**: Document access, style existence, selection requirements
- ✅ **Atomic Rollback**: Document restoration on operation failure

### ✅ 5) Tool Taxonomy - COMPLETED

The clean, coherent tool surface is implemented:

- **Content editing**: `word.edit_markdown` (existing)
- **Formatting**: `word.batch_format` (enhanced with atomic behavior)
- **Review & quality**: `word.highlight` (unified API)
- **File/session**: `word.open`, `word.save`, `word.close` (existing)

## Key Enhancements Beyond Original Plan

1. **True Atomic Behavior**: Uses temporary document saves for rollback capability
2. **Performance Monitoring**: Built-in counters to prevent resource exhaustion  
3. **Comprehensive Validation**: All parameter ranges validated with specific error messages
4. **Enhanced Documentation**: Detailed function docs with supported values and examples
5. **Robust Testing**: Complete test suite covering all validation scenarios

## Migration Status

- ✅ **No Breaking Changes**: All existing tool calls continue to work
- ✅ **Enhanced Functionality**: Existing tools now have better validation and error handling
- ✅ **Backward Compatible**: No legacy tool wrappers needed

The implementation successfully addresses all identified quality issues while maintaining full backward compatibility and adding significant reliability improvements for AI agent workflows.

---

# Original Plan (for reference)

This plan defines: (1) a new unified formatting tool `word.batch_format`, (2) a clean, holistic tool taxonomy, and (3) merging `word.highlight_text` and `word.highlight_all` into a single `word.highlight`.

Guiding constraints from the owner:
- Avoid thin wrappers for backward compatibility unless concretely needed.
- Justify any compatibility step with real, observable inconsistency tied to the current host config (see `mcp.json` excerpt); avoid speculative "what ifs".

The current MCP host JSON only starts the server (command/args/cwd/env). It does not whitelist tool IDs or pin schemas. Therefore, renaming/adding/removing tools does not break this JSON. Any compatibility step below is only included when we can point to a concrete incompatibility.

## 1) New unified tool: `word.batch_format`

Purpose: Apply one or more deterministic formatting operations in a single call, in a defined order, over a chosen scope (document, selection, or text matches).

Contract (logic-level):
- Name: `word.batch_format`
- Inputs:
  - `scope`: "document" | "selection" | "matches"
  - `find` (required iff scope=="matches"): { `text` (string), `case_sensitive` (bool, default false), `whole_word` (bool, default false), `max_matches` (int, 0 = no limit) }
  - `operations`: array of operation objects, executed in order. Supported op types:
  - `font`: { name?, size?, color?, bold?, italic?, underline?, strikethrough?, superscript?, subscript? }
    - `paragraph`: { alignment?, line_spacing_rule?, line_spacing?, space_before?, space_after? }
    - `list`: { type: "bullet" | "numbered" | "none" }
    - `style`: { name: string }
    - `clear_formatting`: { target?: "all" | "font" | "paragraph" (default "all") }
    - `document_base_font`: { name?, size?, color?, bold?, italic?, underline? }
  - `path` (optional): explicit .docx to target; if omitted, applies to the active document managed by the server.
- Output:
  - `{ changed: boolean, summary: { operations: number, ranges_affected: number }, details: [ { op_index, op_type, ranges_affected } ] }`
- Determinism & safety:
  - Single-document lock; operations run sequentially in provided order.
  - Direct COM calls for formatting (no LLM, no heuristics).
  - Save on change; respect `MCP_WORD_AUTO_CLOSE` for lifecycle.
- Error modes:
  - Invalid scope or missing `find` when scope=="matches" -> validation error.
  - Unknown style name -> error with style name echoed.
  - Color parsing errors for invalid values -> error with provided color.
  - Zero matches (scope=="matches") -> success with `changed=false` and `ranges_affected=0`.
  - Superscript/Subscript mutual exclusion -> if both are `true` in a single `font` op or across combined ops for the same range, return a validation error indicating they are mutually exclusive.

Examples (payload shapes):
- Bold and 14pt for all matches of "Cataract":
  - `scope: "matches"`, `find: { text: "Cataract", whole_word: true }`, `operations: [ { font: { bold: true, size: 14 } } ]`
- Convert current selection to a bullet list and add 12pt spacing after paragraphs:
  - `scope: "selection"`, `operations: [ { list: { type: "bullet" } }, { paragraph: { space_after: 12 } } ]`
- Set document base font to Arial 12 black and clear direct formatting everywhere:
  - `scope: "document"`, `operations: [ { document_base_font: { name: "Arial", size: 12, color: "black" } }, { clear_formatting: {} } ]`
- Apply superscript to all matches of "TM":
  - `scope: "matches"`, `find: { text: "TM", whole_word: true }`, `operations: [ { font: { superscript: true } } ]`
- Strike through the current selection:
  - `scope: "selection"`, `operations: [ { font: { strikethrough: true } } ]`

Acceptance criteria:
- Executes multiple ops in-order deterministically across all scopes.
- Returns stable counts; repeated identical calls are idempotent (no further changes, `changed=false`).
- Works with both active doc and explicit `path`.
- Superscript and subscript are mutually exclusive; applying one clears the other.

## 2) Merge highlight tools into `word.highlight`

Decision: Replace `word.highlight_text` and `word.highlight_all` with a single `word.highlight`.

Contract:
- Name: `word.highlight`
- Inputs: `{ text: string, scope: "first" | "all" (default "first"), color?: string|int (default "yellow"), case_sensitive?: bool, whole_word?: bool, max_matches?: int (default 0 for all) }`
- Output: `{ matches: number, color: string|int }`
- Behavior: Deterministic forward search; applies Word highlight color to the first match or all matches.

Compatibility stance (non-speculative):
- Given the provided `mcp.json` (server process only), there is no tool allowlist or schema pinning in config. Hosts will rediscover the tool list on connect. Therefore, removing the old highlight tool IDs will not break this JSON or server startup.
- Action: Do not add thin wrappers. Remove old IDs and document the new `word.highlight` name in README and examples.

Acceptance criteria:
- `word.highlight` replaces both prior behaviors with a single `scope` switch.
- Returns exact count of highlights applied.

## 3) Holistic tool suite (clean taxonomy)

Goal: A small, coherent surface that is easy for agents to route without overlapping functions.

- Content editing
  - `word.edit_markdown`: Structural/content edits via Markdown round-trip (no formatting guarantees).
- Formatting
  - `word.batch_format`: All deterministic font/paragraph/list/style/base-font/clear ops.
- Review & quality
  - `word.highlight`: Deterministic highlighting (merged API).
  - `word.add_comment`, `word.list_comments`, `word.delete_comment` (if present in codebase; otherwise future scope).
- File/session & diagnostics
  - `word.open`, `word.save`, `word.close` (as implemented), `word.list_documents` (if present), `word.get_status`/health (if present).

Rationalization:
- Formatting is centralized in `word.batch_format` to avoid a proliferation of narrowly scoped tools with overlapping parameters.
- Content edits remain in `word.edit_markdown` to preserve existing flow for structural changes.
- Highlight is a single tool to avoid duplicate entry points (first vs all).

## 4) Migration and compatibility

Current reality:
- The owner’s `mcp.json` only controls process startup. There is no evidence of tool allowlists or pinned schemas in config.
- Therefore, introducing `word.batch_format` and replacing the two highlight tools with `word.highlight` should not require wrappers to maintain host startup or discovery.

Plan (no thin wrappers unless a concrete inconsistency is observed):
1. Add `word.batch_format` alongside existing formatting tools for one release window.
   - Purpose: Give agents a new, better entry point without removing any existing functionality immediately.
   - This is not a wrapper; it is a new tool.
2. Switch all README examples and sample scripts to `word.batch_format` and `word.highlight`.
3. Observe tool call logs for N days to detect any calls to removed/legacy IDs after docs update.
   - Concrete signal needed to keep old IDs: If we see external clients still invoking removed names, that is an actual inconsistency.
4. Remove legacy tools in the next minor/major bump after confirming no legacy calls are observed.
   - No wrappers added; straight removal consistent with the current host config.

If a concrete inconsistency appears (examples of non-speculative triggers):
- A client error log shows "Tool not found: word.highlight_text" after the merge.
- An external host that we control includes an allowlist referencing old IDs (visible JSON or code in-repo).

Only in those cases, add a time-limited shim:
- A minimal redirect tool ID that maps parameters 1:1 to the new tool, with a deprecation notice in the server’s tool description. Remove after the client’s config is updated. This is a targeted, evidence-based exception.

## 5) Testing

Add integration tests (Windows-only where COM required):
- `batch_format` happy paths per scope, including combined font+paragraph+list.
- Idempotence check: apply twice -> second run `changed=false`.
- `highlight` first vs all; whole_word and case_sensitive toggles; count verification.
- Error cases: missing `find` for matches, bad style name, invalid color.
- Font flags: superscript, subscript, and strikethrough toggles; verify mutual exclusion (attempting both yields a validation error).

Smoke scripts:
- Update `examples/python/smoke_word_formatting.py` to use `word.batch_format` and `word.highlight`.

## 6) Documentation

- Update `mcp-servers/mcp-server-office/README.md`:
  - New `word.batch_format` spec with examples.
  - `word.highlight` unified API.
  - Mark prior narrow formatting tools as legacy (not deprecated wrappers), and note removal timeline once confirmed unused.

## 7) Delivery checklist

- [ ] Implement `word.batch_format`
- [ ] Implement unified `word.highlight` and remove `highlight_text`/`highlight_all`
- [ ] Update README and examples to use the new tools
- [ ] Add integration tests and run locally on Windows (Python 3.12 + pywin32)
- [ ] Observe logs for legacy tool invocations for N days
- [ ] Remove legacy formatting tools once confirmed unused

Notes on non-speculative compatibility:
- With the provided `mcp.json`, startup compatibility depends only on `command`, `args`, `cwd`, and `env`. Tool renames/additions do not affect startup or discovery here. Therefore, wrappers are not part of this plan unless logs or configs present a concrete mismatch (e.g., a client calling an old tool ID and failing). 
