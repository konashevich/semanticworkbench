# Word MCP Improvements (feat/office-edits)

This document tracks improvements implemented per plan.md, without touching README.

## New unified tools

- word.batch_format
  - Purpose: Apply multiple deterministic formatting ops in-order over a scope.
  - Inputs: scope, operations[], optional find/case/whole_word/max_matches, optional path.
  - Ops: font (name, size, color, bold, italic, underline, strikethrough, superscript, subscript), paragraph (alignment, line_spacing, space_before, space_after), list (bullet|numbered|none), style(name), clear_formatting, document_base_font.
  - Output: { changed, summary { operations, ranges_affected }, details[] }.
  - Notes: Validates superscript/subscript mutual exclusion; saves on change; respects auto-close.

- word.highlight (merged)
  - Replaces: word.highlight_text and word.highlight_all.
  - Inputs: text, scope (first|all), color, case_sensitive, whole_word, max_matches, path.
  - Output: "Highlighted" (first) or "Highlighted N occurrence(s)" (all), or error/no-match messages.
  - Behavior: Deterministic forward search; color supports names or indices.

## Updated capabilities

- Font helpers now support: strikethrough, superscript, subscript (mutually exclusive between super/sub).

## Migration stance

- No thin wrappers were added. The provided mcp.json only starts the server, so tool list changes don’t break startup or discovery.
- If we observe a concrete error like "Tool not found: word.highlight_text" from an external caller, we will add a time-limited shim and document it here.

## Testing notes

- Add integration tests for batch_format (scopes, idempotence, mixed ops) and highlight (first vs all) plus super/sub/strike flags.
- Smoke: update example(s) to invoke batch_format and highlight.

## Operational notes

- COM safety and timeouts unchanged (uses worker lock + cleanup on cancel).
- Color parsing and style application reuse existing helpers.
