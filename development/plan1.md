# Office MCP – Font & Paragraph Formatting Plan (Minimal)

This plan adds essential Word formatting controls to the Office MCP server without overloading scope.

## Scope (what we will support)
- Font styling (document/selection/matches):
  - Family (e.g., Arial)
  - Size (e.g., 12 pt)
  - Color (named colors and hex, default black)
  - Weight/style: bold, italic, underline (toggle)
- Paragraph formatting:
  - Alignment: left, center, right, justify
  - Spacing: line spacing (single, 1.5, double), space before/after (pt)
- Lists:
  - Bulleted or numbered list
  - Remove list formatting
- Styles (lightweight):
  - Apply built-in styles by name (e.g., "Normal", "Heading 1", "Heading 2")
  - Update the Normal style (for base font defaults)
- Clear formatting:
  - Clear direct formatting on scope (retain underlying style)

Non-goals (for now): theme management, custom style creation (except Code already exists), style gallery operations, advanced typography (kerning/ligatures), section/page layout, track changes.

## Proposed MCP tools (simple contracts)

1) set_document_base_font
- Input: name (str), size (number), color (str|hex), bold? (bool), italic? (bool), underline? (bool), path? (str)
- Behavior: Updates ActiveDocument.Styles("Normal").Font and reapplies to paragraphs with Normal where practical. If `path` provided, operate on that file.
- Output: short status string.

2) set_font
- Input: scope: "document" | "selection" | "matches"; name? (str), size? (number), color? (str|hex), bold? (bool), italic? (bool), underline? (bool), find? (str), caseSensitive? (bool), wholeWord? (bool), maxMatches? (int), path? (str)
- Behavior: Applies direct font formatting to the chosen scope. For matches, Find each range and format.
- Output: status string with applied count (where relevant).

3) set_paragraph_format
- Input: scope: "document" | "selection" | "matches"; alignment? (left|center|right|justify), lineSpacing? (single|onePointFive|double|exact:<pt>|atLeast:<pt>), spaceBefore? (pt), spaceAfter? (pt), find?/caseSensitive?/wholeWord?/maxMatches?, path?
- Behavior: Applies paragraph-level properties on scope.
- Output: status string with applied count.

4) set_list_style
- Input: scope: "selection" | "matches" | "paragraphs"; type: "bullet" | "numbered" | "none"; level? (int), find?/caseSensitive?/wholeWord?/maxMatches?, path?
- Behavior: Applies or removes list formatting via ListFormat.ApplyBulletDefault / ApplyNumberDefault / RemoveNumbers / RemoveBullets.
- Output: status string with applied count.

5) apply_style
- Input: scope: "selection" | "matches" | "paragraphs"; styleName (e.g., "Normal", "Heading 1"), find?/caseSensitive?/wholeWord?/maxMatches?, path?
- Behavior: Applies a style to the paragraph(s) intersecting scope.
- Output: status string with applied count.

6) clear_direct_formatting
- Input: scope: "document" | "selection" | "matches"; find?/caseSensitive?/wholeWord?/maxMatches?, path?
- Behavior: Clears direct character + paragraph formatting on scope (keeps style).
- Output: status string with applied count.

Notes:
- "matches" scopes operate on ranges returned by Find. For paragraph tools, apply to the Paragraphs containing each range.
- Keep argument sets small; all parameters optional except the scope/type where needed.

## Implementation outline

- word_editor.py (helpers):
  - parse_color(value) -> (RGB int or WdColor): accept named colors (map) and #RRGGBB.
  - set_range_font(rng, opts): set Font.Name, Size, Color, Bold, Italic, Underline.
  - set_range_paragraph(rng, opts): set Alignment, LineSpacingRule, LineSpacing, SpaceBefore/After.
  - iter_find_ranges(doc, query, flags) -> generator of non-overlapping ranges (forward only, ensure progress with start/end checks).
  - apply_to_scope(doc, scope, find, fn): resolves selection/document/matches and applies fn(range) safely.
  - update_normal_style(doc, font_opts): update Styles("Normal").Font and optionally touch paragraphs using Normal (best-effort).

- server.py (tools):
  - Mirror patterns used by highlight_all/search_and_replace:
    - path vs active doc
    - read-only/protection checks; _prepare_file_for_edit, _make_doc_editable
    - with_doc_lock for path; with_global_lock for active
    - run in worker + timeout cleanup for path
    - Save() when changes were made; AUTO_CLOSE honored
  - Consistent error strings: No doc open, multiple docs open, read-only, file not found.

- Data/Enums:
  - Color map similar to highlight map plus hex parsing; default black.
  - Alignment map: left, center, right, justify -> WdParagraphAlignment constants.
  - LineSpacing map: single, 1.5, double or exact/atLeast with pt.

## Milestones
1) Base font via Normal style (Arial, 12, black) + Save. (set_document_base_font)
2) Character formatting for document/selection/matches (set_font).
3) Paragraph alignment + spacing (set_paragraph_format).
4) Lists: apply/remove bullets/numbering (set_list_style).
5) Apply style and clear formatting (apply_style, clear_direct_formatting).
6) Documentation in README + brief examples.

## Acceptance criteria
- Can set Normal style to Arial/12/black and see edits inherit these defaults.
- Can make a selection bold/italic/underline and change color/size.
- Can align paragraphs center and set 1.5 line spacing with space after 6pt.
- Can convert selected paragraphs to bulleted list and then remove list formatting.
- Can apply "Heading 2" to matched lines by search.
- All tools return succinct status and work with both active doc and explicit path.

## Edge cases & safeguards
- Protected or read-only docs: return clear error (don’t modify).
- Multiple open docs without path: ask for path.
- Locale/style names: try Styles("Normal"/"Heading 1"); if missing, report error (future: BuiltInStyle IDs).
- Find loops: ensure forward progress (use end offsets and stop at doc end).
- Selection absent (comments edge case): select Range(0,0) before insert/apply when needed.

## Test approach (lightweight)
- Manual smoke tests across a sample .docx with headings, lists, and paragraphs.
- Repeatable scripts under tests/ that call tools with known files (best-effort due to COM UI dependency).
- Verify Save persists changes; verify no crash with empty/short docs.

## Open questions
- Do we need RGB support beyond named colors day-one? (Plan says yes via #RRGGBB.)
- Should list level control be included now or deferred? (Deferred unless trivial.)
- Should we reapply Normal style to all paragraphs when updating base font? (Best-effort only.)
