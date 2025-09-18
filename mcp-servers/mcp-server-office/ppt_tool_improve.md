# PowerPoint MCP Tooling Overhaul Plan (Full Replacement)

STRICT RULES:
- 1-based slide indexing. `slide_index` starts at 1. `ppt.presentation.create` always produces slide 1.
- No legacy code retained. No aliases, no fallback, no compatibility wrappers.
- No deterministic coordinate helper abstractions (no semantic corners, no auto-placement heuristics). Agent supplies explicit numeric coordinates (points, or mm/px with unit parsing) for shapes.
- Consistent naming: `ppt.<resource>[.<subresource>].<action>`.
- All responses follow: `{ ok: boolean, ...data | error }` with no version negotiation overhead.
- Errors standardized: `{ ok: false, error: { code, message, details? } }`.

## Current vs Proposed Consolidation

Existing (after recent changes) includes creation, slide ops, shape ops, save/load, reveal, listing. We streamline to a minimal, orthogonal set.

## Final Tool Inventory

### 1. Presentation Lifecycle
1. `ppt.presentation.create(a4_portrait: bool = true, close_existing: bool = false)`
   - Ensures slide 1 exists (blank). Returns: `presentation_index, width_pt, height_pt, slide_count`.
2. `ppt.presentation.list()`
   - Lists open presentations with: index, path (if saved), saved_state (saved/unsaved), slide_count, active: bool.
3. `ppt.presentation.activate(presentation_index: int)`
   - Makes target presentation active. Error if out of range.
4. `ppt.presentation.close(presentation_index: int | None = None, save: bool = false)`
   - Closes one or all (if None and exactly one; error if ambiguous). If `save=true` and unsaved -> error `unsaved` (force explicit save_as first).
5. `ppt.presentation.save(presentation_index: int | None = None)`
   - Quick-save already named. Error codes: `ambiguous`, `unsaved`.
6. `ppt.presentation.save_as(target_path: string, overwrite: bool = false, presentation_index: int | None = None, close_after: bool = false, reveal: bool = false)`
   - Persists or exports. Returns normalized path, size_bytes.
7. `ppt.presentation.reveal(presentation_index: int | None = None)`
   - Brings window front; if multiple and None -> `ambiguous`.

### 2. Slide Management
8. `ppt.slide.add(position: int | None = None, layout: string = "blank")`
   - Inserts new slide. `position` 1-based; None = append.
9. `ppt.slide.delete(slide_index: int)`
   - Removes slide. Error `not-found` if invalid.
10. `ppt.slide.list()`
    - Returns array of slides: `slide_index, layout, shape_count`.

### 3. Shape: Text Boxes
11. `ppt.shape.textbox.add(slide_index: int, left: number | string, top: number | string, width: number | string, height: number | string, text: string, font?: object, paragraph?: object, dpi: number = 96)`
12. `ppt.shape.textbox.update(slide_index: int, shape_id: int, text?: string, font?: object, paragraph?: object)`

### 4. Shape: Images
13. `ppt.shape.image.add(slide_index: int, path: string, left: number | string, top: number | string, width?: number | string, height?: number | string, preserve_aspect: bool = true, dpi: number = 96)`
14. `ppt.shape.delete(slide_index: int, shape_id: int)`
15. `ppt.shape.zorder(slide_index: int, shape_id: int, action: string)` where action ∈ { `bring_to_front`, `send_to_back`, `step_forward`, `step_backward` }.

### 5. Shape Introspection
16. `ppt.shape.list(slide_index: int)`
   - Returns: `shapes: [ { shape_id, type, left_pt, top_pt, width_pt, height_pt, z, has_text, text? } ]`.

### 6. Presentation Content Snapshot
17. `ppt.presentation.content()`
   - Full content dump (slides + shapes + text) for planning or serialization.

### 7. Capabilities / Meta
18. `ppt.capabilities()`
   - Returns static meta: supported layouts (e.g., `["blank","title"]`), coordinate_units (`["pt","mm","px"]`), font_features (booleans), max_slide_dimensions_pt, errors list (codes + brief purpose). No placement enums (per instruction: no deterministic coordinate helpers provided).

## Removed / Not Included
- Any semantic placement tools.
- Any legacy names (`ppt_add_text_box`, etc.).
- Any partial or deprecated aliasing.
- Old combined shape update commands (kept separate purposeful granularity).
- No multi-operation batch tool (keeps surface smaller and transparent).

## Parameter & Validation Rules
- `slide_index` always >= 1 and <= current `slide_count`.
- All coordinate inputs allow: raw number (interpreted as pt), or string with unit suffix (`mm`, `pt`, `px`). Reject inches for simplicity.
- Reject strings that are numeric without unit but cannot parse float -> `invalid-unit`.
- Font object fields (optional): `{ name?: string, size?: number, color?: string, bold?: bool, italic?: bool, underline?: bool, strikethrough?: bool, superscript?: bool, subscript?: bool }`.
- Paragraph object fields (optional): `{ alignment?: "left"|"center"|"right"|"justify", line_spacing?: number, space_before?: number, space_after?: number, bullets?: bool, left_indent?: number, first_line_indent?: number }`.
- Image placement: if width and height omitted -> error `missing-dimension`; if one provided and `preserve_aspect=true` -> compute other; if both provided and `preserve_aspect=true` -> maintain provided width, adjust height proportionally (documented), else accept as-is.
- Z-order action strict enum; invalid -> `invalid-action`.

## Response Shapes
Success (generic): `{ ok: true, ...data }`
Error: `{ ok: false, error: { code, message, details? } }`

Common data patterns:
- Add slide: `{ ok: true, slide_index, slide_count }`
- Add textbox/image: `{ ok: true, shape_id, left_pt, top_pt, width_pt, height_pt }`
- List shapes: `{ ok: true, shapes: [...], slide_index }`
- Capabilities: `{ ok: true, layouts, coordinate_units, font_features, max_slide_width_pt, max_slide_height_pt, error_codes }`

## Error Codes (Authoritative Set)
- `not-found` (slide / shape / presentation)
- `ambiguous` (missing index when multiple resources open)
- `unsaved` (save requested on unnamed presentation)
- `invalid-unit`
- `invalid-dimensions`
- `invalid-action`
- `out-of-range` (coordinates or slide_index beyond bounds)
- `creation-failed`
- `operation-failed`
- `overwrite-denied`
- `io-error` (filesystem issues on save_as)
- `validation-failed` (generic catch for structured validation layering)

## Consolidation Rationale
- Moved from verb-first (`ppt_add_text_box`) to resource-first hierarchical naming to improve agent inference.
- Merged shape deletion into a single `ppt.shape.delete` (covers text, image, future shapes) instead of separate per-type deletions.
- Maintained separate add/update for clarity; no patch-style multi-field sentinel logic needed.
- Kept `ppt.presentation.content` separate from `ppt.shape.list` because one is full hierarchical dump; the other is targeted introspection.
- Excluded batch operations to keep atomicity clear and reduce complexity.
- Retained explicit z-order tool; z-order is specialized enough to justify its own action.

## Implementation Phasing (Recommended)
1. Introduce new tool names alongside existing (hidden flag) internally.
2. Port logic; ensure identical validation semantics.
3. Replace exports (tool registration) with new set.
4. Delete legacy code (no wrappers) and update docs.
5. Regenerate any SDK / client examples.
6. Add capability tests: shape lifecycle, slide add/delete, save cycles, error triggers.

## Open Questions (Decide Before Coding)
- Keep `layout` enum limited (`blank`, `title`) or expand? (Recommendation: keep minimal until real demand.)
- Do we want a `ppt.presentation.duplicate(slide_index, position?)` in v2? (Deferred.)
- Include text extraction in `ppt.shape.list`? (Currently yes: only if `has_text` true include `text`.)

## Documentation Additions
- Root README PowerPoint section: update tool matrix table reflecting final names.
- Migration note (short): “Legacy tool names removed entirely—agents must rely on discovery.”
- Error code reference table.

## Non-Goals
- Coordinate heuristics / semantic placement.
- Layout auto-detection.
- Tool version negotiation.
- Undo/redo or history tracking.

## Acceptance Checklist
- [ ] All legacy tool registrations removed.
- [ ] New 18 tools registered with consistent decorators.
- [ ] `ppt.presentation.create` returns `slide_count >= 1`.
- [ ] All errors match defined codes only.
- [ ] Docs updated (README, implementation summary, quality assessment).
- [ ] Tests: creation, add slide, delete slide, add/update/delete text box, add image (aspect cases), z-order, list shapes, save/save_as, capabilities, error scenarios.

---
This plan is ready for implementation under a full overhaul approach (no backwards compatibility). Modify open questions, then proceed to execution.
