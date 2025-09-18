# Copyright (c) Microsoft. All rights reserved.

from mcp.server.fastmcp import Context, FastMCP
import asyncio
from pathlib import Path
import os
import stat

from mcp_server import settings
from mcp_server.app_interaction.excel_editor import get_active_workbook, get_excel_app, get_workbook_content
from mcp_server.app_interaction.powerpoint_editor import (
    add_text_to_slide,
    get_active_presentation,
    get_powerpoint_app,
    get_presentation_content,
)
from mcp_server.app_interaction.word_editor import (
    get_active_document,
    get_markdown_representation,
    get_word_app,
    parse_font_color,
    write_markdown_to_document,
    replace_document_content,
    set_range_font,
    set_range_paragraph,
    apply_list_format,
    apply_style_to_range,
    update_normal_style,
    apply_to_scope,
    apply_to_scope_strict,
    clear_direct_formatting as clear_direct_fmt,
)
from mcp_server.markdown_edit.comment_analysis import run_comment_analysis
from mcp_server.markdown_edit.feedback_step import run_feedback_step
from mcp_server.markdown_edit.markdown_edit import run_markdown_edit
from mcp_server.types import MarkdownEditRequest
from mcp_server.concurrency import with_global_lock, with_doc_lock, run_in_word_worker, canonical_doc_key

server_name = "Office MCP Server"

# Word Highlight color map (WdColorIndex). Includes the standard palette + 0 (no highlight)
HIGHLIGHT_COLOR_MAP = {
    # clearing
    "none": 0,
    "no_color": 0,
    "clear": 0,
    "0": 0,
    # palette
    "black": 1,
    "blue": 2,
    "turquoise": 3,
    "bright_green": 4,
    "pink": 5,
    "red": 6,
    "yellow": 7,
    "white": 8,
    "dark_blue": 9,
    "teal": 10,
    "green": 11,
    "violet": 12,
    "dark_red": 13,
    "dark_yellow": 14,
    "gray50": 15,
    "grey50": 15,
    "dark_gray": 15,
    "dark_grey": 15,
    "gray25": 16,
    "grey25": 16,
    "light_gray": 16,
    "light_grey": 16,
    # numeric strings
    "1": 1,
    "2": 2,
    "3": 3,
    "4": 4,
    "5": 5,
    "6": 6,
    "7": 7,
    "8": 8,
    "9": 9,
    "10": 10,
    "11": 11,
    "12": 12,
    "13": 13,
    "14": 14,
    "15": 15,
    "16": 16,
}

def _parse_highlight_color(color: str | int | None) -> int:
    if color is None:
        return 7  # default yellow
    if isinstance(color, int):
        return int(color)
    key = str(color).strip().lower()
    return HIGHLIGHT_COLOR_MAP.get(key, 7)


# Helpers to ensure documents open editable
def _prepare_file_for_edit(p: Path) -> None:
    """Best-effort: ensure the file is writable and not marked as Internet zone.

    - Clears NTFS Zone.Identifier ADS to avoid Protected View on Windows.
    - Clears read-only file attribute.
    """
    try:
        if os.name == "nt":
            # Remove Mark-of-the-Web alternate data stream if present
            try:
                os.remove(str(p) + ":Zone.Identifier")
            except Exception:
                pass
        try:
            mode = os.stat(p).st_mode
            if mode & stat.S_IWRITE == 0:
                os.chmod(p, mode | stat.S_IWRITE)
        except Exception:
            pass
    except Exception:
        pass


def _make_doc_editable(doc) -> None:
    """Best-effort: disable read-only recommendations and protection on a Word doc."""
    try:
        try:
            doc.ReadOnlyRecommended = False
        except Exception:
            pass
        try:
            # Clear "Mark as Final"
            doc.Final = False
        except Exception:
            pass
        try:
            prot = getattr(doc, "ProtectionType", -1)
            if prot not in (-1, 0):
                try:
                    doc.Unprotect(Password="")
                except Exception:
                    try:
                        doc.Unprotect("")
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass


def _paths_equal(p: Path, doc_fullname: str) -> bool:
    """Robust path comparison between a filesystem Path and Word's Document.FullName.

    Handles case differences and separator normalization on Windows.
    """
    try:
        return canonical_doc_key(str(p)) == canonical_doc_key(doc_fullname)
    except Exception:
        return str(p) == doc_fullname


AUTO_CLOSE = os.getenv("MCP_WORD_AUTO_CLOSE", "0").strip().lower() in ("1", "true", "yes")
DEFAULT_TIMEOUT = int(os.getenv("MCP_WORD_OP_TIMEOUT", "75"))


async def _run_on_worker_with_cleanup(p: Path, fn):
    """Run fn on the worker for path p with timeout and cleanup on CancelledError.

    If the asyncio task is cancelled, best-effort close any open document matching p
    in the worker instance to avoid leaving lock files or duplicate windows.
    """
    key = canonical_doc_key(str(p))
    try:
        return await asyncio.wait_for(run_in_word_worker(key, fn), timeout=DEFAULT_TIMEOUT)
    except asyncio.CancelledError:
        # Best-effort cleanup: close any open doc matching path on the worker
        def _cleanup(word):
            try:
                for i in range(1, max(1, getattr(word.Documents, "Count", 0)) + 1):
                    try:
                        d = word.Documents(i)
                        if _paths_equal(p, d.FullName):
                            try:
                                d.Close(SaveChanges=0)
                            except Exception:
                                pass
                    except Exception:
                        continue
            except Exception:
                pass
            return None

        try:
            await run_in_word_worker(key, _cleanup)
        except Exception:
            pass
        raise


def create_mcp_server() -> FastMCP:
    mcp = FastMCP(name=server_name, log_level=settings.log_level, host="127.0.0.1")

    # region: helpers (sandbox)
    def _sandbox_ok(p: Path) -> tuple[bool, str]:
        root = os.getenv("MCP_OFFICE_SANDBOX_ROOT", "").strip()
        if not root:
            return True, ""
        try:
            r = Path(os.path.expanduser(os.path.expandvars(root))).resolve()
            pp = p.resolve()
            try:
                # Python 3.11+
                if pp.is_relative_to(r):
                    return True, ""
                return False, f"Path {pp} is outside sandbox root {r}"
            except Exception:
                # Fallback for older Path
                if str(pp).lower().startswith(str(r).lower() + os.sep) or str(pp).lower() == str(r).lower():
                    return True, ""
                return False, f"Path {pp} is outside sandbox root {r}"
        except Exception as e:
            return False, f"Sandbox root invalid: {e}"
    # endregion

    @mcp.tool()
    async def open_word_document(path: str) -> str:
        """Opens the specified .docx file in Word and returns its markdown representation.

        Args:
            path: Absolute or workspace-relative path to a .docx file.
        Returns:
            Markdown content of the opened document or an error message.
        """
        try:
            from mcp_server.path_utils import resolve_user_path
            p = resolve_user_path(path)
            if not p.is_file():
                return f"ERROR: File not found: {p}"
            if p.suffix.lower() != ".docx":
                return f"ERROR: Only .docx files supported. Got: {p.suffix}"
            from mcp_server.app_interaction.word_editor import get_word_app, get_markdown_representation

            async def _task():
                # Execute on a dedicated Word worker for this path
                def _open_on_worker(word):
                    # When using worker pool, we receive the worker's Word instance
                    if word is None:
                        from mcp_server.app_interaction.word_editor import get_word_app
                        word = get_word_app()
                    _prepare_file_for_edit(p)
                    # If already open, try to find matching document
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                try:
                                    # If previously opened read-only, close and reopen writable
                                    if getattr(d, "ReadOnly", False):
                                        d.Close(SaveChanges=0)
                                        d = word.Documents.Open(str(p), ReadOnly=False)
                                    _make_doc_editable(d)
                                except Exception:
                                    pass
                                return get_markdown_representation(d)
                        except Exception:
                            continue
                    # Open explicitly in editable mode
                    try:
                        doc = word.Documents.Open(str(p), ReadOnly=False)
                    except Exception:
                        # Fallback to simple open
                        doc = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        try:
                            _make_doc_editable(doc)
                        except Exception:
                            pass
                        result = get_markdown_representation(doc)
                        # Optionally close immediately to release any file locks
                        if AUTO_CLOSE:
                            try:
                                doc.Close(SaveChanges=0)
                            except Exception:
                                pass
                        return result
                    finally:
                        # Keep the document open for future operations
                        pass

                # Prevent indefinite hangs by bounding the open duration and ensure cleanup on cancel
                return await _run_on_worker_with_cleanup(p, _open_on_worker)

            # Per-document serialization to allow safe parallelism across different files
            return await with_doc_lock(canonical_doc_key(str(p)), _task)
        except Exception as e:
            return f"ERROR: Failed to open document: {e}"

    @mcp.tool()
    async def create_word_document(
        path: str,
        content: str = "",
        content_format: str = "markdown",
        overwrite: bool = False,
        make_active: bool = True,
        visible: bool = False,
        template: str = "default",
    ) -> dict | str:
        """Create a new .docx at `path` (optionally seeded with content).

        Args:
            path: Target .docx path.
            content: Initial content text.
            content_format: 'markdown' | 'plain' | 'none'.
            overwrite: Allow replacing existing file.
            make_active: Keep document open/active after creation.
            visible: Show Word window.
            template: Currently only 'default' supported (others ignored).
        Returns:
            Success dict with metadata or error string prefixed with 'ERROR:'.
        """
        from mcp_server.path_utils import resolve_user_path
        try:
            p = resolve_user_path(path)
            if p.suffix.lower() != ".docx":
                # Auto-append .docx if user omitted extension
                p = p.with_suffix(".docx")
            ok, msg = _sandbox_ok(p)
            if not ok:
                return f"ERROR: {msg}"
            if p.exists() and not overwrite:
                return f"ERROR: File already exists: {p}"
            try:
                p.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return f"ERROR: Failed to create parent directory: {e}"

            cf = (content_format or "markdown").strip().lower()
            allowed_cf = {"markdown", "plain", "text", "none"}
            if cf not in allowed_cf:
                return f"ERROR: Unsupported content_format: {cf} (allowed: markdown|plain|none)"
            tmpl = (template or "default").strip().lower()
            allowed_templates = {"default", "blank", "minimal_heading", "research"}
            if tmpl not in allowed_templates:
                return f"ERROR: Unsupported template: {tmpl} (allowed: default|blank|minimal_heading|research)"

            def _derive_template_seed(tmpl_name: str) -> str:
                if tmpl_name == "blank":
                    return ""
                if tmpl_name == "minimal_heading":
                    return "# Title\n\n"
                if tmpl_name == "research":
                    return ("# Title\n\n## Abstract\n\n## Introduction\n\n## Methods\n\n"
                            "## Results\n\n## Discussion\n\n")
                return ""  # default

            def _do_on_worker(word):
                created_path = None
                try:
                    # new document
                    doc = word.Documents.Add()
                    try:
                        if visible:
                            word.Visible = True
                    except Exception:
                        pass
                    # seed content or template if no explicit content
                    seed_text = content
                    if not seed_text:
                        seed_text = _derive_template_seed(tmpl)
                    if seed_text:
                        if cf == "markdown":
                            write_markdown_to_document(doc, seed_text)
                        elif cf in ("plain", "text"):
                            replace_document_content(doc, seed_text)
                        # cf == none -> ignore even if template provided
                    # save (use SaveAs2 when available)
                    try:
                        doc.SaveAs2(str(p))  # type: ignore[attr-defined]
                    except Exception:
                        doc.SaveAs(str(p))
                    created_path = str(p)
                    # activation/visibility
                    try:
                        if make_active:
                            doc.Activate()
                        else:
                            # Optionally close if not keeping active
                            doc.Close(SaveChanges=0)
                    except Exception:
                        pass
                    size = 0
                    try:
                        size = p.stat().st_size
                    except Exception:
                        pass
                    return {
                        "created": True,
                        "path": str(p),
                        "activated": bool(make_active),
                        "size_bytes": int(size),
                        "notes": "created",
                    }
                except Exception as ex:
                    # best-effort cleanup of partially created file
                    try:
                        if created_path and Path(created_path).exists():
                            Path(created_path).unlink(missing_ok=True)
                    except Exception:
                        pass
                    return {"error": f"create_word_document failed: {ex}"}

            async def _task_create():
                return await _run_on_worker_with_cleanup(p, _do_on_worker)

            res = await with_doc_lock(canonical_doc_key(str(p)), _task_create)
            # normalize error
            if isinstance(res, dict) and "error" in res:
                return f"ERROR: {res['error']}"
            return res
        except Exception as e:
            return f"ERROR: create_word_document failed: {e}"

    @mcp.tool()
    async def save_word_document_as(
        target_path: str,
        source_path: str = "",
        format: str = "docx",
        overwrite: bool = False,
        include_comments: bool = True,
        close_after: bool = False,
        reveal: bool = False,
    ) -> dict | str:
        """Save the specified (or active single) document to a new path/format.

        Supported formats: docx, pdf, md (markdown export via server).
        """
        from mcp_server.path_utils import resolve_user_path
        try:
            t = resolve_user_path(target_path)
            ok, msg = _sandbox_ok(t)
            if not ok:
                return f"ERROR: {msg}"
            try:
                t.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return f"ERROR: Failed to create parent directory: {e}"
            fmt = (format or "docx").strip().lower()
            if fmt not in ("docx", "pdf", "md"):
                return f"ERROR: Unsupported format: {fmt}"
            # Auto-normalize extension if missing / mismatched for docx or pdf; md handled separately
            expected_ext = ".docx" if fmt == "docx" else (".pdf" if fmt == "pdf" else ".md")
            if t.suffix.lower() != expected_ext:
                t = t.with_suffix(expected_ext)
            if t.exists() and not overwrite:
                return f"ERROR: Target already exists: {t}"

            # Word constants
            WD_FORMAT_PDF = 17
            WD_FORMAT_DOCX = 12  # wdFormatXMLDocument

            def _save_from_doc(doc, word):
                # Handle markdown export separately
                if fmt == "md":
                    md = get_markdown_representation(doc)
                    try:
                        with open(t, "w", encoding="utf-8") as f:
                            f.write(md)
                    except Exception as e:
                        return {"error": f"Failed to write markdown file: {e}"}
                else:
                    # choose format
                    fcode = WD_FORMAT_PDF if fmt == "pdf" else WD_FORMAT_DOCX
                    # Prefer SaveCopyAs to avoid changing source doc binding
                    try:
                        doc.SaveCopyAs(FileName=str(t), FileFormat=fcode)  # type: ignore[attr-defined]
                    except Exception:
                        try:
                            # Fallback to SaveAs2 (will rebind document path)
                            doc.SaveAs2(FileName=str(t), FileFormat=fcode)  # type: ignore[attr-defined]
                        except Exception:
                            doc.SaveAs(FileName=str(t), FileFormat=fcode)
                try:
                    if reveal:
                        word.Visible = True
                except Exception:
                    pass
                try:
                    if close_after and fmt != "md":
                        # If we used SaveAs2, document may now point to target; close it safely
                        try:
                            doc.Close(SaveChanges=0)
                        except Exception:
                            pass
                except Exception:
                    pass
                size = 0
                try:
                    size = t.stat().st_size
                except Exception:
                    pass
                return {"saved": True, "target_path": str(t), "bytes": int(size), "format": fmt, "notes": "saved"}

            if source_path:
                s = resolve_user_path(source_path)
                if not s.exists():
                    return f"ERROR: Source file not found: {s}"
                # run on worker keyed by source
                def _do_on_worker(word):
                    # find or open source doc
                    target = None
                    for i in range(1, max(1, getattr(word.Documents, "Count", 0)) + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(s, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(s), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(s))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass
                    return _save_from_doc(target, word)

                async def _task_save_path():
                    return await _run_on_worker_with_cleanup(s, _do_on_worker)

                res = await with_doc_lock(canonical_doc_key(str(s)), _task_save_path)
            else:
                # Use the only active document
                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    w = get_word_app()
                    if w.Documents.Count == 0:
                        return {"__error__": "No document is open. Provide 'source_path' or open one."}
                    if w.Documents.Count > 1:
                        return {"__error__": "Multiple documents are open. Provide 'source_path' to target the right file."}
                    d = get_active_document(w)
                    return _save_from_doc(d, w)

                async def _task_save_active():
                    return await asyncio.to_thread(_do_active)

                res = await with_global_lock(_task_save_active)

            if isinstance(res, dict) and "__error__" in res:
                return f"ERROR: {res['__error__']}"
            if isinstance(res, dict) and "error" in res:
                return f"ERROR: {res['error']}"
            return res
        except Exception as e:
            return f"ERROR: save_word_document_as failed: {e}"

    @mcp.tool()
    async def list_word_documents() -> list[dict] | str:
        """List open Word documents with basic metadata.

        Returns list of { path, read_only, active, modified }."""
        try:
            import asyncio
            def _do():
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                w = get_word_app()
                docs = []
                count = getattr(w.Documents, "Count", 0)
                active_full = None
                try:
                    if count:
                        active_full = w.ActiveDocument.FullName
                except Exception:
                    pass
                for i in range(1, count + 1):
                    try:
                        d = w.Documents(i)
                        docs.append({
                            "path": d.FullName,
                            "read_only": bool(getattr(d, "ReadOnly", False)),
                            "active": _paths_equal(Path(d.FullName), active_full) if active_full else False,
                            "modified": bool(getattr(d, "Saved", False) is False),
                        })
                    except Exception:
                        continue
                return docs
            return await asyncio.to_thread(_do)
        except Exception as e:
            return f"ERROR: list_word_documents failed: {e}"

    @mcp.tool()
    async def save_word_document(path: str = "") -> str:
        """Save the specified Word document or the single active one if no path given."""
        try:
            import asyncio
            from mcp_server.path_utils import resolve_user_path
            if path:
                p = resolve_user_path(path)
                def _do_on_worker(word):
                    for i in range(1, max(1, getattr(word.Documents, "Count", 0)) + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                try:
                                    d.Save()
                                    return "Saved"
                                except Exception as ex:
                                    return f"ERROR: Save failed: {ex}"
                        except Exception:
                            continue
                    return "ERROR: Document not open"
                async def _task():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)
                return await with_doc_lock(canonical_doc_key(str(p)), _task)
            else:
                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    w = get_word_app()
                    if w.Documents.Count == 0:
                        return "ERROR: No document open"
                    if w.Documents.Count > 1:
                        return "ERROR: Multiple documents open; provide 'path'"
                    try:
                        w.ActiveDocument.Save()
                        return "Saved"
                    except Exception as ex:
                        return f"ERROR: Save failed: {ex}"
                async def _task_active():
                    return await asyncio.to_thread(_do_active)
                return await with_global_lock(_task_active)
        except Exception as e:
            return f"ERROR: save_word_document failed: {e}"

    @mcp.tool()
    async def edit_word_document(task: str, ctx: Context) -> str:
        """Edits the active Word document according to the given task and returns a summary.

        Provide only the task instructions (a few sentences). The current document content is fetched automatically.
        """
        try:
            async def _task():
                markdown_edit_output = await run_markdown_edit(
                    markdown_edit_request=MarkdownEditRequest(context=ctx, task=task)
                )
                output_string = markdown_edit_output.change_summary + "\n" + markdown_edit_output.output_message
                return output_string

            # Global lock (active document may be unknown/unsaved)
            return await with_global_lock(_task)
        except Exception as e:
            return f"ERROR: edit_word_document failed: {e}"

    @mcp.tool()
    async def add_comments_to_word_document(ctx: Context) -> str:
        """
        Runs a routine that will add feedback as comments to the currently open Word Document.
        """
        try:
            async def _task():
                comment_output = await run_feedback_step(
                    markdown_edit_request=MarkdownEditRequest(context=ctx),
                )
                return comment_output.feedback_summary

            return await with_global_lock(_task)
        except Exception as e:
            return f"ERROR: add_comments_to_word_document failed: {e}"

    @mcp.tool()
    async def analyze_comments(ctx: Context) -> str:
        """
        Runs a routine that analyze the comments in the Word document and determine how they could be solved.
        """
        try:
            async def _task():
                comment_analysis_output = await run_comment_analysis(
                    markdown_edit_request=MarkdownEditRequest(context=ctx),
                )
                return comment_analysis_output.edit_instructions + "\n" + comment_analysis_output.assistant_hints

            return await with_global_lock(_task)
        except Exception as e:
            return f"ERROR: analyze_comments failed: {e}"

    # TODO: It might be good to consider having the document content always be available to the assistant if the document is "connected".
    @mcp.tool()
    async def get_word_content(ctx: Context) -> str:
        """
        Returns the content of the open Word document. Use this tool when you just need the content of the document.
        You should use this after making edits to get the current state of the document.
        """
        try:
            async def _task():
                word = get_word_app()
                doc = get_active_document(word)
                markdown_from_word = get_markdown_representation(doc)
                return markdown_from_word

            return await with_global_lock(_task)
        except Exception as e:
            return f"ERROR: get_word_content failed: {e}"

    @mcp.tool()
    async def highlight(
        text: str,
        scope: str = "first",
        color: str | int = "yellow",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> dict:
        """Highlight text deterministically with enhanced validation.

        Args:
            text: Text to search for (required, non-empty)
            scope: "first" or "all" (default "first")
            color: Highlight color - CSS color names, hex codes, or Word color constants
            case_sensitive: Match case exactly if True
            whole_word: Match whole words only if True  
            max_matches: Limit number of matches (0 = no limit)
            path: Optional path to specific document

        Returns:
            { matches: int, color: int, scope: str } or { error: str }
        
        Supported highlight colors: yellow, red, blue, green, pink, turquoise, bright_green, 
        gray, dark_blue, dark_red, dark_yellow, dark_green, violet, teal, gray_25, gray_50, 
        white, black, none/clear (to remove highlighting)
        """
        # Enhanced validation
        s = (scope or "first").strip().lower()
        if s not in ("first", "all"):
            return {"error": "Invalid scope. Use 'first' or 'all'"}
        
        if not (text or "").strip():
            return {"error": "Text parameter is required and cannot be empty"}
        
        # Performance limit validation
        if max_matches > 10000:
            return {"error": "Maximum 10,000 matches allowed for performance reasons"}
        
        try:
            import asyncio
            hci = _parse_highlight_color(color)
            
            # Enhanced color validation - validate all color inputs
            valid_colors = list(HIGHLIGHT_COLOR_MAP.keys())
            
            # Validate color input
            color_error = None
            if isinstance(color, str):
                if color.strip().lower() not in valid_colors:
                    color_error = f"Invalid highlight color '{color}'. Supported colors: {', '.join(sorted(set(valid_colors)))}, or integers 0-16"
            elif isinstance(color, (int, float)):
                try:
                    int_color = int(color)
                    if int_color < 0 or int_color > 16:
                        color_error = f"Invalid highlight color '{color}'. Supported integer range: 0-16"
                except (ValueError, TypeError):
                    color_error = f"Invalid highlight color '{color}'. Must be a valid color name or integer 0-16"
            else:
                color_error = f"Invalid highlight color type. Must be string or integer, got {type(color).__name__}"
            
            if color_error:
                return {"error": color_error}
            
            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return {"error": f"File not found: {p}"}

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    # Find or open the target document
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass
                    rng = target.Content
                    rng.Find.ClearFormatting()
                    flags = {
                        'FindText': text,
                        'MatchCase': case_sensitive,
                        'MatchWholeWord': whole_word,
                        'Wrap': 0,
                        'Forward': True,
                    }
                    count = 0
                    found = rng.Find.Execute(**flags)
                    while found:
                        current_end = rng.End
                        rng.HighlightColorIndex = hci
                        count += 1
                        if s == 'first' or (max_matches and count >= max_matches):
                            break
                        start_pos = current_end
                        if start_pos >= target.Content.End:
                            break
                        rng = target.Range(start_pos, target.Content.End)
                        rng.Find.ClearFormatting()
                        found = rng.Find.Execute(**flags)
                        if found and rng.End <= current_end:
                            break
                        # Performance check
                        if count >= 10000:
                            break
                    if count:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return count

                async def _task_hl_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_hl_path)
            else:
                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass
                    rng = doc.Content
                    rng.Find.ClearFormatting()
                    flags = {
                        'FindText': text,
                        'MatchCase': case_sensitive,
                        'MatchWholeWord': whole_word,
                        'Wrap': 0,
                        'Forward': True,
                    }
                    count = 0
                    found = rng.Find.Execute(**flags)
                    while found:
                        current_end = rng.End
                        rng.HighlightColorIndex = hci
                        count += 1
                        if s == 'first' or (max_matches and count >= max_matches):
                            break
                        start_pos = current_end
                        if start_pos >= doc.Content.End:
                            break
                        rng = doc.Range(start_pos, doc.Content.End)
                        rng.Find.ClearFormatting()
                        found = rng.Find.Execute(**flags)
                        if found and rng.End <= current_end:
                            break
                        # Performance check
                        if count >= 10000:
                            break
                    return count

                async def _task_hl_active():
                    return await asyncio.to_thread(_do_active)

                applied = await with_global_lock(_task_hl_active)
                if applied == -1:
                    return {"error": "No document is open. Provide 'path' or open one."}
                if applied == -2:
                    return {"error": "Multiple documents are open. Provide 'path' to target the right file."}
                if applied == -3:
                    return {"error": "Active document is read-only. Provide 'path' to edit a writable copy."}
            return {"matches": int(applied), "color": int(hci), "scope": s}
        except Exception as e:
            return {"error": f"highlight failed: {e}"}

    @mcp.tool()
    async def batch_format(
        scope: str = "document",
        operations: list[dict] | None = None,
        find: dict | str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> dict:
        """Apply multiple deterministic formatting operations in order to a scope with atomic behavior.

        Atomic Behavior: All-or-nothing execution. If any operation fails validation or execution, 
        no changes are applied to the document.

        Performance Limits: Maximum 50 operations per call, maximum 10,000 text matches processed.

        Supported op keys within each item of `operations`:
        - font: { name?, size?, color?, bold?, italic?, underline?, strikethrough?, superscript?, subscript? }
        - paragraph: { alignment?, line_spacing_rule?, line_spacing?, space_before?, space_after? }
        - list: { type: "bullet" | "numbered" | "none" }
        - style: { name }
        - clear_formatting: { target?: "all" | "font" | "paragraph" }
        - document_base_font: { name?, size?, color?, bold?, italic?, underline? }

        Color formats supported: CSS color names ("red", "blue"), hex codes ("#FF0000"), Word color constants (16711680).
        Font sizes: 1-1638 points. Paragraph spacing: 0-1584 points.
        """
        try:
            ops = operations or []
            if not isinstance(ops, list) or not ops:
                return {"changed": False, "summary": {"operations": 0, "ranges_affected": 0}, "details": []}

            # Performance limit validation
            if len(ops) > 50:
                return {"error": "Maximum 50 operations allowed per call"}

            # Scope validation
            s = (scope or "document").strip().lower()
            if s not in ("document", "selection", "matches"):
                return {"error": "Invalid scope. Use 'document', 'selection', or 'matches'"}

            def _validate_ops(ops_: list[dict]) -> str | None:
                for idx, op in enumerate(ops_):
                    if not isinstance(op, dict) or not op:
                        return f"Invalid operation at index {idx}"
                    
                    # Font validation
                    if "font" in op:
                        fnt = op["font"]
                        if not isinstance(fnt, dict):
                            return f"Invalid font object at index {idx}"
                        
                        # Mutual exclusion check
                        if fnt.get("superscript") and fnt.get("subscript"):
                            return "Superscript and subscript are mutually exclusive"
                        
                        # Font size validation
                        if "size" in fnt and fnt["size"] is not None:
                            try:
                                size = int(fnt["size"])
                                if size < 1 or size > 1638:
                                    return f"Font size must be between 1 and 1638 points at index {idx}"
                            except (ValueError, TypeError):
                                return f"Invalid font size at index {idx}"
                        
                        # Color validation
                        if "color" in fnt and fnt["color"] is not None and fnt["color"] != "":
                            color_val = parse_font_color(fnt["color"])
                            if color_val is None:
                                return f"Invalid font color '{fnt['color']}' at index {idx}. Use CSS color names, hex codes (#FF0000), or Word color constants"
                    
                    # Paragraph validation
                    if "paragraph" in op:
                        para = op["paragraph"]
                        if not isinstance(para, dict):
                            return f"Invalid paragraph object at index {idx}"
                        
                        # Alignment validation
                        if "alignment" in para and para["alignment"]:
                            align = str(para["alignment"]).lower()
                            if align not in ("left", "center", "right", "justify"):
                                return f"Invalid alignment '{para['alignment']}' at index {idx}. Use: left, center, right, justify"
                        
                        # Spacing validation
                        for spacing_key in ("space_before", "space_after"):
                            if spacing_key in para and para[spacing_key] is not None:
                                try:
                                    spacing = float(para[spacing_key])
                                    if spacing < 0 or spacing > 1584:
                                        return f"{spacing_key} must be between 0 and 1584 points at index {idx}"
                                except (ValueError, TypeError):
                                    return f"Invalid {spacing_key} value at index {idx}"
                        
                        # Line spacing validation
                        if "line_spacing_rule" in para and para["line_spacing_rule"]:
                            rule = str(para["line_spacing_rule"]).lower()
                            if rule not in ("single", "1.5", "double", "multiple", "exact", "minimum"):
                                return f"Invalid line_spacing_rule '{para['line_spacing_rule']}' at index {idx}"
                    
                    # List validation
                    if "list" in op:
                        list_obj = op["list"]
                        if not isinstance(list_obj, dict):
                            return f"Invalid list object at index {idx}"
                        t = str(list_obj.get("type", "")).lower()
                        if t not in ("bullet", "numbered", "none"):
                            return f"Invalid list type '{list_obj.get('type')}' at index {idx}. Use: bullet, numbered, none"
                    
                    # Style validation will happen during execution for style existence
                    
                    # Clear formatting validation
                    if "clear_formatting" in op:
                        cf = op["clear_formatting"]
                        if isinstance(cf, dict) and "target" in cf:
                            target = str(cf["target"]).lower()
                            if target not in ("all", "font", "paragraph"):
                                return f"Invalid clear_formatting target '{cf['target']}' at index {idx}. Use: all, font, paragraph"
                    
                    # Document base font validation
                    if "document_base_font" in op:
                        base = op["document_base_font"]
                        if not isinstance(base, dict):
                            return f"Invalid document_base_font object at index {idx}"
                        
                        # Size validation
                        if "size" in base and base["size"] is not None:
                            try:
                                size = int(base["size"])
                                if size < 1 or size > 1638:
                                    return f"Document base font size must be between 1 and 1638 points at index {idx}"
                            except (ValueError, TypeError):
                                return f"Invalid document base font size at index {idx}"
                        
                        # Color validation
                        if "color" in base and base["color"] is not None and base["color"] != "":
                            color_val = parse_font_color(base["color"])
                            if color_val is None:
                                return f"Invalid document base font color '{base['color']}' at index {idx}"
                
                return None

            msg = _validate_ops(ops)
            if msg:
                return {"error": msg}

            # Decode find input (support object form)
            find_text: str = ""
            f_case = bool(case_sensitive)
            f_whole = bool(whole_word)
            f_max = int(max_matches or 0)
            if isinstance(find, dict):
                find_text = str(find.get("text", "") or "")
                f_case = bool(find.get("case_sensitive", f_case))
                f_whole = bool(find.get("whole_word", f_whole))
                f_max = int(find.get("max_matches", f_max) or 0)
            else:
                find_text = str(find or "")
            if s == "matches" and not find_text.strip():
                return {"error": "When scope=='matches', provide find.text (or 'find' string)"}

            # Performance limit for matches
            if f_max > 10000:
                return {"error": "Maximum 10,000 matches allowed for performance reasons"}

            # Check for superscript/subscript conflicts across operations for same scope
            wants_super = any(isinstance(op, dict) and "font" in op and dict(op["font"]).get("superscript") for op in ops)
            wants_sub = any(isinstance(op, dict) and "font" in op and dict(op["font"]).get("subscript") for op in ops)
            if wants_super and wants_sub:
                return {"error": "Superscript and subscript are mutually exclusive for the same target range"}

            def _apply_ops_atomic(word, doc):
                """Apply all operations atomically with rollback on failure."""
                import tempfile
                import os
                from mcp_server.app_interaction.word_editor import get_selection
                
                # First, validate that we can execute all operations without actually changing anything
                validation_errors = []
                
                # Validate style existence
                for idx, op in enumerate(ops):
                    if "style" in op:
                        style_name = (op["style"].get("name") or "").strip()
                        if style_name:
                            try:
                                _ = doc.Styles(style_name)
                            except Exception:
                                validation_errors.append(f"Unknown style '{style_name}' at operation {idx}")
                
                if validation_errors:
                    raise ValueError("; ".join(validation_errors))
                
                # Check if selection scope has an actual selection
                if s == "selection":
                    sel = get_selection(word)
                    if sel is None or sel.Range.Start == sel.Range.End:
                        raise ValueError("Selection scope requires an active text selection")
                
                # Save document content for potential rollback (without closing document)
                undo_record_started = False
                try:
                    # Start an undo record for atomic rollback
                    doc.UndoRecord.StartCustomRecord("Batch Format Operations")
                    undo_record_started = True
                except Exception:
                    undo_record_started = False  # Fallback: no atomic rollback capability
                
                try:
                    # Track distinct ranges changed across the entire batch
                    distinct_changed: set[tuple[int, int]] = set()
                    total_ranges = 0
                    details: list[dict] = []
                    total_matches_processed = 0  # Track actual text matches, not operations

                    def _make_fn(op: dict):
                        if "font" in op:
                            f = op["font"]
                            # Validate color early if provided
                            color_in = f.get("color") if isinstance(f, dict) else None
                            if color_in is not None and color_in != "":
                                cv = parse_font_color(color_in)
                                if cv is None:
                                    raise ValueError(f"Invalid font color: {color_in}")
                                color_val = cv
                            else:
                                color_val = None
                            return "font", (lambda rng: set_range_font(
                                rng,
                                name=(f.get("name") or "").strip() or None,
                                size=f.get("size"),
                                color=color_val,
                                bold=f.get("bold"),
                                italic=f.get("italic"),
                                underline=f.get("underline"),
                                strikethrough=f.get("strikethrough"),
                                superscript=f.get("superscript"),
                                subscript=f.get("subscript"),
                            ))
                        if "paragraph" in op:
                            p = op["paragraph"]
                            # Accept alias line_spacing_rule as shorthand
                            ls_val = (p.get("line_spacing_rule") or p.get("line_spacing") or "").strip() or None
                            return "paragraph", (lambda rng: set_range_paragraph(
                                rng,
                                alignment=(p.get("alignment") or "").strip() or None,
                                line_spacing=ls_val,
                                space_before=p.get("space_before"),
                                space_after=p.get("space_after"),
                            ))
                        if "list" in op:
                            t = str(op["list"].get("type", "")).lower()
                            return "list", (lambda rng: apply_list_format(rng, t))
                        if "style" in op:
                            s_ = (op["style"].get("name") or "").strip()
                            return "style", (lambda rng: apply_style_to_range(rng, s_))
                        if "clear_formatting" in op:
                            cf = op["clear_formatting"] if isinstance(op["clear_formatting"], dict) else {}
                            tgt = (cf.get("target") or "all").strip().lower()
                            if tgt not in ("all", "font", "paragraph"):
                                tgt = "all"
                            return "clear_formatting", (lambda rng, t=tgt: clear_direct_fmt(rng, target=t))
                        if "document_base_font" in op:
                            b = op["document_base_font"]
                            # Validate color if provided
                            color_in = b.get("color") if isinstance(b, dict) else None
                            color_val_db = None
                            if color_in is not None and color_in != "":
                                cv = parse_font_color(color_in)
                                if cv is None:
                                    raise ValueError(f"Invalid base font color: {color_in}")
                                color_val_db = cv
                            return "document_base_font", (lambda _rng: update_normal_style(
                                doc,
                                name=(b.get("name") or "").strip() or None,
                                size=b.get("size"),
                                color=color_val_db,
                                bold=b.get("bold"),
                                italic=b.get("italic"),
                                underline=b.get("underline"),
                            ))
                        return "unknown", lambda _r: None

                    def _collect_changed(rng):
                        try:
                            s_pos = int(getattr(rng, 'Start', -1))
                            e_pos = int(getattr(rng, 'End', -1))
                            if s_pos >= 0 and e_pos >= 0 and e_pos >= s_pos:
                                distinct_changed.add((s_pos, e_pos))
                        except Exception:
                            pass

                    # Count total text matches that will be processed for performance validation
                    if s == "matches" and find_text:
                        from mcp_server.app_interaction.word_editor import iter_find_ranges
                        match_count = 0
                        for _ in iter_find_ranges(doc, find_text, match_case=f_case, whole_word=f_whole, max_matches=f_max):
                            match_count += 1
                            if match_count > 10000:
                                raise ValueError("Performance limit exceeded: more than 10,000 text matches would be processed")
                        total_matches_processed = match_count * len(ops)  # Each operation processes all matches

                    for idx, op in enumerate(ops):
                        op_type, fn = _make_fn(op)
                        if op_type == "unknown":
                            details.append({"op_index": idx, "op_type": op_type, "ranges_affected": 0})
                            continue
                        
                        # Use enhanced apply_to_scope that errors on empty selection
                        applied = apply_to_scope_strict(
                            word,
                            doc,
                            s,
                            find_text=find_text,
                            match_case=f_case,
                            whole_word=f_whole,
                            max_matches=f_max,
                            fn=fn,
                            on_applied=_collect_changed,
                        )
                        total_ranges += applied
                        details.append({"op_index": idx, "op_type": op_type, "ranges_affected": applied})
                    
                    return len(distinct_changed), details, total_matches_processed
                
                except Exception as e:
                    # Rollback using Word's Undo mechanism (preserves document reference)
                    if undo_record_started:
                        try:
                            doc.UndoRecord.EndCustomRecord()
                            doc.Undo()  # Undo all operations in the record
                            raise ValueError(f"Operation failed and changes were undone: {e}")
                        except Exception:
                            # If undo fails, just raise the original error
                            pass
                    raise e
                finally:
                    # End undo record if it was started
                    if undo_record_started:
                        try:
                            doc.UndoRecord.EndCustomRecord()
                        except Exception:
                            pass

            if path:
                from mcp_server.path_utils import resolve_user_path
                import os
                p = resolve_user_path(path)
                if not p.is_file():
                    return {"error": f"File not found: {p}"}

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    # Find or open the target document
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass
                    try:
                        rngs, det, matches_processed = _apply_ops_atomic(word, target)
                    except Exception as ex:
                        return {"__error__": str(ex)}
                    if rngs:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return rngs, det, matches_processed

                async def _task_batch_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                res = await with_doc_lock(canonical_doc_key(str(p)), _task_batch_path)
                if isinstance(res, dict) and "__error__" in res:
                    return {"error": res["__error__"]}
                total, details, matches_processed = res
            else:
                import asyncio
                from mcp_server.app_interaction.word_editor import get_selection, apply_to_scope_strict
                
                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1, []
                    if word.Documents.Count > 1:
                        return -2, []
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3, []
                    except Exception:
                        pass
                    try:
                        return _apply_ops_atomic(word, doc)
                    except Exception as ex:
                        return {"__error__": str(ex)}

                async def _task_batch_active():
                    return await asyncio.to_thread(_do_active)

                res = await with_global_lock(_task_batch_active)
                if isinstance(res, dict) and "__error__" in res:
                    return {"error": res["__error__"]}
                total, details, matches_processed = res
                if total == -1:
                    return {"error": "No document is open. Provide 'path' or open one."}
                if total == -2:
                    return {"error": "Multiple documents are open. Provide 'path' to target the right file."}
                if total == -3:
                    return {"error": "Active document is read-only. Provide 'path' to edit a writable copy."}
            return {"changed": total > 0, "summary": {"text_matches_processed": matches_processed, "ranges_affected": int(total) }, "details": details}
        except Exception as e:
            return {"error": f"batch_format failed: {e}"}

    @mcp.tool()
    async def search_and_replace(
        find: str,
        replace: str,
        all: bool = True,
        case_sensitive: bool = False,
        whole_word: bool = False,
        preview_only: bool = False,
        preserve_case: bool = False,
        path: str = "",
    ) -> str:
        """Perform deterministic search & replace in the active Word document.

        Args:
            find: Text to locate (must be non-empty).
            replace: Replacement text (can be empty to delete matches).
            all: If True replace every occurrence; if False only the first.
            case_sensitive: Match case exactly if True.
            whole_word: Match only whole words if True.
            preview_only: If True, do not modify—just report the count it *would* change.

        Returns:
            Summary message: counts or error.
        """
        if not find:
            return "ERROR: 'find' text is required"
        if find == replace and not preview_only:
            return "No-op: find and replace text are identical"
        try:
            import asyncio

            def _preserve(original: str, replacement: str) -> str:
                if not original:
                    return replacement
                if original.isupper():
                    return replacement.upper()
                if original.islower():
                    return replacement.lower()
                if original[0].isupper() and original[1:].islower():
                    return replacement[:1].upper() + replacement[1:].lower()
                return replacement

            def _replace_in_doc(doc):
                try:
                    protection = doc.ProtectionType
                except Exception:
                    protection = -1
                if protection not in (-1, 0):
                    return (0, "ERROR: Document is protected/read-only")
                rng = doc.Content
                rng.Find.ClearFormatting()
                flags = {
                    'FindText': find,
                    'MatchCase': case_sensitive,
                    'MatchWholeWord': whole_word,
                    'Wrap': 0,  # wdFindStop
                    'Forward': True,
                }
                count = 0
                found = rng.Find.Execute(**flags)
                while found:
                    count += 1
                    current_end = rng.End
                    if not preview_only:
                        new_text = _preserve(rng.Text, replace) if preserve_case else replace
                        rng.Text = new_text
                        start_pos = rng.End
                        rng = doc.Range(start_pos, doc.Content.End)
                    else:
                        start_pos = rng.End
                        rng = doc.Range(start_pos, doc.Content.End)
                    if not all:
                        break
                    rng.Find.ClearFormatting()
                    found = rng.Find.Execute(**flags)
                    # Safety: ensure forward progress
                    if found and rng.End <= current_end:
                        break
                return (count, None)

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    # Find or open target document
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass
                    count, err = _replace_in_doc(target)
                    if not preview_only and count:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return count, err

                async def _task_sr_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                count, err = await with_doc_lock(canonical_doc_key(str(p)), _task_sr_path)
            else:
                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return 0, "ERROR: No document is open. Provide 'path' or open one."
                    if word.Documents.Count > 1:
                        return 0, "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                    doc = get_active_document(word)
                    return _replace_in_doc(doc)

                async def _task_sr_active():
                    return await asyncio.to_thread(_do_active)

                count, err = await with_global_lock(_task_sr_active)
            if err:
                return err
            if count == 0:
                return "No matches found"
            if preview_only:
                scope = "all" if all else "first"
                return f"Preview: would replace {count} {scope} occurrence(s)"
            else:
                scope = "all" if all else "first"
                return f"Replaced {count} {scope} occurrence(s)"
        except Exception as e:
            return f"ERROR: search_and_replace failed: {e}"

    # region New deterministic formatting tools

    @mcp.tool()
    async def set_document_base_font(
        name: str = "",
        size: float | int | None = None,
        color: str = "",
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        path: str = "",
    ) -> str:
        """Update the Normal style (base font) in a Word document.

        Any provided property will be updated. If none are provided, it's a no-op.
        """
        try:
            has_change = any([
                bool(name.strip()),
                size is not None,
                bool(color.strip()),
                bold is not None,
                italic is not None,
                underline is not None,
            ])
            if not has_change:
                return "No changes requested"

            color_val = parse_font_color(color) if color else None

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass
                    ok = update_normal_style(
                        target,
                        name=name.strip() or None,
                        size=size,
                        color=color_val,
                        bold=bold,
                        italic=italic,
                        underline=underline,
                    )
                    if ok:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return ok

                async def _task_base_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                updated = await with_doc_lock(canonical_doc_key(str(p)), _task_base_path)
                return "Updated Normal style" if updated else "Failed to update Normal style"
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass
                    ok = update_normal_style(
                        doc,
                        name=name.strip() or None,
                        size=size,
                        color=color_val,
                        bold=bold,
                        italic=italic,
                        underline=underline,
                    )
                    return 1 if ok else 0

                async def _task_base_active():
                    return await asyncio.to_thread(_do_active)

                res = await with_global_lock(_task_base_active)
                if res == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if res == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if res == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
                return "Updated Normal style" if res == 1 else "Failed to update Normal style"
        except Exception as e:
            return f"ERROR: set_document_base_font failed: {e}"

    @mcp.tool()
    async def set_font(
        scope: str = "document",
        name: str = "",
        size: float | int | None = None,
        color: str = "",
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        find: str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> str:
        """Apply character formatting deterministically to a scope (document/selection/matches)."""
        try:
            has_change = any([
                bool(name.strip()),
                size is not None,
                bool(color.strip()),
                bold is not None,
                italic is not None,
                underline is not None,
            ])
            if not has_change:
                return "No changes requested"

            color_val = parse_font_color(color) if color else None

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass

                    def _apply(rng):
                        return set_range_font(
                            rng,
                            name=name.strip() or None,
                            size=size,
                            color=color_val,
                            bold=bold,
                            italic=italic,
                            underline=underline,
                        )

                    applied = apply_to_scope(
                        word,
                        target,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    if applied:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return applied

                async def _task_setfont_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_setfont_path)
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass

                    def _apply(rng):
                        return set_range_font(
                            rng,
                            name=name.strip() or None,
                            size=size,
                            color=color_val,
                            bold=bold,
                            italic=italic,
                            underline=underline,
                        )

                    applied = apply_to_scope(
                        word,
                        doc,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    return applied

                async def _task_setfont_active():
                    return await asyncio.to_thread(_do_active)

                applied = await with_global_lock(_task_setfont_active)
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            return ("Applied formatting to document" if applied == 1 and scope != "matches" else f"Applied formatting to {applied} target(s)") if applied else "No targets found"
        except Exception as e:
            return f"ERROR: set_font failed: {e}"

    @mcp.tool()
    async def set_paragraph_format(
        scope: str = "document",
        alignment: str = "",
        line_spacing: str = "",
        space_before: float | int | None = None,
        space_after: float | int | None = None,
        find: str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> str:
        """Apply paragraph alignment/spacing to a scope deterministically."""
        try:
            has_change = any([
                bool(alignment.strip()),
                bool(line_spacing.strip()),
                space_before is not None,
                space_after is not None,
            ])
            if not has_change:
                return "No changes requested"

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass

                    def _apply(rng):
                        return set_range_paragraph(
                            rng,
                            alignment=alignment.strip() or None,
                            line_spacing=line_spacing.strip() or None,
                            space_before=space_before,
                            space_after=space_after,
                        )

                    applied = apply_to_scope(
                        word,
                        target,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    if applied:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return applied

                async def _task_setpara_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_setpara_path)
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass

                    def _apply(rng):
                        return set_range_paragraph(
                            rng,
                            alignment=alignment.strip() or None,
                            line_spacing=line_spacing.strip() or None,
                            space_before=space_before,
                            space_after=space_after,
                        )

                    applied = apply_to_scope(
                        word,
                        doc,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    return applied

                async def _task_setpara_active():
                    return await asyncio.to_thread(_do_active)

                applied = await with_global_lock(_task_setpara_active)
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            return ("Applied paragraph formatting to document" if applied == 1 and scope != "matches" else f"Applied paragraph formatting to {applied} target(s)") if applied else "No targets found"
        except Exception as e:
            return f"ERROR: set_paragraph_format failed: {e}"

    @mcp.tool()
    async def set_list_style(
        scope: str = "selection",
        type: str = "bullet",
        find: str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> str:
        """Apply or remove list formatting (bullet/numbered/none) on the scope."""
        try:
            lt = type.strip().lower()
            if lt not in ("bullet", "numbered", "none"):
                return "ERROR: Invalid list type. Use 'bullet', 'numbered', or 'none'."

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass

                    def _apply(rng):
                        return apply_list_format(rng, lt)

                    applied = apply_to_scope(
                        word,
                        target,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    if applied:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return applied

                async def _task_setlist_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_setlist_path)
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass

                    def _apply(rng):
                        return apply_list_format(rng, lt)

                    applied = apply_to_scope(
                        word,
                        doc,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    return applied

                async def _task_setlist_active():
                    return await asyncio.to_thread(_do_active)

                applied = await with_global_lock(_task_setlist_active)
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            return ("Applied list formatting to document" if applied == 1 and scope != "matches" else f"Applied list formatting to {applied} target(s)") if applied else "No targets found"
        except Exception as e:
            return f"ERROR: set_list_style failed: {e}"

    @mcp.tool()
    async def apply_style(
        scope: str = "selection",
        style_name: str = "",
        find: str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> str:
        """Apply a built-in style (e.g., 'Normal', 'Heading 1') to the scope."""
        try:
            if not style_name.strip():
                return "ERROR: style_name is required"

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass

                    # Pre-validate style existence
                    try:
                        _ = target.Styles(style_name)
                    except Exception:
                        return {"__error__": f"Unknown style: {style_name}"}

                    def _apply(rng):
                        return apply_style_to_range(rng, style_name)

                    applied = apply_to_scope(
                        word,
                        target,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    if applied:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return applied

                async def _task_applystyle_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                res = await with_doc_lock(canonical_doc_key(str(p)), _task_applystyle_path)
                if isinstance(res, dict) and "__error__" in res:
                    return res["__error__"]
                applied = res
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass
                    # Pre-validate style existence
                    try:
                        _ = doc.Styles(style_name)
                    except Exception:
                        return {"__error__": f"Unknown style: {style_name}"}

                    def _apply(rng):
                        return apply_style_to_range(rng, style_name)

                    applied = apply_to_scope(
                        word,
                        doc,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    return applied

                async def _task_applystyle_active():
                    return await asyncio.to_thread(_do_active)

                res = await with_global_lock(_task_applystyle_active)
                if isinstance(res, dict) and "__error__" in res:
                    return res["__error__"]
                applied = res
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            return ("Applied style to document" if applied == 1 and scope != "matches" else f"Applied style to {applied} target(s)") if applied else "No targets found"
        except Exception as e:
            return f"ERROR: apply_style failed: {e}"

    @mcp.tool()
    async def clear_direct_formatting(
        scope: str = "document",
        find: str = "",
        case_sensitive: bool = False,
        whole_word: bool = False,
        max_matches: int = 0,
        path: str = "",
    ) -> str:
        """Clear direct character and paragraph formatting on the scope (retains styles)."""
        try:
            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"

                def _do_on_worker(word):
                    _prepare_file_for_edit(p)
                    target = None
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                target = d
                                break
                        except Exception:
                            continue
                    if target is None:
                        try:
                            target = word.Documents.Open(str(p), ReadOnly=False)
                        except Exception:
                            target = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                    try:
                        _make_doc_editable(target)
                    except Exception:
                        pass

                    def _apply(rng):
                        return clear_direct_fmt(rng)

                    applied = apply_to_scope(
                        word,
                        target,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    if applied:
                        try:
                            target.Save()
                        except Exception:
                            pass
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return applied

                async def _task_clearfmt_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_clearfmt_path)
            else:
                import asyncio

                def _do_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return -1
                    if word.Documents.Count > 1:
                        return -2
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return -3
                    except Exception:
                        pass

                    def _apply(rng):
                        return clear_direct_fmt(rng)

                    applied = apply_to_scope(
                        word,
                        doc,
                        scope,
                        find_text=find,
                        match_case=case_sensitive,
                        whole_word=whole_word,
                        max_matches=max_matches,
                        fn=_apply,
                    )
                    return applied

                async def _task_clearfmt_active():
                    return await asyncio.to_thread(_do_active)

                applied = await with_global_lock(_task_clearfmt_active)
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            return ("Cleared direct formatting on document" if applied == 1 and scope != "matches" else f"Cleared direct formatting on {applied} target(s)") if applied else "No targets found"
        except Exception as e:
            return f"ERROR: clear_direct_formatting failed: {e}"

    # endregion

    @mcp.tool()
    async def get_powerpoint_content() -> str:
        """
        Returns the content of all slides in the active PowerPoint presentation.
        """
        powerpoint = get_powerpoint_app()
        presentation = get_active_presentation(powerpoint)
        return get_presentation_content(presentation)

    @mcp.tool()
    async def add_powerpoint_slide(slide_number: int, text: str) -> bool:
        """
        Adds a new slide at the specified position with the given text. Always call get_powerpoint_content to get the latest content and slide numbers.
        DO NOT use Markdown formatting for the text, it will not be rendered correctly. Use plaintext.
        At a maximum, add two bullet points to each slide.

        Args:
            slide_number: The position where to add the new slide
            text: The text to add to the slide

        Returns:
            True if the slide was added successfully, False otherwise
        """
        powerpoint = get_powerpoint_app()
        presentation = get_active_presentation(powerpoint)

        # Add new blank slide
        presentation.Slides.Add(slide_number, 12)

        # Add the text to the new slide
        add_text_to_slide(presentation, slide_number, text)
        return True

    @mcp.tool()
    async def remove_powerpoint_slide(slide_number: int) -> bool:
        """
        Removes the slide at the specified position. Always call get_powerpoint_content to get the latest content and slide numbers.

        Args:
            slide_number: The position of the slide to remove

        Returns:
            True if the slide was removed successfully, False otherwise
        """
        try:
            powerpoint = get_powerpoint_app()
            presentation = get_active_presentation(powerpoint)

            if slide_number <= 0 or slide_number > presentation.Slides.Count:
                return False

            presentation.Slides(slide_number).Delete()
            return True
        except Exception:
            return False

    @mcp.tool()
    async def get_excel_content() -> str:
        """
        Returns the content of the first worksheet in the active Excel workbook as a markdown table.

        Returns:
            A string containing the worksheet data formatted as a markdown table.
            First row is treated as headers.
        """
        excel = get_excel_app()
        workbook = get_active_workbook(excel)
        return get_workbook_content(workbook)

    @mcp.tool()
    async def close_word(path: str = "", save: bool = False) -> str:
        """Close a specific Word document by path, or all documents if none specified.

        Args:
            path: Optional absolute path to a .docx. If empty, closes all open docs.
            save: Whether to save changes before closing.
        """
        try:
            import asyncio

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)

                def _do_on_worker(word):
                    # Close the specific document if open in this instance
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                d.Close(SaveChanges=int(bool(save)))
                                return 1
                        except Exception:
                            continue
                    return 0

                async def _task_close_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                closed = await with_doc_lock(canonical_doc_key(str(p)), _task_close_path)
                if closed:
                    return "Closed document"
                return "Document not open"
            else:
                def _do_close_all():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    from mcp_server.app_interaction.word_editor import get_word_app
                    word = get_word_app()
                    count = int(word.Documents.Count)
                    try:
                        word.DisplayAlerts = 0
                    except Exception:
                        pass
                    # Close all documents in this instance
                    for _ in range(count):
                        try:
                            word.ActiveDocument.Close(SaveChanges=int(bool(save)))
                        except Exception:
                            break
                    return count

                async def _task_close_all():
                    return await asyncio.to_thread(_do_close_all)

                closed = await with_global_lock(_task_close_all)
                if closed:
                    return f"Closed {closed} document(s)"
                return "No documents open"
        except Exception as e:
            return f"ERROR: close_word failed: {e}"

    @mcp.tool()
    async def reveal_word(path: str = "") -> str:
        """Make the Word window for a given path visible and activate it (debug aid).

        If path is omitted, reveals the current Word instance managed by the active document.
        """
        try:
            import asyncio
            if path:
                p = Path(path).expanduser()

                def _do_on_worker(word):
                    # Ensure instance is visible and activate matching document
                    try:
                        word.Visible = True
                    except Exception:
                        pass
                    for i in range(1, word.Documents.Count + 1):
                        try:
                            d = word.Documents(i)
                            if _paths_equal(p, d.FullName):
                                d.Activate()
                                try:
                                    word.WindowState = 0  # wdWindowStateNormal
                                except Exception:
                                    pass
                                return "Revealed"
                        except Exception:
                            continue
                    return "Document not open in this instance"

                async def _task_reveal():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                return await with_doc_lock(canonical_doc_key(str(p)), _task_reveal)
            else:
                def _do_reveal_active():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    from mcp_server.app_interaction.word_editor import get_word_app
                    word = get_word_app()
                    try:
                        word.Visible = True
                        word.WindowState = 0
                    except Exception:
                        pass
                    return "Revealed"

                async def _task_reveal_active():
                    return await asyncio.to_thread(_do_reveal_active)

                return await with_global_lock(_task_reveal_active)
        except Exception as e:
            return f"ERROR: reveal_word failed: {e}"

    return mcp
