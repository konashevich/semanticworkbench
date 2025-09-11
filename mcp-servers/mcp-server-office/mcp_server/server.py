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
    set_range_font,
    set_range_paragraph,
    apply_list_format,
    apply_style_to_range,
    update_normal_style,
    apply_to_scope,
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
    async def highlight_text(text: str, color: str | int = "yellow", path: str = "") -> str:
        """Highlight the first occurrence of the given text in the active Word document.

        Args:
            text: The exact text (case sensitive) to highlight.
            color: One of the Word highlight colors (name or index). Use 0/none/clear to remove highlight.
            path: Optional absolute path to a .docx file to target explicitly.
        Returns:
            Status message about the highlight result.
        """
        try:
            from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
            import asyncio

            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"
                hci = _parse_highlight_color(color)

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
                    found = rng.Find.Execute(FindText=text, MatchCase=True, MatchWholeWord=False, Wrap=0, Forward=True)
                    if not found:
                        if AUTO_CLOSE:
                            try:
                                target.Close(SaveChanges=0)
                            except Exception:
                                pass
                        return "Text not found"
                    rng.HighlightColorIndex = hci
                    try:
                        target.Save()
                    except Exception:
                        pass
                    msg = "Highlighted"
                    if AUTO_CLOSE:
                        try:
                            target.Close(SaveChanges=0)
                        except Exception:
                            pass
                    return msg

                async def _task_hl_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                return await with_doc_lock(canonical_doc_key(str(p)), _task_hl_path)
            else:
                hci = _parse_highlight_color(color)

                def _do_highlight():
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                    word = get_word_app()
                    if word.Documents.Count == 0:
                        return "ERROR: No document is open. Provide 'path' or open one."
                    if word.Documents.Count > 1:
                        return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                    doc = get_active_document(word)
                    try:
                        if getattr(doc, "ReadOnly", False):
                            return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
                    except Exception:
                        pass
                    rng = doc.Content
                    rng.Find.ClearFormatting()
                    found = rng.Find.Execute(FindText=text, MatchCase=True, MatchWholeWord=False, Wrap=0, Forward=True)
                    if not found:
                        return "Text not found"
                    rng.HighlightColorIndex = hci
                    return "Highlighted"

                async def _task_hl_active():
                    return await asyncio.to_thread(_do_highlight)

                return await with_global_lock(_task_hl_active)
        except Exception as e:
            return f"ERROR: highlight_text failed: {e}"

    @mcp.tool()
    async def highlight_all(text: str, case_sensitive: bool = False, whole_word: bool = False, max_matches: int = 0, color: str | int = "yellow", path: str = "") -> str:
        """Highlight every occurrence of a string in the active Word document.

        Args:
            text: The text to search for (must be non-empty).
            case_sensitive: Match case exactly if True.
            whole_word: Match whole words only if True.
            max_matches: 0 = no limit, otherwise stop after this many highlights.
            color: One of the Word highlight colors (name or index). Use 0/none/clear to remove highlight.
            path: Optional absolute path to a .docx file to target explicitly.

        Returns:
            Summary string with number of highlights applied or an error message.
        """
        if not text.strip():
            return "ERROR: highlight_all requires non-empty text"
        try:
            import asyncio
            if path:
                from mcp_server.path_utils import resolve_user_path
                p = resolve_user_path(path)
                if not p.is_file():
                    return f"ERROR: File not found: {p}"
                hci = _parse_highlight_color(color)

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
                        if max_matches and count >= max_matches:
                            break
                        start_pos = current_end
                        if start_pos >= target.Content.End:
                            break
                        rng = target.Range(start_pos, target.Content.End)
                        rng.Find.ClearFormatting()
                        found = rng.Find.Execute(**flags)
                        if found and rng.End <= current_end:
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

                async def _task_hla_path():
                    return await _run_on_worker_with_cleanup(p, _do_on_worker)

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_hla_path)
            else:
                hci = _parse_highlight_color(color)
                def _do_all_active():
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
                        if max_matches and count >= max_matches:
                            break
                        start_pos = current_end
                        if start_pos >= doc.Content.End:
                            break
                        rng = doc.Range(start_pos, doc.Content.End)
                        rng.Find.ClearFormatting()
                        found = rng.Find.Execute(**flags)
                        if found and rng.End <= current_end:
                            break
                    return count

                async def _task_hla_active():
                    return await asyncio.to_thread(_do_all_active)

                applied = await with_global_lock(_task_hla_active)
                if applied == -1:
                    return "ERROR: No document is open. Provide 'path' or open one."
                if applied == -2:
                    return "ERROR: Multiple documents are open. Provide 'path' to target the right file."
                if applied == -3:
                    return "ERROR: Active document is read-only. Provide 'path' to edit a writable copy."
            if applied == 0:
                return "No matches highlighted"
            return f"Highlighted {applied} occurrence(s)"
        except Exception as e:
            return f"ERROR: highlight_all failed: {e}"

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
                        set_range_font(
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
                        set_range_font(
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
                        set_range_paragraph(
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
                        set_range_paragraph(
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
                        apply_list_format(rng, lt)

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
                        apply_list_format(rng, lt)

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

                    applied = 0

                    def _apply(rng):
                        nonlocal applied
                        if apply_style_to_range(rng, style_name):
                            applied += 1

                    # Use selection/document/matches
                    _ = apply_to_scope(
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

                applied = await with_doc_lock(canonical_doc_key(str(p)), _task_applystyle_path)
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
                    applied = 0

                    def _apply(rng):
                        nonlocal applied
                        if apply_style_to_range(rng, style_name):
                            applied += 1

                    _ = apply_to_scope(
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

                applied = await with_global_lock(_task_applystyle_active)
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
                        clear_direct_fmt(rng)

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
                        clear_direct_fmt(rng)

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
