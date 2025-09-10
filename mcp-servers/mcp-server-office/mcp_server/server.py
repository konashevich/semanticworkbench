# Copyright (c) Microsoft. All rights reserved.

from mcp.server.fastmcp import Context, FastMCP
import asyncio
from pathlib import Path

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
)
from mcp_server.markdown_edit.comment_analysis import run_comment_analysis
from mcp_server.markdown_edit.feedback_step import run_feedback_step
from mcp_server.markdown_edit.markdown_edit import run_markdown_edit
from mcp_server.types import MarkdownEditRequest

server_name = "Office MCP Server"


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
            p = Path(path).expanduser()
            if not p.is_file():
                return f"ERROR: File not found: {p}"
            if p.suffix.lower() != ".docx":
                return f"ERROR: Only .docx files supported. Got: {p.suffix}"
            # Open Word and the document in a background thread (COM is blocking)
            from mcp_server.app_interaction.word_editor import get_word_app, get_markdown_representation
            def _open():
                # Ensure COM initialized in this worker thread
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                word = get_word_app()
                # If already open, try to find matching document
                for i in range(1, word.Documents.Count + 1):
                    try:
                        d = word.Documents(i)
                        if Path(d.FullName) == p:
                            return d, word
                    except Exception:
                        continue
                doc = word.Documents.Open(str(p))  # type: ignore[attr-defined]
                return doc, word

            doc, _word = await asyncio.to_thread(_open)
            return await asyncio.to_thread(get_markdown_representation, doc)
        except Exception as e:
            return f"ERROR: Failed to open document: {e}"

    @mcp.tool()
    async def edit_word_document(task: str, ctx: Context) -> str:
        """Edits the active Word document according to the given task and returns a summary.

        Provide only the task instructions (a few sentences). The current document content is fetched automatically.
        """
        try:
            markdown_edit_output = await run_markdown_edit(
                markdown_edit_request=MarkdownEditRequest(context=ctx, task=task)
            )
            output_string = markdown_edit_output.change_summary + "\n" + markdown_edit_output.output_message
            return output_string
        except Exception as e:
            return f"ERROR: edit_word_document failed: {e}"

    @mcp.tool()
    async def add_comments_to_word_document(ctx: Context) -> str:
        """
        Runs a routine that will add feedback as comments to the currently open Word Document.
        """
        try:
            comment_output = await run_feedback_step(
                markdown_edit_request=MarkdownEditRequest(context=ctx),
            )
            return comment_output.feedback_summary
        except Exception as e:
            return f"ERROR: add_comments_to_word_document failed: {e}"

    @mcp.tool()
    async def analyze_comments(ctx: Context) -> str:
        """
        Runs a routine that analyze the comments in the Word document and determine how they could be solved.
        """
        try:
            comment_analysis_output = await run_comment_analysis(
                markdown_edit_request=MarkdownEditRequest(context=ctx),
            )
            return comment_analysis_output.edit_instructions + "\n" + comment_analysis_output.assistant_hints
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
            word = get_word_app()
            doc = get_active_document(word)
            markdown_from_word = get_markdown_representation(doc)
            return markdown_from_word
        except Exception as e:
            return f"ERROR: get_word_content failed: {e}"

    @mcp.tool()
    async def highlight_text(text: str) -> str:
        """Highlight the first occurrence of the given text in the active Word document.

        Args:
            text: The exact text (case sensitive) to highlight.
        Returns:
            Status message about the highlight result.
        """
        try:
            from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
            import asyncio

            def _do_highlight():
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                word = get_word_app()
                doc = get_active_document(word)
                rng = doc.Content
                rng.Find.ClearFormatting()
                found = rng.Find.Execute(FindText=text, MatchCase=True, MatchWholeWord=False)
                if not found:
                    return "Text not found"
                # rng is now the found range
                rng.HighlightColorIndex = 7  # wdYellow
                return "Highlighted"

            result = await asyncio.to_thread(_do_highlight)
            return result
        except Exception as e:
            return f"ERROR: highlight_text failed: {e}"

    @mcp.tool()
    async def highlight_all(text: str, case_sensitive: bool = False, whole_word: bool = False, max_matches: int = 0) -> str:
        """Highlight every occurrence of a string in the active Word document.

        Args:
            text: The text to search for (must be non-empty).
            case_sensitive: Match case exactly if True.
            whole_word: Match whole words only if True.
            max_matches: 0 = no limit, otherwise stop after this many highlights.

        Returns:
            Summary string with number of highlights applied or an error message.
        """
        if not text.strip():
            return "ERROR: highlight_all requires non-empty text"
        try:
            import asyncio
            def _do_all():
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
                word = get_word_app()
                doc = get_active_document(word)
                rng = doc.Content
                rng.Find.ClearFormatting()
                flags = {
                    'FindText': text,
                    'MatchCase': case_sensitive,
                    'MatchWholeWord': whole_word,
                }
                count = 0
                found = rng.Find.Execute(**flags)
                while found:
                    try:
                        rng.HighlightColorIndex = 7  # wdYellow
                        count += 1
                        if max_matches and count >= max_matches:
                            break
                        # Move search start to end of current match
                        start_pos = rng.End
                        rng = doc.Range(start_pos, doc.Content.End)
                        rng.Find.ClearFormatting()
                        found = rng.Find.Execute(**flags)
                    except Exception:
                        break
                return count
            applied = await asyncio.to_thread(_do_all)
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

            def _do_replace():
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                from mcp_server.app_interaction.word_editor import get_word_app, get_active_document
                word = get_word_app()
                doc = get_active_document(word)
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
                }
                count = 0
                found = rng.Find.Execute(**flags)
                while found:
                    count += 1
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
                return (count, None)

            count, err = await asyncio.to_thread(_do_replace)
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

    return mcp
