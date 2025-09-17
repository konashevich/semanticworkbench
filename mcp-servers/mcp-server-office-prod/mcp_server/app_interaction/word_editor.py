# Copyright (c) Microsoft. All rights reserved.


import sys
from contextlib import suppress

def _ensure_com_initialized():
    """Idempotently initialize COM for the current thread.

    Word (and other Office apps) require an STA. Repeated calls are safe.
    """
    try:
        import pythoncom  # type: ignore

        # CoInitialize is equivalent to CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)
        with suppress(Exception):
            pythoncom.CoInitialize()
    except Exception:
        # If pythoncom isn't available or initialization fails, we let later code raise a clearer error.
        pass

from mcp_server.types import WordCommentData


def get_word_app():
    if sys.platform != "win32":
        raise EnvironmentError("This script only works on Windows.")

    import win32com.client as win32

    _ensure_com_initialized()

    """Connect to Word if it is running, or start a new instance."""
    try:
        # Try connecting to an existing instance of Word
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        # If not running, create a new instance
        word = win32.Dispatch("Word.Application")
    # Keep Word hidden by default to prevent empty windows from appearing during headless operations
    # (server tools will show or target documents explicitly when needed)
    word.Visible = False
    return word


def get_active_document(word):
    """Return the active Word document. Does not create a new one if none is open."""
    if word.Documents.Count == 0:
        raise RuntimeError("No active document")
    return word.ActiveDocument


def get_document_content(doc):
    """Return the content of the document."""
    return doc.Content.Text


def replace_document_content(doc, content):
    """Replace the content of the document with the given content."""
    doc.Content.Text = content


def get_markdown_representation(doc, include_comments: bool = False) -> str:
    """
    Get the markdown representation of the document.
    Supports Headings, plaintext, bulleted/numbered lists, bold, italic, and code blocks.
    """
    markdown_text = []
    in_code_block = False
    for i in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(i)
        style_name = paragraph.Style.NameLocal

        # Handle Code style for code blocks
        if style_name == "Code":
            if not in_code_block:
                markdown_text.append("```")
                in_code_block = True
            markdown_text.append(paragraph.Range.Text.rstrip())
            continue
        elif in_code_block:
            # Close code block when style changes
            markdown_text.append("```")
            in_code_block = False

        # Process paragraph style first
        prefix = ""
        if "Heading" in style_name:
            try:
                level = int(style_name.split("Heading")[1].strip())
                prefix = "#" * level + " "
            except (ValueError, IndexError):
                pass

        para_text = ""
        para_range = paragraph.Range
        # For performance, check if there's any formatting at all
        if para_range.Text.strip():
            if para_range.Font.Bold or para_range.Font.Italic:
                # Process words instead of characters for better performance
                current_run = {"text": "", "bold": False, "italic": False}

                # Get all words in this paragraph
                for w in range(1, para_range.Words.Count + 1):
                    word_range = para_range.Words(w)
                    word_text = word_range.Text  # Keep original with potential spaces

                    # Skip if empty
                    if not word_text.strip():
                        continue

                    # Get formatting for this word
                    is_bold = word_range.Font.Bold
                    is_italic = word_range.Font.Italic

                    # If formatting changed, start a new run
                    if is_bold != current_run["bold"] or is_italic != current_run["italic"]:
                        # Finish the previous run if it exists
                        if current_run["text"]:
                            if current_run["bold"] and current_run["italic"]:
                                para_text += f"***{current_run['text'].rstrip()}***"
                            elif current_run["bold"]:
                                para_text += f"**{current_run['text'].rstrip()}**"
                            elif current_run["italic"]:
                                para_text += f"*{current_run['text'].rstrip()}*"
                            else:
                                para_text += current_run["text"].rstrip()

                            # Add a space if the previous run ended with a space
                            if current_run["text"].endswith(" "):
                                para_text += " "

                        # Start a new run with the current word
                        current_run = {"text": word_text, "bold": is_bold, "italic": is_italic}
                    else:
                        # Continue the current run - but be careful with spaces
                        if current_run["text"]:
                            current_run["text"] += word_text
                        else:
                            current_run["text"] = word_text

                # Process the final run
                if current_run["text"]:
                    if current_run["bold"] and current_run["italic"]:
                        para_text += f"***{current_run['text'].rstrip()}***"
                    elif current_run["bold"]:
                        para_text += f"**{current_run['text'].rstrip()}**"
                    elif current_run["italic"]:
                        para_text += f"*{current_run['text'].rstrip()}*"
                    else:
                        para_text += current_run["text"].rstrip()
            else:
                # No special formatting, just get the text
                para_text = para_range.Text.strip()
        else:
            para_text = para_range.Text.strip()

        if not para_text:
            continue

        # Handle list formatting
        if paragraph.Range.ListFormat.ListType == 2:
            markdown_text.append(f"- {para_text}")
        elif paragraph.Range.ListFormat.ListType == 3:
            markdown_text.append(f"1. {para_text}")
        else:
            markdown_text.append(f"{prefix}{para_text}")

    # Close any open code block at the end of document
    if in_code_block:
        markdown_text.append("```")

    if include_comments:
        comment_section = get_comments_markdown_representation(doc)
        if comment_section:
            markdown_text.append(comment_section)

    return "\n".join(markdown_text)


def _write_formatted_text(selection, text):
    """
    Helper function to write text with markdown formatting (bold, italic) to the document.
    Processes text in chunks for better performance.

    Args:
        selection: Word selection object where text will be inserted
        text: Markdown-formatted text string to process
    """

    segments = []
    i = 0
    while i < len(text):
        # Bold+Italic (***text***)
        if i + 2 < len(text) and text[i : i + 3] == "***" and "***" in text[i + 3 :]:
            end_pos = text.find("***", i + 3)
            if end_pos != -1:
                segments.append(("bold_italic", text[i + 3 : end_pos]))
                i = end_pos + 3
                continue

        # Bold (**text**)
        elif i + 1 < len(text) and text[i : i + 2] == "**" and "**" in text[i + 2 :]:
            end_pos = text.find("**", i + 2)
            if end_pos != -1:
                segments.append(("bold", text[i + 2 : end_pos]))
                i = end_pos + 2
                continue

        # Italic (*text*)
        elif text[i] == "*" and i + 1 < len(text) and text[i + 1] != "*" and "*" in text[i + 1 :]:
            end_pos = text.find("*", i + 1)
            if end_pos != -1:
                segments.append(("italic", text[i + 1 : end_pos]))
                i = end_pos + 1
                continue

        # Find the next special marker or end of string
        next_marker = float("inf")
        for marker in ["***", "**", "*"]:
            pos = text.find(marker, i)
            if pos != -1 and pos < next_marker:
                next_marker = pos

        # Add plain text segment
        if next_marker == float("inf"):
            segments.append(("plain", text[i:]))
            break
        else:
            segments.append(("plain", text[i:next_marker]))
            i = next_marker

    # Now write all segments with minimal formatting changes
    current_format = None
    for format_type, content in segments:
        if format_type != current_format:
            selection.Font.Bold = False
            selection.Font.Italic = False

            if format_type == "bold" or format_type == "bold_italic":
                selection.Font.Bold = True
            if format_type == "italic" or format_type == "bold_italic":
                selection.Font.Italic = True

            current_format = format_type
        selection.TypeText(content)

    selection.Font.Bold = False
    selection.Font.Italic = False


def write_markdown_to_document(doc, markdown_text: str) -> None:
    """Writes markdown text to a Word document with appropriate formatting.

    Converts markdown syntax to Word formatting, including:
    - Headings (# to Heading styles)
    - Lists (bulleted and numbered)
    - Text formatting (bold, italic)
    - Code blocks (``` to Code style)
    """
    comments = get_document_comments(doc)

    doc.Content.Delete()

    word_app = doc.Application
    selection = word_app.Selection

    # Create "Code" style if it doesn't exist
    try:
        # Check if Code style exists
        code_style = word_app.ActiveDocument.Styles("Code")
    except Exception:
        # Create the Code style
        code_style = word_app.ActiveDocument.Styles.Add("Code", 1)
        code_style.Font.Name = "Cascadia Code"
        code_style.Font.Size = 10
        code_style.ParagraphFormat.SpaceAfter = 0
        code_style.QuickStyle = True
        code_style.LinkStyle = True

    # This fixes an issue where if there are comments on a doc, there is no selection
    # which causes insertion to fail
    doc.Range(0, 0).Select()

    # Ensure we start with normal style
    selection.Style = word_app.ActiveDocument.Styles("Normal")

    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1

        if not line:
            continue

        if line.startswith("```"):
            i_start = i

            # Find the end of the code block
            while i < len(lines) and not lines[i].strip().startswith("```"):
                i += 1

            # Process all lines in the code block
            for j in range(i_start, i):
                code_line = lines[j]
                selection.TypeText(code_line)
                selection.Style = word_app.ActiveDocument.Styles("Code")
                selection.TypeParagraph()

            # Skip the closing code fence
            if i < len(lines):
                i += 1

            # Restore normal style for next paragraph
            selection.Style = word_app.ActiveDocument.Styles("Normal")
            continue

        # Check if the line is a heading
        if line.startswith("#"):
            heading_level = 0
            for char in line:
                if char == "#":
                    heading_level += 1
                else:
                    break

            # Remove the # characters and any leading space
            text = line[heading_level:].strip()

            _write_formatted_text(selection, text)

            # Get the current paragraph and set its style
            current_paragraph = selection.Paragraphs.Last
            if 1 <= heading_level <= 9:  # Word supports Heading 1-9
                current_paragraph.Style = f"Heading {heading_level}"

            selection.TypeParagraph()

        # Check if line is a bulleted list item
        elif line.startswith(("- ", "* ")):
            # Extract the text after the bullet marker
            text = line[2:].strip()

            _write_formatted_text(selection, text)

            # Apply bullet formatting
            selection.Range.ListFormat.ApplyBulletDefault()

            selection.TypeParagraph()
            selection.Style = word_app.ActiveDocument.Styles("Normal")

        # Check if line is a numbered list item
        elif line.strip().startswith(tuple(f"{i}. " for i in range(1, 100)) + tuple(f"{i})" for i in range(1, 100))):
            # Extract the text after the number and period/parenthesis
            text = ""
            if ". " in line:
                parts = line.strip().split(". ", 1)
                if len(parts) > 1:
                    text = parts[1]
            elif ") " in line:
                parts = line.strip().split(") ", 1)
                if len(parts) > 1:
                    text = parts[1]

            if text:
                _write_formatted_text(selection, text)

                # Apply numbered list formatting
                selection.Range.ListFormat.ApplyNumberDefault()
                selection.TypeParagraph()
                selection.Style = word_app.ActiveDocument.Styles("Normal")
            else:
                # If parsing failed, just add the line as normal text
                _write_formatted_text(selection, line)
                selection.TypeParagraph()

        else:
            # Regular paragraph text with formatting support
            _write_formatted_text(selection, line)
            selection.TypeParagraph()

    # Move cursor to the beginning of the document
    doc.Range(0, 0).Select()

    # Reapply comments. Note this implicitly will remove any comments that have locations not in the new text.
    for comment in comments:
        add_document_comment(doc, comment)


def add_document_comment(
    doc,
    comment_data: WordCommentData,
) -> bool:
    """
    Add a comment to specific text within a Word document.

    Returns:
        bool: True if comment was added successfully, False otherwise
    """
    try:
        content_range = doc.Content
        found_range = None

        # Find the specified occurrence of the text
        found_count = 0

        # Use FindText to locate the text
        content_range.Find.ClearFormatting()
        found = content_range.Find.Execute(FindText=comment_data.location_text, MatchCase=True, MatchWholeWord=False)

        while found:
            found_count += 1
            if found_count == comment_data.occurrence:
                found_range = content_range.Duplicate
                break

            # Continue searching from the end of the current match
            content_range.Collapse(Direction=0)  # Collapse to end
            found = content_range.Find.Execute(
                FindText=comment_data.location_text, MatchCase=True, MatchWholeWord=False
            )

        if not found_range:
            return False

        # Add a comment to the found range
        comment = doc.Comments.Add(Range=found_range, Text=comment_data.comment_text)
        if comment_data.author:
            comment.Author = comment_data.author
        return True
    except Exception:
        return False


def get_document_comments(doc) -> list[WordCommentData]:
    """
    Retrieve all comments from a Word document.
    """
    comments: list[WordCommentData] = []

    try:
        if doc.Comments.Count == 0:
            return comments

        for i in range(1, doc.Comments.Count + 1):
            try:
                comment = doc.Comments(i)

                comment_text = ""
                try:
                    comment_text = comment.Range.Text
                except Exception:
                    pass

                author = "Unknown"
                try:
                    author = comment.Author
                except Exception:
                    pass

                date = ""
                try:
                    date = str(comment.Date)
                except Exception:
                    pass

                reference_text = ""
                try:
                    if hasattr(comment, "Scope"):
                        reference_text = comment.Scope.Text
                except Exception:
                    pass

                comment_info = WordCommentData(
                    id=str(i),
                    comment_text=comment_text,
                    location_text=reference_text,
                    date=date,
                    author=author,
                )
                comments.append(comment_info)
            except Exception:
                continue

        return comments
    except Exception as e:
        print(f"Error retrieving comments: {e}")
        return comments


def get_comments_markdown_representation(doc) -> str:
    comments = get_document_comments(doc)

    if not comments:
        return ""

    comment_section = "\n\n<comments>\n"
    for i, comment in enumerate(comments, 1):
        comment_section += f'<comment id={i} author="{comment.author}">\n'
        comment_section += f"  <location_text>{comment.location_text}</location_text>\n"
        comment_section += f"  <comment_text>{comment.comment_text}</comment_text>\n"
        comment_section += "</comment>\n"
    comment_section.rstrip()
    comment_section += "</comments>"
    return comment_section


def delete_comments_containing_text(doc, search_text: str, case_sensitive: bool = False) -> int:
    """
    Delete comments containing specific text.

    Args:
        doc: Word document object
        search_text: Text to search for in comments
        case_sensitive: Whether the search should be case-sensitive

    Returns:
        int: Number of comments deleted
    """
    deleted_count = 0
    try:
        if not case_sensitive:
            search_text = search_text.lower()

        # Work backwards to avoid index shifting issues when deleting
        for i in range(doc.Comments.Count, 0, -1):
            comment = doc.Comments(i)
            comment_text = comment.Range.Text

            if not case_sensitive:
                comment_text = comment_text.lower()

            if search_text in comment_text:
                comment.Delete()
                deleted_count += 1
        return deleted_count
    except Exception:
        return deleted_count


# region Formatting helpers

def _rgb_from_hex(color: str) -> int:
    """Parse #RRGGBB to Word RGB int (r + g*256 + b*65536)."""
    s = color.strip()
    if s.startswith("#"):
        s = s[1:]
    if len(s) != 6:
        raise ValueError("Invalid hex color")
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return r + g * 256 + b * 65536


_NAMED_COLORS = {
    # common names
    "black": _rgb_from_hex("#000000"),
    "white": _rgb_from_hex("#ffffff"),
    "red": _rgb_from_hex("#ff0000"),
    "green": _rgb_from_hex("#008000"),
    "blue": _rgb_from_hex("#0000ff"),
    "yellow": _rgb_from_hex("#ffff00"),
    "gray": _rgb_from_hex("#808080"),
    "grey": _rgb_from_hex("#808080"),
    "lightgray": _rgb_from_hex("#d3d3d3"),
    "lightgrey": _rgb_from_hex("#d3d3d3"),
    "darkgray": _rgb_from_hex("#a9a9a9"),
    "darkgrey": _rgb_from_hex("#a9a9a9"),
}


def parse_font_color(color: str | int | None) -> int | None:
    """Resolve a color value for Word Font.Color.

    Accepts:
      - None: no change
      - int: pass-through
      - named color string
      - #RRGGBB
    """
    if color is None:
        return None
    if isinstance(color, int):
        return int(color)
    s = str(color).strip()
    if not s:
        return None
    if s.startswith("#"):
        try:
            return _rgb_from_hex(s)
        except Exception:
            return None
    return _NAMED_COLORS.get(s.lower(), None)


def get_selection(word):
    """Return the Application.Selection if available, else None."""
    try:
        sel = word.Selection
        # Some situations can have no selection object; ensure it has a Range
        _ = sel.Range  # may raise
        return sel
    except Exception:
        return None


def iter_find_ranges(doc, find_text: str, *, match_case: bool = False, whole_word: bool = False, max_matches: int = 0):
    """Yield non-overlapping ranges that match find_text in the document, forward-only."""
    if not find_text:
        return
    rng = doc.Content
    rng.Find.ClearFormatting()
    flags = {
        "FindText": find_text,
        "MatchCase": match_case,
        "MatchWholeWord": whole_word,
        "Wrap": 0,  # wdFindStop
        "Forward": True,
    }
    count = 0
    found = rng.Find.Execute(**flags)
    while found:
        current_end = rng.End
        yield rng.Duplicate
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


def set_range_font(
    rng,
    *,
    name: str | None = None,
    size: float | int | None = None,
    color: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    superscript: bool | None = None,
    subscript: bool | None = None,
    ) -> bool:
    """Apply character-level formatting to a range. Returns True if any change applied."""
    changed = False
    try:
        f = rng.Font
        if name and getattr(f, "Name", None) != name:
            f.Name = name
            changed = True
        if size is not None:
            try:
                if float(getattr(f, "Size", 0)) != float(size):
                    f.Size = float(size)
                    changed = True
            except Exception:
                pass
        if color is not None:
            try:
                if int(getattr(f, "Color", -1)) != int(color):
                    f.Color = int(color)
                    changed = True
            except Exception:
                pass
        if bold is not None and bool(getattr(f, "Bold", False)) != bool(bold):
            f.Bold = bool(bold)
            changed = True
        if italic is not None and bool(getattr(f, "Italic", False)) != bool(italic):
            f.Italic = bool(italic)
            changed = True
        if underline is not None:
            target = 1 if underline else 0
            try:
                if int(getattr(f, "Underline", 0)) != target:
                    f.Underline = target  # wdUnderlineSingle=1, wdUnderlineNone=0
                    changed = True
            except Exception:
                pass
        if strikethrough is not None:
            try:
                if bool(getattr(f, "StrikeThrough", False)) != bool(strikethrough):
                    f.StrikeThrough = bool(strikethrough)
                    changed = True
            except Exception:
                pass
        # Superscript and subscript are mutually exclusive at the Font level
        if superscript is not None:
            try:
                if bool(getattr(f, "Superscript", False)) != bool(superscript):
                    f.Superscript = bool(superscript)
                    changed = True
                if superscript:
                    with suppress(Exception):
                        if bool(getattr(f, "Subscript", False)):
                            f.Subscript = False
                            changed = True
            except Exception:
                pass
        if subscript is not None:
            try:
                if bool(getattr(f, "Subscript", False)) != bool(subscript):
                    f.Subscript = bool(subscript)
                    changed = True
                if subscript:
                    with suppress(Exception):
                        if bool(getattr(f, "Superscript", False)):
                            f.Superscript = False
                            changed = True
            except Exception:
                pass
    except Exception:
        pass
    return changed


_ALIGN_MAP = {
    "left": 0,  # wdAlignParagraphLeft
    "center": 1,  # wdAlignParagraphCenter
    "right": 2,  # wdAlignParagraphRight
    "justify": 3,  # wdAlignParagraphJustify
}

_LINE_SPACING_RULE = {
    "single": (0, 0.0),  # (rule, pt)
    "onepointfive": (1, 0.0),
    "1.5": (1, 0.0),
    "double": (2, 0.0),
}


def _parse_line_spacing(value: str) -> tuple[int, float] | None:
    if not value:
        return None
    s = value.strip().lower()
    if s in _LINE_SPACING_RULE:
        return _LINE_SPACING_RULE[s]
    # exact:14 or atleast:12
    if s.startswith("exact:"):
        try:
            pt = float(s.split(":", 1)[1])
            return (4, pt)  # wdLineSpaceExactly
        except Exception:
            return None
    if s.startswith("atleast:"):
        try:
            pt = float(s.split(":", 1)[1])
            return (3, pt)  # wdLineSpaceAtLeast
        except Exception:
            return None
    return None


def set_range_paragraph(rng, *, alignment: str | None = None, line_spacing: str | None = None, space_before: float | int | None = None, space_after: float | int | None = None) -> bool:
    changed = False
    try:
        p = rng.ParagraphFormat
        if alignment:
            val = _ALIGN_MAP.get(str(alignment).lower())
            if val is not None and getattr(p, "Alignment", None) != val:
                p.Alignment = val
                changed = True
        if line_spacing:
            res = _parse_line_spacing(line_spacing)
            if res:
                rule, pt = res
                try:
                    if getattr(p, "LineSpacingRule", None) != rule:
                        p.LineSpacingRule = rule
                        changed = True
                except Exception:
                    pass
                if pt:
                    try:
                        if float(getattr(p, "LineSpacing", 0.0)) != float(pt):
                            p.LineSpacing = float(pt)
                            changed = True
                    except Exception:
                        pass
        if space_before is not None:
            try:
                if float(getattr(p, "SpaceBefore", 0.0)) != float(space_before):
                    p.SpaceBefore = float(space_before)
                    changed = True
            except Exception:
                pass
        if space_after is not None:
            try:
                if float(getattr(p, "SpaceAfter", 0.0)) != float(space_after):
                    p.SpaceAfter = float(space_after)
                    changed = True
            except Exception:
                pass
    except Exception:
        pass
    return changed


def apply_list_format(rng, list_type: str) -> bool:
    changed = False
    try:
        lf = rng.ListFormat
        t = str(list_type).lower()
        before = getattr(lf, "ListType", 0)
        if t == "bullet":
            with suppress(Exception):
                lf.ApplyBulletDefault()
            after = getattr(lf, "ListType", 0)
            changed = before != after
        elif t == "numbered":
            with suppress(Exception):
                lf.ApplyNumberDefault()
            after = getattr(lf, "ListType", 0)
            changed = before != after
        elif t == "none":
            # Best-effort remove any list formatting
            with suppress(Exception):
                lf.RemoveNumbers()
            with suppress(Exception):
                lf.RemoveBullets()
            after = getattr(lf, "ListType", 0)
            changed = before != after
    except Exception:
        pass
    return changed


def apply_style_to_range(rng, style_name: str) -> bool:
    """Apply a style idempotently, respecting style type.

    - Character style: apply to the entire range as a character style.
    - Paragraph style: apply to each paragraph in the range as needed.
    Falls back to range-level assignment if style metadata isn't available.
    """
    # Try to detect style type (1=paragraph, 2=character)
    style_type = None
    try:
        doc = getattr(rng, "Document", None)
        if doc is not None:
            try:
                s = doc.Styles(style_name)
                style_type = getattr(s, "Type", None)
            except Exception:
                pass
    except Exception:
        pass

    # Character style path
    if style_type == 2:
        try:
            current = None
            try:
                current = rng.Style
            except Exception:
                pass
            if str(current) != style_name:
                rng.Style = style_name
                return True
        except Exception:
            pass
        return False

    # Paragraph style (or unknown type): apply per paragraph where needed
    changed = False
    try:
        para = rng.Paragraphs
        if para is not None and para.Count > 0:
            for i in range(1, int(para.Count) + 1):
                try:
                    cur = None
                    try:
                        cur = para(i).Style
                    except Exception:
                        pass
                    if str(cur) != style_name:
                        para(i).Style = style_name
                        changed = True
                except Exception:
                    continue
            return changed
    except Exception:
        pass

    # Fallback: set on range if paragraphs not available
    try:
        current = None
        try:
            current = rng.Style
        except Exception:
            pass
        if str(current) != style_name:
            rng.Style = style_name
            return True
    except Exception:
        pass
    return False


def update_normal_style(doc, *, name: str | None = None, size: float | int | None = None, color: int | None = None, bold: bool | None = None, italic: bool | None = None, underline: bool | None = None) -> bool:
    try:
        styles = doc.Styles
        normal = styles("Normal")
        f = normal.Font
        changed = False
        if name and getattr(f, "Name", None) != name:
            f.Name = name
            changed = True
        if size is not None:
            try:
                if float(getattr(f, "Size", 0)) != float(size):
                    f.Size = float(size)
                    changed = True
            except Exception:
                pass
        if color is not None:
            try:
                if int(getattr(f, "Color", -1)) != int(color):
                    f.Color = int(color)
                    changed = True
            except Exception:
                pass
        if bold is not None and bool(getattr(f, "Bold", False)) != bool(bold):
            f.Bold = bool(bold)
            changed = True
        if italic is not None and bool(getattr(f, "Italic", False)) != bool(italic):
            f.Italic = bool(italic)
            changed = True
        if underline is not None:
            target = 1 if underline else 0
            try:
                if int(getattr(f, "Underline", 0)) != target:
                    f.Underline = target
                    changed = True
            except Exception:
                pass
        return changed
    except Exception:
        return False

def _snapshot_font(f):
    """Return a tuple snapshot of key Font properties for idempotence checks."""
    try:
        return (
            str(getattr(f, "Name", "")),
            float(getattr(f, "Size", 0.0) or 0.0),
            int(getattr(f, "Color", -1) or -1),
            int(getattr(f, "ColorIndex", -1) or -1),
            bool(getattr(f, "Bold", False)),
            bool(getattr(f, "Italic", False)),
            int(getattr(f, "Underline", 0) or 0),
            bool(getattr(f, "StrikeThrough", False)),
            bool(getattr(f, "Superscript", False)),
            bool(getattr(f, "Subscript", False)),
            bool(getattr(f, "Hidden", False)),
            bool(getattr(f, "SmallCaps", False)),
            bool(getattr(f, "AllCaps", False)),
            bool(getattr(f, "Shadow", False)),
            bool(getattr(f, "Outline", False)),
            bool(getattr(f, "Emboss", False)),
            bool(getattr(f, "Engrave", False)),
            float(getattr(f, "Kerning", 0.0) or 0.0),
            float(getattr(f, "Spacing", 0.0) or 0.0),
            int(getattr(f, "Scaling", 100) or 100),
            float(getattr(f, "Position", 0.0) or 0.0),
        )
    except Exception:
        return None

def _snapshot_par(p):
    """Return a tuple snapshot of key ParagraphFormat properties for idempotence checks."""
    try:
        tabs: list[tuple[float, int, int]] = []
        try:
            ts = getattr(p, "TabStops", None)
            if ts is not None:
                count = int(getattr(ts, "Count", 0) or 0)
                for i in range(1, count + 1):
                    try:
                        stop = ts(i)
                        pos = float(getattr(stop, "Position", 0.0) or 0.0)
                        align = int(getattr(stop, "Alignment", 0) or 0)
                        leader = int(getattr(stop, "Leader", 0) or 0)
                        tabs.append((pos, align, leader))
                    except Exception:
                        continue
        except Exception:
            pass
        # Shading (background/foreground/texture)
        shading = (
            int(getattr(getattr(p, "Shading", None), "BackgroundPatternColor", -1) or -1),
            int(getattr(getattr(p, "Shading", None), "ForegroundPatternColor", -1) or -1),
            int(getattr(getattr(p, "Shading", None), "Texture", 0) or 0),
        )
        # Borders: capture enabled flag only (deep border state diff would be heavy)
        try:
            borders_enabled = bool(getattr(getattr(p, "Borders", None), "Enable", False))
        except Exception:
            borders_enabled = False
        return (
            int(getattr(p, "Alignment", 0) or 0),
            int(getattr(p, "LineSpacingRule", 0) or 0),
            float(getattr(p, "LineSpacing", 0.0) or 0.0),
            float(getattr(p, "SpaceBefore", 0.0) or 0.0),
            float(getattr(p, "SpaceAfter", 0.0) or 0.0),
            float(getattr(p, "LeftIndent", 0.0) or 0.0),
            float(getattr(p, "RightIndent", 0.0) or 0.0),
            float(getattr(p, "FirstLineIndent", 0.0) or 0.0),
            bool(getattr(p, "WidowControl", False)),
            bool(getattr(p, "KeepWithNext", False)),
            bool(getattr(p, "KeepTogether", False)),
            int(getattr(p, "OutlineLevel", 0) or 0),
            bool(getattr(p, "PageBreakBefore", False)),
            bool(getattr(p, "NoLineNumber", False)),
            bool(getattr(p, "Hyphenation", False)),
            bool(getattr(p, "MirrorIndents", False)),
            bool(getattr(p, "DisableLineHeightGrid", False)),
            bool(getattr(p, "AutoAdjustRightIndent", False)),
            float(getattr(p, "CharacterUnitLeftIndent", 0.0) or 0.0),
            float(getattr(p, "CharacterUnitRightIndent", 0.0) or 0.0),
            float(getattr(p, "CharacterUnitFirstLineIndent", 0.0) or 0.0),
            tuple(tabs),
            shading,
            borders_enabled,
        )
    except Exception:
        return None

def clear_direct_formatting(rng, target: str = "all") -> bool:
    """Clear direct formatting, returning True only if anything actually changed.

    target: 'all' | 'font' | 'paragraph'
    """
    changed = False
    try:
        t = str(target).lower().strip() if target else "all"
        font_before = _snapshot_font(getattr(rng, "Font", None)) if t in ("all", "font") else None
        par_before = _snapshot_par(getattr(rng, "ParagraphFormat", None)) if t in ("all", "paragraph") else None
        if t in ("all", "font"):
            with suppress(Exception):
                rng.Font.Reset()
        if t in ("all", "paragraph"):
            with suppress(Exception):
                rng.ParagraphFormat.Reset()
        font_after = _snapshot_font(getattr(rng, "Font", None)) if font_before is not None else None
        par_after = _snapshot_par(getattr(rng, "ParagraphFormat", None)) if par_before is not None else None
        if font_before is not None and font_after is not None and font_before != font_after:
            changed = True
        if par_before is not None and par_after is not None and par_before != par_after:
            changed = True
    except Exception:
        pass
    return changed

def apply_to_scope(word, doc, scope: str, *, find_text: str = "", match_case: bool = False, whole_word: bool = False, max_matches: int = 0, fn=None, on_applied=None) -> int:
    """Apply a function to ranges based on scope. Returns count applied."""
    applied = 0
    s = str(scope).strip().lower()
    try:
        if s == "document":
            if fn:
                ok = fn(doc.Content)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(doc.Content)
                applied = 1 if ok else 0
        elif s == "selection":
            sel = get_selection(word)
            if sel is None:
                return 0
            if fn:
                ok = fn(sel.Range)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(sel.Range)
                applied = 1 if ok else 0
        elif s == "matches":
            if not find_text:
                return 0
            for rng in iter_find_ranges(doc, find_text, match_case=match_case, whole_word=whole_word, max_matches=max_matches):
                if fn and fn(rng):
                    if on_applied:
                        with suppress(Exception):
                            on_applied(rng)
                    applied += 1
        else:
            # default: document
            if fn:
                ok = fn(doc.Content)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(doc.Content)
                applied = 1 if ok else 0
    except Exception:
        pass
    return applied


def apply_to_scope_strict(word, doc, scope: str, *, find_text: str = "", match_case: bool = False, whole_word: bool = False, max_matches: int = 0, fn=None, on_applied=None) -> int:
    """Apply a function to ranges based on scope with strict selection validation. Returns count applied."""
    applied = 0
    s = str(scope).strip().lower()
    try:
        if s == "document":
            if fn:
                ok = fn(doc.Content)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(doc.Content)
                applied = 1 if ok else 0
        elif s == "selection":
            sel = get_selection(word)
            if sel is None or sel.Range.Start == sel.Range.End:
                raise ValueError("Selection scope requires an active text selection")
            if fn:
                ok = fn(sel.Range)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(sel.Range)
                applied = 1 if ok else 0
        elif s == "matches":
            if not find_text:
                return 0
            for rng in iter_find_ranges(doc, find_text, match_case=match_case, whole_word=whole_word, max_matches=max_matches):
                if fn and fn(rng):
                    if on_applied:
                        with suppress(Exception):
                            on_applied(rng)
                    applied += 1
        else:
            # default: document
            if fn:
                ok = fn(doc.Content)
                if ok and on_applied:
                    with suppress(Exception):
                        on_applied(doc.Content)
                applied = 1 if ok else 0
    except Exception:
        pass
    return applied


# endregion
