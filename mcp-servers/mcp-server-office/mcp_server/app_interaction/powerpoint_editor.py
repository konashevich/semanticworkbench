# Copyright (c) Microsoft. All rights reserved.

import sys
from typing import Any, Optional
import re

POINTS_PER_INCH = 72.0
MM_PER_INCH = 25.4

# Color name mapping (same as Word editor)
def _rgb_from_hex(hex_str: str) -> int:
    hex_str = hex_str.lstrip("#")
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color: {hex_str}")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return (b << 16) + (g << 8) + r  # BGR for Office

_NAMED_COLORS = {
    "black": 0x000000,
    "white": 0xFFFFFF,
    "red": 0x0000FF,
    "green": 0x008000,
    "blue": 0xFF0000,
    "yellow": 0x00FFFF,
    "cyan": 0xFFFF00,
    "magenta": 0xFF00FF,
    "silver": 0xC0C0C0,
    "gray": 0x808080,
    "grey": 0x808080,
    "maroon": 0x000080,
    "olive": 0x008080,
    "lime": 0x00FF00,
    "aqua": 0xFFFF00,
    "teal": 0x808000,
    "navy": 0x800000,
    "fuchsia": 0xFF00FF,
    "purple": 0x800080,
    "orange": 0x0080FF,
    "darkred": 0x00008B,
    "darkgreen": 0x006400,
    "darkblue": 0x8B0000,
    "darkgray": 0xA9A9A9,
    "darkgrey": 0xA9A9A9,
}

# Validation helpers
class PPTValidationError(Exception):
    def __init__(self, code: str, message: str, details: dict | None = None):
        self.code = code
        self.message = message
        self.details = details or {}
        super().__init__(message)

def validate_coordinates(left: float, top: float, width: float, height: float, slide_width: float, slide_height: float) -> None:
    """Validate shape coordinates against slide bounds."""
    if left < 0 or top < 0:
        raise PPTValidationError("out-of-bounds", f"Negative coordinates not allowed: left={left}, top={top}")
    if width <= 0 or height <= 0:
        raise PPTValidationError("invalid-size", f"Width and height must be positive: width={width}, height={height}")
    if left + width > slide_width:
        raise PPTValidationError("out-of-bounds", f"Shape extends beyond slide width: {left + width} > {slide_width}")
    if top + height > slide_height:
        raise PPTValidationError("out-of-bounds", f"Shape extends beyond slide height: {top + height} > {slide_height}")

def validate_font_size(size: float) -> None:
    """Validate font size is within acceptable range."""
    if size < 1 or size > 1638:
        raise PPTValidationError("invalid-font-size", f"Font size must be between 1 and 1638 points, got {size}")

def validate_color(color: str | int) -> int:
    """Parse and validate color value."""
    if isinstance(color, int):
        return color
    s = str(color).strip().lower()
    if s.startswith("#"):
        try:
            return _rgb_from_hex(s)
        except Exception:
            raise PPTValidationError("invalid-color", f"Invalid hex color: {color}")
    if s in _NAMED_COLORS:
        return _NAMED_COLORS[s]
    raise PPTValidationError("invalid-color", f"Unknown color: {color}")

def round_precision(value: float, decimals: int = 3) -> float:
    """Round coordinate values to prevent floating point drift."""
    return round(value, decimals)


def mm_to_points(mm: float) -> float:
    return mm * POINTS_PER_INCH / MM_PER_INCH


def px_to_points(px: float, dpi: float = 96.0) -> float:
    return px * POINTS_PER_INCH / dpi


_UNIT_RE = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*(mm|pt|px)?\s*$", re.IGNORECASE)


def parse_unit(value: str | float | int, dpi: float = 96.0) -> float:
    """Parse a numeric value with optional unit suffix to PowerPoint points.

    Accepts raw numbers (assumed points) or strings ending with mm/pt/px.
    """
    if isinstance(value, (int, float)):
        return float(value)
    m = _UNIT_RE.match(str(value))
    if not m:
        raise PPTValidationError("invalid-unit", f"Invalid unit value: {value}")
    num = float(m.group(1))
    unit = (m.group(2) or "pt").lower()
    if unit == "pt":
        return num
    if unit == "mm":
        return mm_to_points(num)
    if unit == "px":
        return px_to_points(num, dpi=dpi)
    raise PPTValidationError("invalid-unit", f"Unsupported unit: {unit}")


def get_powerpoint_app():
    """Connect to PowerPoint if it is running, or start a new instance."""
    if sys.platform != "win32":
        raise EnvironmentError("This script only works on Windows.")

    import win32com.client as win32

    try:
        # Try connecting to an existing instance of PowerPoint
        powerpoint = win32.GetActiveObject("PowerPoint.Application")
    except Exception:
        # If not running, create a new instance
        powerpoint = win32.Dispatch("PowerPoint.Application")
    # Make PowerPoint visible
    powerpoint.Visible = True
    return powerpoint


def get_active_presentation(powerpoint):
    """Return an active PowerPoint presentation, or create one if none is open."""
    if powerpoint.Presentations.Count == 0:
        # If there are no presentations, add a new one
        return powerpoint.Presentations.Add()
    else:
        return powerpoint.ActivePresentation


def get_slide_content(presentation, slide_number):
    """Return the text content of a slide."""
    content = []
    slide = presentation.Slides(slide_number)
    for shape in slide.Shapes:
        if shape.HasTextFrame:
            if shape.TextFrame.HasText:
                content.append(shape.TextFrame.TextRange.Text)
    return "\n".join(content)


def get_presentation_content(presentation):
    """Return the content of all slides in the presentation."""
    content = []
    for i in range(1, presentation.Slides.Count + 1):
        slide_content = get_slide_content(presentation, i)
        content.append(f"Slide {i}:\n{slide_content}")
    return "\n\n".join(content)


def set_slide_size(presentation, width_points: float, height_points: float):
    """Set the slide size (Design -> Slide Size) to specific dimensions in points."""
    presentation.PageSetup.SlideWidth = width_points
    presentation.PageSetup.SlideHeight = height_points


def create_blank_slide(presentation, position: Optional[int] = None, layout: str = "blank"):
    layout_map = {"blank": 12, "title": 1}
    layout_id = layout_map.get(layout.lower().strip(), 12)
    if position is None:
        position = presentation.Slides.Count + 1
    slide = presentation.Slides.Add(position, layout_id)
    return slide


def add_text_box_at(slide, left: float, top: float, width: float, height: float, text: str):
    tb = slide.Shapes.AddTextbox(Orientation=1, Left=left, Top=top, Width=width, Height=height)
    rng = tb.TextFrame.TextRange
    rng.Text = text
    return tb


def apply_font_format(text_range, font: dict[str, Any]):
    """Apply font formatting with full Word parity."""
    f = text_range.Font
    if "name" in font and font["name"]:
        f.Name = font["name"]
    if "size" in font and font["size"] is not None:
        validate_font_size(float(font["size"]))
        f.Size = float(font["size"])
    if "bold" in font and font["bold"] is not None:
        f.Bold = -1 if font["bold"] else 0
    if "italic" in font and font["italic"] is not None:
        f.Italic = -1 if font["italic"] else 0
    if "underline" in font and font["underline"] is not None:
        f.Underline = -1 if font["underline"] else 0
    if "strikethrough" in font and font["strikethrough"] is not None:
        f.Strike = -1 if font["strikethrough"] else 0
    if "superscript" in font and font["superscript"] is not None:
        if font["superscript"]:
            f.Superscript = -1
            f.Subscript = 0
    if "subscript" in font and font["subscript"] is not None:
        if font["subscript"]:
            f.Subscript = -1
            f.Superscript = 0
    if "color" in font and font["color"]:
        color_val = validate_color(font["color"])  # May raise PPTValidationError
        f.Color.RGB = color_val


def apply_paragraph_format(text_range, paragraph: dict[str, Any]):
    """Apply paragraph formatting with enhanced options."""
    pf = text_range.ParagraphFormat
    if "alignment" in paragraph and paragraph["alignment"]:
        align_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
        pf.Alignment = align_map.get(paragraph["alignment"].lower(), 1)
    if "line_spacing" in paragraph and paragraph["line_spacing"]:
        try:
            ls = float(paragraph["line_spacing"])
            if ls >= 0.25 and ls <= 132:  # PowerPoint limits
                pf.SpaceWithin = ls
        except Exception:
            pass
    if "space_before" in paragraph and paragraph["space_before"] is not None:
        try:
            sb = float(paragraph["space_before"])
            if sb >= 0 and sb <= 1584:  # PowerPoint limits in points
                pf.SpaceBefore = sb
        except Exception:
            pass
    if "space_after" in paragraph and paragraph["space_after"] is not None:
        try:
            sa = float(paragraph["space_after"])
            if sa >= 0 and sa <= 1584:
                pf.SpaceAfter = sa
        except Exception:
            pass
    if "bullets" in paragraph and paragraph["bullets"] is not None:
        pf.Bullet.Visible = -1 if paragraph["bullets"] else 0
    if "left_indent" in paragraph and paragraph["left_indent"] is not None:
        try:
            li = float(paragraph["left_indent"])
            if li >= 0:
                pf.LeftIndent = li
        except Exception:
            pass
    if "first_line_indent" in paragraph and paragraph["first_line_indent"] is not None:
        try:
            fli = float(paragraph["first_line_indent"])
            pf.FirstLineIndent = fli
        except Exception:
            pass


def list_slide_shapes(slide):
    items = []
    for idx in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(idx)
        items.append({
            "id": sh.Id,
            "name": getattr(sh, "Name", ""),
            "type": getattr(sh, "Type", None),
            "left_pt": float(getattr(sh, "Left", 0.0)),
            "top_pt": float(getattr(sh, "Top", 0.0)),
            "width_pt": float(getattr(sh, "Width", 0.0)),
            "height_pt": float(getattr(sh, "Height", 0.0)),
            "z_order": getattr(sh, "ZOrderPosition", None),
        })
    return items


def set_shape_position(sh, left: float, top: float, width: Optional[float] = None, height: Optional[float] = None):
    sh.Left = left
    sh.Top = top
    if width is not None:
        sh.Width = width
    if height is not None:
        sh.Height = height


def add_image_at(slide, path: str, left: float, top: float, width: Optional[float], height: Optional[float], preserve_aspect: bool = True, dpi: float = 96.0):
    """Add image with enhanced validation and DPI support."""
    # Pre-validate bounds before adding
    if width is not None and height is not None:
        # Check final bounds
        slide_w = slide.Parent.PageSetup.SlideWidth
        slide_h = slide.Parent.PageSetup.SlideHeight
        validate_coordinates(left, top, width, height, slide_w, slide_h)
    
    pic = slide.Shapes.AddPicture(FileName=path, LinkToFile=False, SaveWithDocument=True, Left=left, Top=top, Width=-1, Height=-1)
    
    # Get natural dimensions for aspect calculations
    natural_width = pic.Width
    natural_height = pic.Height
    
    # Apply sizing with validation
    if width is not None and height is not None:
        pic.Width = width
        pic.Height = height
    elif width is not None:
        if preserve_aspect and natural_width > 0:
            ratio = natural_height / natural_width
            final_height = width * ratio
            # Validate the computed dimensions
            slide_w = slide.Parent.PageSetup.SlideWidth
            slide_h = slide.Parent.PageSetup.SlideHeight
            validate_coordinates(left, top, width, final_height, slide_w, slide_h)
            pic.Width = width
            pic.Height = final_height
        else:
            pic.Width = width
    elif height is not None:
        if preserve_aspect and natural_height > 0:
            ratio = natural_width / natural_height
            final_width = height * ratio
            # Validate the computed dimensions
            slide_w = slide.Parent.PageSetup.SlideWidth
            slide_h = slide.Parent.PageSetup.SlideHeight
            validate_coordinates(left, top, final_width, height, slide_w, slide_h)
            pic.Height = height
            pic.Width = final_width
        else:
            pic.Height = height
    
    return pic


## Deprecated legacy demo functions removed (add_text_to_slide, main) to keep module lean.
