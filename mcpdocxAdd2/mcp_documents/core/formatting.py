"""
Text and paragraph formatting utilities
"""

from typing import Optional, Dict, Any
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from mcp_documents.core.utils import parse_color


def apply_paragraph_formatting(paragraph, formatting: Dict[str, Any]) -> None:
    """
    Apply formatting to a paragraph.
    
    Supported formatting options:
    - alignment: 'LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'
    - left_indent, right_indent, first_line_indent: Indentation in inches
    - space_before, space_after: Spacing in points
    - line_spacing: Line spacing as multiple or points
    - keep_together, keep_with_next, page_break_before, widow_control: Boolean pagination options
    """
    if not formatting:
        return
    
    para_format = paragraph.paragraph_format
    
    # Alignment
    alignment = formatting.get("alignment")
    if alignment:
        alignment_map = {
            "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
            "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
            "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
            "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if alignment.upper() in alignment_map:
            para_format.alignment = alignment_map[alignment.upper()]
    
    # Indentation
    if "left_indent" in formatting and formatting["left_indent"] is not None:
        para_format.left_indent = Inches(float(formatting["left_indent"]))
    
    if "right_indent" in formatting and formatting["right_indent"] is not None:
        para_format.right_indent = Inches(float(formatting["right_indent"]))
    
    if "first_line_indent" in formatting and formatting["first_line_indent"] is not None:
        para_format.first_line_indent = Inches(float(formatting["first_line_indent"]))
    
    # Spacing
    if "space_before" in formatting and formatting["space_before"] is not None:
        para_format.space_before = Pt(float(formatting["space_before"]))
    
    if "space_after" in formatting and formatting["space_after"] is not None:
        para_format.space_after = Pt(float(formatting["space_after"]))
    
    # Line spacing
    if "line_spacing" in formatting and formatting["line_spacing"] is not None:
        try:
            spacing_float = float(formatting["line_spacing"])
            para_format.line_spacing = spacing_float
        except ValueError:
            para_format.line_spacing = Pt(float(formatting["line_spacing"]))
    
    # Pagination options
    if "keep_together" in formatting:
        para_format.keep_together = bool(formatting["keep_together"])
    
    if "keep_with_next" in formatting:
        para_format.keep_with_next = bool(formatting["keep_with_next"])
    
    if "page_break_before" in formatting:
        para_format.page_break_before = bool(formatting["page_break_before"])
    
    if "widow_control" in formatting:
        para_format.widow_control = bool(formatting["widow_control"])


def apply_run_formatting(run, formatting: Dict[str, Any]) -> None:
    """
    Apply formatting to a run of text.
    
    Supported formatting options:
    - name: Font name
    - size: Font size in points
    - bold, italic, underline: Boolean style options
    - color: Color as hex (#RRGGBB) or rgb(r,g,b)
    """
    if not formatting:
        return
    
    font = run.font
    
    # Font name
    if "name" in formatting and formatting["name"]:
        font.name = formatting["name"]
        # Also set East Asian font for CJK support
        run._element.rPr.rFonts.set(qn('w:eastAsia'), formatting["name"])
    
    # Font size
    if "size" in formatting and formatting["size"] is not None:
        font.size = Pt(float(formatting["size"]))
    
    # Font styles
    if "bold" in formatting:
        font.bold = bool(formatting["bold"])
    
    if "italic" in formatting:
        font.italic = bool(formatting["italic"])
    
    if "underline" in formatting:
        font.underline = bool(formatting["underline"])
    
    # Font color
    if "color" in formatting and formatting["color"]:
        rgb = parse_color(formatting["color"])
        if rgb:
            font.color.rgb = RGBColor(*rgb)


def format_text_run(
    run,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    color: Optional[str] = None
) -> None:
    """
    Convenience function to format a text run.
    """
    formatting = {}
    
    if bold:
        formatting["bold"] = True
    if italic:
        formatting["italic"] = True
    if underline:
        formatting["underline"] = True
    if font_name:
        formatting["name"] = font_name
    if font_size:
        formatting["size"] = font_size
    if color:
        formatting["color"] = color
    
    apply_run_formatting(run, formatting)
