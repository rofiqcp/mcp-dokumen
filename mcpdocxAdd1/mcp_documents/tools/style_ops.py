"""
Style and formatting tools - page layout, margins, styles
"""

from typing import Optional
from docx.shared import Cm, Inches
from mcp_documents.core import DocumentProcessor


def register_style_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register style and formatting tools with the MCP server."""
    
    @mcp_server.tool()
    async def set_page_margins(
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
        section_index: int = 0
    ) -> str:
        """
        Set page margins for the document.
        
        Args:
            top: Top margin in centimeters
            bottom: Bottom margin in centimeters
            left: Left margin in centimeters
            right: Right margin in centimeters
            section_index: Section to apply to (default: 0)
        
        Returns:
            Success message or error
        """
        return doc_processor.set_page_margins(top, bottom, left, right, section_index)
    
    @mcp_server.tool()
    async def set_page_size(
        width: float,
        height: float,
        section_index: int = 0
    ) -> str:
        """
        Set page size for the document.
        
        Args:
            width: Page width in centimeters
            height: Page height in centimeters
            section_index: Section to apply to (default: 0)
        
        Returns:
            Success message or error
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        try:
            sections = doc_processor.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            
            section = sections[section_index]
            section.page_width = Cm(width)
            section.page_height = Cm(height)
            
            return f"Page size set to {width}cm x {height}cm"
        except Exception as e:
            return f"Error: {e}"
    
    @mcp_server.tool()
    async def set_page_orientation(
        orientation: str,
        section_index: int = 0
    ) -> str:
        """
        Set page orientation.
        
        Args:
            orientation: "portrait" or "landscape"
            section_index: Section to apply to (default: 0)
        
        Returns:
            Success message or error
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        try:
            from docx.enum.section import WD_ORIENT
            
            sections = doc_processor.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            
            section = sections[section_index]
            
            if orientation.lower() == "landscape":
                section.orientation = WD_ORIENT.LANDSCAPE
                # Swap dimensions
                new_width = section.page_height
                new_height = section.page_width
                section.page_width = new_width
                section.page_height = new_height
            elif orientation.lower() == "portrait":
                section.orientation = WD_ORIENT.PORTRAIT
                # Swap dimensions if needed
                if section.page_width > section.page_height:
                    new_width = section.page_height
                    new_height = section.page_width
                    section.page_width = new_width
                    section.page_height = new_height
            else:
                return f"Error: Invalid orientation. Use 'portrait' or 'landscape'"
            
            return f"Page orientation set to {orientation}"
        except Exception as e:
            return f"Error: {e}"
    
    @mcp_server.tool()
    async def get_page_info(section_index: int = 0) -> str:
        """
        Get page layout information.
        
        Args:
            section_index: Section to query (default: 0)
        
        Returns:
            Page size, margins, and orientation
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        try:
            sections = doc_processor.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            
            section = sections[section_index]
            
            # Convert to cm
            def to_cm(emu):
                return round(emu / 360000, 2)
            
            info = []
            info.append(f"Section {section_index}:")
            info.append(f"  Page size: {to_cm(section.page_width)}cm x {to_cm(section.page_height)}cm")
            info.append(f"  Orientation: {section.orientation}")
            info.append(f"  Margins:")
            info.append(f"    Top: {to_cm(section.top_margin)}cm")
            info.append(f"    Bottom: {to_cm(section.bottom_margin)}cm")
            info.append(f"    Left: {to_cm(section.left_margin)}cm")
            info.append(f"    Right: {to_cm(section.right_margin)}cm")
            
            return "\n".join(info)
        except Exception as e:
            return f"Error: {e}"
    
    @mcp_server.tool()
    async def list_styles() -> str:
        """
        List available paragraph styles in the document.
        
        Returns:
            List of style names
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        try:
            from docx.enum.style import WD_STYLE_TYPE
            
            styles = doc_processor.current_document.styles
            paragraph_styles = []
            
            for style in styles:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    paragraph_styles.append(style.name)
            
            return "Paragraph styles:\n" + "\n".join(f"  - {s}" for s in sorted(paragraph_styles))
        except Exception as e:
            return f"Error: {e}"
    
    @mcp_server.tool()
    async def add_section() -> str:
        """
        Add a new section to the document.
        
        Returns:
            Success message with section count
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        try:
            from docx.enum.section import WD_SECTION
            
            doc_processor.current_document.add_section(WD_SECTION.NEW_PAGE)
            count = len(doc_processor.current_document.sections)
            
            return f"Section added. Document now has {count} sections."
        except Exception as e:
            return f"Error: {e}"
