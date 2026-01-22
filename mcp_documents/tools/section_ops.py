"""
Section operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_section_tools(mcp_server, processor: DocumentProcessor):
    """Register section tools with the MCP server."""

    @mcp_server.tool()
    def add_section_break(section_type: str = "NEW_PAGE") -> str:
        """
        Add a new section break to the document.
        
        Args:
            section_type: Type of section break:
                - 'NEW_PAGE': Start on new page (default)
                - 'CONTINUOUS': Continue on same page
                - 'ODD_PAGE': Start on next odd page
                - 'EVEN_PAGE': Start on next even page
        
        Returns:
            Success or error message
        """
        return processor.add_section(section_type)

    @mcp_server.tool()
    def list_sections() -> str:
        """
        List all sections in the document with their properties.
        Shows orientation, page size, and margins for each section.
        
        Returns:
            Section listing or error message
        """
        return processor.list_sections()

    @mcp_server.tool()
    def set_section_properties(section_index: int = 0,
                                orientation: Optional[str] = None,
                                width_inches: Optional[float] = None,
                                height_inches: Optional[float] = None,
                                top_margin: Optional[float] = None,
                                bottom_margin: Optional[float] = None,
                                left_margin: Optional[float] = None,
                                right_margin: Optional[float] = None) -> str:
        """
        Set properties for a specific section.
        
        Args:
            section_index: Index of the section to modify
            orientation: 'PORTRAIT' or 'LANDSCAPE'
            width_inches: Page width in inches
            height_inches: Page height in inches
            top_margin: Top margin in inches
            bottom_margin: Bottom margin in inches
            left_margin: Left margin in inches
            right_margin: Right margin in inches
        
        Returns:
            Success or error message
        """
        return processor.set_section_properties(
            section_index=section_index,
            orientation=orientation,
            width_inches=width_inches,
            height_inches=height_inches,
            top_margin=top_margin,
            bottom_margin=bottom_margin,
            left_margin=left_margin,
            right_margin=right_margin
        )

    @mcp_server.tool()
    def add_zoned_header(section_index: int = 0,
                          left_text: str = "",
                          center_text: str = "",
                          right_text: str = "") -> str:
        """
        Add a header with left, center, and right zones.
        
        Args:
            section_index: Section to add header to
            left_text: Text for left zone
            center_text: Text for center zone
            right_text: Text for right zone
        
        Returns:
            Success or error message
        """
        return processor.add_zoned_header(
            section_index=section_index,
            left_text=left_text,
            center_text=center_text,
            right_text=right_text
        )

    @mcp_server.tool()
    def add_zoned_footer(section_index: int = 0,
                          left_text: str = "",
                          center_text: str = "",
                          right_text: str = "") -> str:
        """
        Add a footer with left, center, and right zones.
        
        Args:
            section_index: Section to add footer to
            left_text: Text for left zone
            center_text: Text for center zone
            right_text: Text for right zone
        
        Returns:
            Success or error message
        """
        return processor.add_zoned_footer(
            section_index=section_index,
            left_text=left_text,
            center_text=center_text,
            right_text=right_text
        )
