"""
Header and footer operation tools
"""

from typing import Optional
from mcp_documents.core import DocumentProcessor


def register_header_footer_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register header and footer tools with the MCP server."""
    
    @mcp_server.tool()
    async def add_header(
        text: str,
        section_index: int = 0,
        alignment: str = "center"
    ) -> str:
        """
        Add a header to the document section.
        
        Args:
            text: Header text content
            section_index: Section to add header to (default: 0)
            alignment: Text alignment - "left", "center", "right"
        
        Returns:
            Success message or error
        """
        return doc_processor.add_header(text, section_index, alignment)
    
    @mcp_server.tool()
    async def add_footer(
        text: str,
        section_index: int = 0,
        alignment: str = "center"
    ) -> str:
        """
        Add a footer to the document section.
        
        Args:
            text: Footer text content
            section_index: Section to add footer to (default: 0)
            alignment: Text alignment - "left", "center", "right"
        
        Returns:
            Success message or error
        """
        return doc_processor.add_footer(text, section_index, alignment)
    
    @mcp_server.tool()
    async def add_page_numbers(
        position: str = "footer",
        alignment: str = "center",
        format_string: str = "Page {page}",
        section_index: int = 0
    ) -> str:
        """
        Add page numbers to header or footer.
        
        Args:
            position: "header" or "footer"
            alignment: "left", "center", or "right"
            format_string: Format - use {page} for page number
            section_index: Section to add to
        
        Returns:
            Success message or error
        """
        return doc_processor.add_page_numbers(position, alignment, format_string, section_index)
    
    @mcp_server.tool()
    async def remove_header(section_index: int = 0) -> str:
        """
        Remove header from a section.
        
        Args:
            section_index: Section to remove header from
        
        Returns:
            Success message or error
        """
        return doc_processor.remove_header(section_index)
    
    @mcp_server.tool()
    async def remove_footer(section_index: int = 0) -> str:
        """
        Remove footer from a section.
        
        Args:
            section_index: Section to remove footer from
        
        Returns:
            Success message or error
        """
        return doc_processor.remove_footer(section_index)
    
    @mcp_server.tool()
    async def get_header_text(section_index: int = 0) -> str:
        """
        Get the header text from a section.
        
        Args:
            section_index: Section to get header from
        
        Returns:
            Header text or error
        """
        return doc_processor.get_header_text(section_index)
    
    @mcp_server.tool()
    async def get_footer_text(section_index: int = 0) -> str:
        """
        Get the footer text from a section.
        
        Args:
            section_index: Section to get footer from
        
        Returns:
            Footer text or error
        """
        return doc_processor.get_footer_text(section_index)
    
    @mcp_server.tool()
    async def set_different_first_page_header(
        first_page_text: str,
        other_pages_text: str,
        section_index: int = 0
    ) -> str:
        """
        Set different header for first page vs other pages.
        
        Args:
            first_page_text: Header text for first page
            other_pages_text: Header text for other pages
            section_index: Section to apply to
        
        Returns:
            Success message or error
        """
        return doc_processor.set_different_first_page_header(
            first_page_text, other_pages_text, section_index
        )
