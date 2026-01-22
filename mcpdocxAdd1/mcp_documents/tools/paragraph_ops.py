"""
Paragraph text operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_paragraph_tools(mcp_server, processor: DocumentProcessor):
    """Register paragraph text tools with the MCP server."""

    @mcp_server.tool()
    def get_paragraph_text(index: int) -> str:
        """
        Get text from a specific paragraph.
        
        Args:
            index: Paragraph index (0-based)
        
        Returns:
            Paragraph text or error message
        """
        return processor.get_paragraph_text(index)

    @mcp_server.tool()
    def set_paragraph_text(index: int, text: str) -> str:
        """
        Replace text of a specific paragraph.
        
        Args:
            index: Paragraph index (0-based)
            text: New text content
        
        Returns:
            Success or error message
        """
        return processor.set_paragraph_text(index, text)

    @mcp_server.tool()
    def remove_paragraph(index: int) -> str:
        """
        Remove a paragraph by index.
        
        Args:
            index: Paragraph index (0-based)
        
        Returns:
            Success or error message
        """
        return processor.delete_paragraph(index)

    @mcp_server.tool()
    def insert_paragraph_after(index: int, text: str,
                                style: Optional[str] = None) -> str:
        """
        Insert a paragraph after the specified index.
        
        Args:
            index: Paragraph index to insert after
            text: Text content for new paragraph
            style: Optional style name
        
        Returns:
            Success or error message
        """
        return processor.insert_paragraph_after(index, text, style)

    @mcp_server.tool()
    def find_text(search_text: str, case_sensitive: bool = False) -> str:
        """
        Find all occurrences of text in the document.
        
        Args:
            search_text: Text to search for
            case_sensitive: Whether to match case
        
        Returns:
            List of occurrences or error message
        """
        return processor.find_text(search_text, case_sensitive)
