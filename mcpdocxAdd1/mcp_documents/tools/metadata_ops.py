"""
Metadata operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_metadata_tools(mcp_server, processor: DocumentProcessor):
    """Register metadata tools with the MCP server."""

    @mcp_server.tool()
    def get_metadata() -> str:
        """
        Get document metadata/properties.
        Includes title, author, subject, keywords, dates, etc.
        
        Returns:
            Document metadata or error message
        """
        return processor.get_metadata()

    @mcp_server.tool()
    def set_metadata(title: Optional[str] = None,
                      author: Optional[str] = None,
                      subject: Optional[str] = None,
                      keywords: Optional[str] = None,
                      category: Optional[str] = None,
                      comments: Optional[str] = None) -> str:
        """
        Set document metadata/properties.
        
        Args:
            title: Document title
            author: Document author
            subject: Document subject
            keywords: Keywords (comma-separated)
            category: Document category
            comments: Document comments
        
        Returns:
            Success or error message
        """
        return processor.set_metadata(
            title=title,
            author=author,
            subject=subject,
            keywords=keywords,
            category=category,
            comments=comments
        )

    @mcp_server.tool()
    def strip_personal_info() -> str:
        """
        Remove personal information from document metadata.
        Clears author, last modified by, and comments.
        
        Returns:
            Success or error message
        """
        return processor.strip_personal_info()
