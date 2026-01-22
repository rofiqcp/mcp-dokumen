"""
Comment operation tools
"""

from typing import Optional, List
from mcp_documents.core import DocumentProcessor


def register_comment_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register comment operation tools with the MCP server."""
    
    @mcp_server.tool()
    async def add_comment(
        comment_text: str,
        paragraph_index: int,
        author: str = "MCP Documents",
        initials: str = "MCP"
    ) -> str:
        """
        Add a comment to a paragraph.
        
        Args:
            comment_text: The comment text
            paragraph_index: Index of paragraph to comment on (0-based)
            author: Comment author name
            initials: Author initials
        
        Returns:
            Success message or error
        """
        return doc_processor.add_comment(comment_text, paragraph_index, author, initials)
    
    @mcp_server.tool()
    async def get_all_comments() -> str:
        """
        Get all comments from the document.
        
        Returns:
            List of all comments with metadata
        """
        return doc_processor.get_all_comments()
    
    @mcp_server.tool()
    async def get_comments_by_author(author: str) -> str:
        """
        Get comments filtered by author name.
        
        Args:
            author: Author name to filter by
        
        Returns:
            List of comments by the specified author
        """
        return doc_processor.get_comments_by_author(author)
    
    @mcp_server.tool()
    async def delete_all_comments() -> str:
        """
        Delete all comments from the document.
        
        Returns:
            Success message or error
        """
        return doc_processor.delete_all_comments()
    
    @mcp_server.tool()
    async def get_comment_count() -> str:
        """
        Get the number of comments in the document.
        
        Returns:
            Comment count
        """
        return doc_processor.get_comment_count()
