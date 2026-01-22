"""
Bookmark operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_bookmark_tools(mcp_server, processor: DocumentProcessor):
    """Register bookmark tools with the MCP server."""

    @mcp_server.tool()
    def add_bookmark(paragraph_index: int, bookmark_name: str) -> str:
        """
        Add a bookmark to a paragraph in the document.
        Bookmarks can be used for internal navigation.
        
        Args:
            paragraph_index: Index of the paragraph to bookmark
            bookmark_name: Name for the bookmark (no spaces)
        
        Returns:
            Success or error message
        """
        return processor.add_bookmark(paragraph_index, bookmark_name)

    @mcp_server.tool()
    def add_hyperlink_to_bookmark(text: str, bookmark_name: str) -> str:
        """
        Add a hyperlink that navigates to a bookmark in the document.
        
        Args:
            text: Display text for the hyperlink
            bookmark_name: Name of the target bookmark
        
        Returns:
            Success or error message
        """
        return processor.add_hyperlink_to_bookmark(text, bookmark_name)

    @mcp_server.tool()
    def list_bookmarks() -> str:
        """
        List all bookmarks in the current document.
        
        Returns:
            Comma-separated list of bookmark names or error message
        """
        return processor.list_bookmarks()
