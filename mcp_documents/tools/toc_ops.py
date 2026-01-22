"""
Table of Contents operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_toc_tools(mcp_server, processor: DocumentProcessor):
    """Register Table of Contents tools with the MCP server."""

    @mcp_server.tool()
    def add_table_of_contents(title: str = "Table of Contents", 
                              heading_levels: int = 3) -> str:
        """
        Add a Table of Contents to the document.
        The TOC will include headings up to the specified level.
        
        Args:
            title: Title for the TOC section
            heading_levels: Number of heading levels to include (1-9)
        
        Returns:
            Success or error message
        
        Note: Open the document in Word and update the field to populate the TOC.
        """
        return processor.add_table_of_contents(title=title, 
                                               heading_levels=heading_levels)

    @mcp_server.tool()
    def update_toc() -> str:
        """
        Get instructions for updating the Table of Contents.
        TOC must be updated in Microsoft Word or compatible application.
        
        Returns:
            Instructions for updating TOC
        """
        return processor.update_toc()
