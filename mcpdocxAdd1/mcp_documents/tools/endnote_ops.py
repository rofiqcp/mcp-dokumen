"""
Endnote operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_endnote_tools(mcp_server, processor: DocumentProcessor):
    """Register endnote tools with the MCP server."""

    @mcp_server.tool()
    def add_endnote(paragraph_index: int, endnote_text: str) -> str:
        """
        Add an endnote to a specific paragraph.
        Endnotes appear at the end of the document.
        
        Args:
            paragraph_index: Index of the paragraph to add endnote to
            endnote_text: Text content of the endnote
        
        Returns:
            Success or error message
        """
        return processor.add_endnote(paragraph_index, endnote_text)

    @mcp_server.tool()
    def render_endnotes() -> str:
        """
        Render all endnotes at the end of the document.
        Creates an 'Endnotes' section with numbered entries.
        
        Returns:
            Success or error message
        """
        return processor.render_endnotes()

    @mcp_server.tool()
    def add_native_footnote(paragraph_index: int, footnote_text: str) -> str:
        """
        Add a native Word footnote to a specific paragraph.
        
        Args:
            paragraph_index: Index of the paragraph to add footnote to
            footnote_text: Text content of the footnote
        
        Returns:
            Success or error message
        """
        return processor.add_native_footnote(paragraph_index, footnote_text)
