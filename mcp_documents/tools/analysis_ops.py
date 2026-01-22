"""
Document analysis operations for MCP Document Server
"""

from typing import Dict
from mcp_documents.core.document import DocumentProcessor


def register_analysis_tools(mcp_server, processor: DocumentProcessor):
    """Register document analysis tools with the MCP server."""

    @mcp_server.tool()
    def get_document_outline() -> str:
        """
        Get the document outline based on headings.
        Returns a hierarchical view of all headings.
        
        Returns:
            Document outline or error message
        """
        return processor.get_document_outline()

    @mcp_server.tool()
    def get_document_statistics() -> str:
        """
        Get comprehensive statistics about the document.
        Includes paragraph count, word count, table count, etc.
        
        Returns:
            Document statistics or error message
        """
        return processor.get_document_statistics()

    @mcp_server.tool()
    def get_formatting_report() -> str:
        """
        Get a report of fonts and styles used in the document.
        Useful for checking formatting consistency.
        
        Returns:
            Formatting report or error message
        """
        return processor.get_formatting_report()
