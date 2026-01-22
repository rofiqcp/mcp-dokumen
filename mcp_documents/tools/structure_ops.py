"""
Structure operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_structure_tools(mcp_server, processor: DocumentProcessor):
    """Register document structure tools with the MCP server."""

    @mcp_server.tool()
    def get_document_structure() -> str:
        """
        Get structural overview of the document.
        Includes sections, paragraphs, tables, and heading outline.
        
        Returns:
            Document structure or error message
        """
        return processor.get_document_structure()

    @mcp_server.tool()
    def get_raw_xml() -> str:
        """
        Get the raw XML structure of the document body.
        Useful for debugging and advanced manipulation.
        
        Returns:
            XML string (may be truncated) or error message
        """
        return processor.get_raw_xml()

    @mcp_server.tool()
    def list_docx_files(directory: str) -> str:
        """
        List all .docx files in a directory.
        
        Args:
            directory: Path to directory to scan
        
        Returns:
            List of files or error message
        """
        return processor.list_docx_files(directory)

    @mcp_server.tool()
    def copy_document(dest_path: str) -> str:
        """
        Create a copy of the current document.
        
        Args:
            dest_path: Destination file path
        
        Returns:
            Success or error message
        """
        return processor.copy_document(dest_path)

    @mcp_server.tool()
    def get_word_count() -> str:
        """
        Get detailed word count statistics.
        Includes words, characters, paragraphs, and lines.
        
        Returns:
            Word count statistics or error message
        """
        return processor.get_word_count()
