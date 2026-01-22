"""
Document protection operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_protection_tools(mcp_server, processor: DocumentProcessor):
    """Register document protection tools with the MCP server."""

    @mcp_server.tool()
    def protect_document(protection_type: str = "readOnly", 
                         password: Optional[str] = None) -> str:
        """
        Protect the document with specified restrictions.
        
        Args:
            protection_type: Type of protection:
                - 'readOnly': Document is read-only
                - 'comments': Only comments allowed
                - 'trackedChanges': Only tracked changes allowed
                - 'forms': Only form fields can be edited
            password: Optional password to protect the document
        
        Returns:
            Success or error message
        """
        return processor.protect_document(password=password, 
                                          protection_type=protection_type)

    @mcp_server.tool()
    def unprotect_document() -> str:
        """
        Remove protection from the current document.
        
        Returns:
            Success or error message
        """
        return processor.unprotect_document()

    @mcp_server.tool()
    def add_editable_range(start_paragraph: int, end_paragraph: int, editor: str = "everyone") -> str:
        """Mark a range of paragraphs as editable."""
        return processor.add_editable_range(start_paragraph, end_paragraph, editor)
