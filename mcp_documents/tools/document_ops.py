"""
Document lifecycle tools - create, open, save, close, info
"""

from typing import Optional
from mcp_documents.core import DocumentProcessor


def register_document_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register document lifecycle tools with the MCP server."""
    
    @mcp_server.tool()
    async def create_document(
        file_path: str,
        title: Optional[str] = None
    ) -> str:
        """
        Create a new Word document.
        
        Args:
            file_path: Full path where to save the document
            title: Optional title to add as heading
        
        Returns:
            Success message or error
        """
        return doc_processor.create_document(file_path, title)
    
    @mcp_server.tool()
    async def open_document(file_path: str) -> str:
        """
        Open an existing Word document for editing.
        
        Args:
            file_path: Full path to the document
        
        Returns:
            Success message or error
        """
        return doc_processor.open_document(file_path)
    
    @mcp_server.tool()
    async def save_document(file_path: Optional[str] = None) -> str:
        """
        Save the current document.
        
        Args:
            file_path: Optional new path to save as (uses current path if not specified)
        
        Returns:
            Success message or error
        """
        return doc_processor.save_document(file_path)
    
    @mcp_server.tool()
    async def close_document() -> str:
        """
        Close the current document without saving.
        
        Returns:
            Success message
        """
        return doc_processor.close_document()
    
    @mcp_server.tool()
    async def get_document_info() -> str:
        """
        Get information about the current document.
        
        Returns:
            Document information including sections, paragraphs, tables count
        """
        return doc_processor.get_document_info()
    
    @mcp_server.tool()
    async def get_document_text() -> str:
        """
        Get all text content from the current document.
        
        Returns:
            All paragraph text from the document
        """
        paragraphs = doc_processor.get_paragraphs()
        if not paragraphs:
            return "No content or no document open"
        return "\n\n".join(paragraphs)
    
    @mcp_server.tool()
    async def list_open_documents() -> str:
        """
        List all currently open documents.
        
        Returns:
            List of open document paths
        """
        if not doc_processor.documents:
            return "No documents open"
        
        docs = list(doc_processor.documents.keys())
        current = doc_processor.current_file_path
        
        lines = []
        for doc in docs:
            marker = " (current)" if doc == current else ""
            lines.append(f"- {doc}{marker}")
        
        return "Open documents:\n" + "\n".join(lines)
    
    @mcp_server.tool()
    async def switch_document(file_path: str) -> str:
        """
        Switch to a different open document.
        
        Args:
            file_path: Path to the document to switch to
        
        Returns:
            Success message or error
        """
        if file_path not in doc_processor.documents:
            return f"Error: Document not open: {file_path}"
        
        doc_processor.current_document = doc_processor.documents[file_path]
        doc_processor.current_file_path = file_path
        
        return f"Switched to: {file_path}"
