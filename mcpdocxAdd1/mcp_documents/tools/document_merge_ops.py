"""
Document merge and insert tools.
"""

from typing import List
from mcp_documents.core import DocumentProcessor


def register_document_merge_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register document merge tools."""
    
    @mcp_server.tool()
    async def merge_documents(file_paths: List[str]) -> str:
        """Merge multiple documents into the current document."""
        return doc_processor.merge_documents(file_paths)
    
    @mcp_server.tool()
    async def insert_document(file_path: str) -> str:
        """Insert a single document at the end of current document."""
        return doc_processor.insert_document(file_path)
