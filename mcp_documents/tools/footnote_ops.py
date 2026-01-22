"""
Footnote tools.
"""

from mcp_documents.core import DocumentProcessor


def register_footnote_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register footnote tools."""
    
    @mcp_server.tool()
    async def add_footnote(text: str, paragraph_index: int) -> str:
        """Add a footnote to a paragraph."""
        return doc_processor.add_footnote(text, paragraph_index)
