"""
Hyperlink tools.
"""

from mcp_documents.core import DocumentProcessor


def register_hyperlink_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register hyperlink tools."""
    
    @mcp_server.tool()
    async def add_hyperlink(
        paragraph_index: int,
        text: str,
        url: str,
        position: str = "end"
    ) -> str:
        """Add a hyperlink to a paragraph."""
        return doc_processor.add_hyperlink(paragraph_index, text, url, position)
