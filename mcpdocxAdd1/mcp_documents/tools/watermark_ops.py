"""
Watermark tools.
"""

from mcp_documents.core import DocumentProcessor


def register_watermark_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register watermark tools."""
    
    @mcp_server.tool()
    async def add_text_watermark(text: str = "CONFIDENTIAL", section_index: int = 0) -> str:
        """Add a simple text watermark via header."""
        return doc_processor.add_text_watermark(text, section_index)
