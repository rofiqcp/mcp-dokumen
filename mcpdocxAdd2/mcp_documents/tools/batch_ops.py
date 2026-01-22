"""
Batch operations for MCP Document Server
"""

from typing import Dict, Optional
from mcp_documents.core.document import DocumentProcessor


def register_batch_tools(mcp_server, processor: DocumentProcessor):
    """Register batch operation tools with the MCP server."""

    @mcp_server.tool()
    def batch_replace(replacements: Dict[str, str]) -> str:
        """
        Replace multiple strings in the document at once.
        
        Args:
            replacements: Dictionary of old -> new text replacements
                         Example: {"old1": "new1", "old2": "new2"}
        
        Returns:
            Success message with replacement count or error message
        """
        return processor.batch_replace(replacements)

    @mcp_server.tool()
    def batch_format_paragraphs(style_name: str,
                                 start_index: int = 0,
                                 end_index: Optional[int] = None) -> str:
        """
        Apply a style to a range of paragraphs.
        
        Args:
            style_name: Name of the style to apply (e.g., 'Normal', 'Heading 1')
            start_index: Starting paragraph index (default 0)
            end_index: Ending paragraph index (exclusive, default all remaining)
        
        Returns:
            Success message with count or error message
        """
        return processor.batch_format_paragraphs(
            style_name=style_name,
            start_index=start_index,
            end_index=end_index
        )
