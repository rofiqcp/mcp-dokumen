"""
List operation tools - bulleted and numbered lists
"""

from typing import Optional, List
from mcp_documents.core import DocumentProcessor


def register_list_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register list operation tools with the MCP server."""
    
    @mcp_server.tool()
    async def add_bulleted_list(
        items: List[str],
        style: str = "List Bullet"
    ) -> str:
        """
        Add a bulleted list to the document.
        
        Args:
            items: List of item texts
            style: List style - "List Bullet", "List Bullet 2", "List Bullet 3"
        
        Returns:
            Success message or error
        """
        return doc_processor.add_bulleted_list(items, style)
    
    @mcp_server.tool()
    async def add_numbered_list(
        items: List[str],
        style: str = "List Number"
    ) -> str:
        """
        Add a numbered list to the document.
        
        Args:
            items: List of item texts
            style: List style - "List Number", "List Number 2", "List Number 3"
        
        Returns:
            Success message or error
        """
        return doc_processor.add_numbered_list(items, style)
    
    @mcp_server.tool()
    async def add_list_item(
        text: str,
        list_type: str = "bullet"
    ) -> str:
        """
        Add a single list item.
        
        Args:
            text: Item text
            list_type: "bullet" or "number"
        
        Returns:
            Success message or error
        """
        return doc_processor.add_list_item(text, list_type)
    
    @mcp_server.tool()
    async def add_multilevel_list(
        items: List[dict]
    ) -> str:
        """
        Add a multi-level list with nested items.
        
        Args:
            items: List of dicts with 'text' and 'level' keys
                   Example: [{"text": "Item 1", "level": 0}, {"text": "Sub item", "level": 1}]
        
        Returns:
            Success message or error
        """
        return doc_processor.add_multilevel_list(items)
