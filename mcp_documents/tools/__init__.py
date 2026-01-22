"""
Tools module for MCP Documents
"""

from mcp_documents.tools.document_ops import register_document_tools
from mcp_documents.tools.content_ops import register_content_tools
from mcp_documents.tools.table_ops import register_table_tools
from mcp_documents.tools.style_ops import register_style_tools
from mcp_documents.tools.conversion import register_conversion_tools

__all__ = [
    "register_document_tools",
    "register_content_tools",
    "register_table_tools",
    "register_style_tools",
    "register_conversion_tools",
]
