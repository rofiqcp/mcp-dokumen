"""
MCP Documents - Unified Document Processing Server

A comprehensive Model Context Protocol server for document manipulation.
"""

__version__ = "1.0.0"

from mcp_documents.core.document import DocumentProcessor
from mcp_documents.core.formatting import apply_paragraph_formatting, apply_run_formatting

__all__ = [
    "__version__",
    "DocumentProcessor",
    "apply_paragraph_formatting",
    "apply_run_formatting",
]
