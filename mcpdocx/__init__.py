"""
MCP DOCX Unified Server - 90 Tools for DOCX Operations
A comprehensive MCP server for reading, generating, and editing Word documents.
"""

__version__ = "1.0.0"
__author__ = "MCP DOCX Team"

from mcpdocx.processor import DocumentProcessor
from mcpdocx.server import create_server, run_server

__all__ = ["DocumentProcessor", "create_server", "run_server", "__version__"]
