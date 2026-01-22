"""
MCP Documents Server - Unified Document Processing MCP Server

A comprehensive Model Context Protocol server for document manipulation
including Word documents, PDFs, images, and format conversions.
"""

import argparse
import logging
import sys
from typing import Optional

from mcp.server.fastmcp import FastMCP

from mcp_documents.core import DocumentProcessor
from mcp_documents.tools.document_ops import register_document_tools
from mcp_documents.tools.content_ops import register_content_tools
from mcp_documents.tools.table_ops import register_table_tools
from mcp_documents.tools.style_ops import register_style_tools
from mcp_documents.tools.conversion import register_conversion_tools

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create and configure the MCP server with all tools."""
    
    # Create FastMCP server
    server = FastMCP("mcp-documents")
    
    # Create document processor (shared state)
    doc_processor = DocumentProcessor()
    
    # Register all tool groups
    register_document_tools(server, doc_processor)
    register_content_tools(server, doc_processor)
    register_table_tools(server, doc_processor)
    register_style_tools(server, doc_processor)
    register_conversion_tools(server, doc_processor)
    
    logger.info("MCP Documents server initialized with all tools")
    
    return server, doc_processor


def run_server():
    """Run the MCP server using stdio transport."""
    server, _ = create_server()
    
    logger.info("Starting MCP Documents server...")
    server.run()


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="MCP Documents - Unified Document Processing Server"
    )
    parser.add_argument(
        "--version",
        action="version",
        version="MCP Documents 1.0.0"
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging"
    )
    
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    run_server()


if __name__ == "__main__":
    main()
