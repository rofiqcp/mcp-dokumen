"""
MCP Documents Server - TIER 1: ESSENTIAL/BASIC
Essential document operations for basic CRUD and simple formatting
"""

import argparse
import logging
from mcp.server.fastmcp import FastMCP
from mcp_documents.core import DocumentProcessor

# TIER 1: Essential tools only
from mcp_documents.tools.document_ops import register_document_tools
from mcp_documents.tools.content_ops import register_content_tools
from mcp_documents.tools.table_ops import register_table_tools
from mcp_documents.tools.paragraph_ops import register_paragraph_tools
from mcp_documents.tools.structure_ops import register_structure_tools

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create Tier 1 server with essential tools only."""
    server = FastMCP("mcp-documents-basic")
    doc_processor = DocumentProcessor()
    
    # Register Tier 1 tools
    register_document_tools(server, doc_processor)
    register_content_tools(server, doc_processor)
    register_table_tools(server, doc_processor)
    register_paragraph_tools(server, doc_processor)
    register_structure_tools(server, doc_processor)
    
    # Partial style tools (basic formatting only)
    @server.tool()
    def apply_style(paragraph_index: int, style_name: str) -> str:
        """Apply a style to a paragraph."""
        return doc_processor.apply_style(paragraph_index, style_name)
    
    @server.tool()
    def list_styles() -> str:
        """List all available styles in the document."""
        return doc_processor.list_styles()
    
    # Partial conversion tools (basic formats only)
    @server.tool()
    def convert_to_pdf(output_path: str = None) -> str:
        """Convert current document to PDF."""
        return doc_processor.convert_to_pdf(output_path)
    
    @server.tool()
    def convert_to_markdown(output_path: str = None) -> str:
        """Convert current document to Markdown."""
        return doc_processor.convert_to_markdown(output_path)
    
    @server.tool()
    def convert_to_html(output_path: str = None) -> str:
        """Convert current document to HTML."""
        return doc_processor.convert_to_html(output_path)
    
    logger.info("MCP Documents TIER 1 (Basic) initialized")
    return server, doc_processor


def run_server():
    """Run the Tier 1 MCP server."""
    server, _ = create_server()
    logger.info("Starting MCP Documents TIER 1 server...")
    server.run()


def main():
    parser = argparse.ArgumentParser(description="MCP Documents TIER 1 - Essential Operations")
    parser.add_argument("--version", action="version", version="MCP Documents TIER 1 v1.0.0")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    run_server()


if __name__ == "__main__":
    main()
