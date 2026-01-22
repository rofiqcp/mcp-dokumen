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
from mcp_documents.tools.header_footer_ops import register_header_footer_tools
from mcp_documents.tools.list_ops import register_list_tools
from mcp_documents.tools.comment_ops import register_comment_tools
from mcp_documents.tools.advanced_table_ops import register_advanced_table_tools
from mcp_documents.tools.footnote_ops import register_footnote_tools
from mcp_documents.tools.watermark_ops import register_watermark_tools
from mcp_documents.tools.hyperlink_ops import register_hyperlink_tools
from mcp_documents.tools.document_merge_ops import register_document_merge_tools
from mcp_documents.tools.track_changes_ops import register_track_changes_tools
from mcp_documents.tools.protection_ops import register_protection_tools
from mcp_documents.tools.endnote_ops import register_endnote_tools
from mcp_documents.tools.toc_ops import register_toc_tools
from mcp_documents.tools.mail_merge_ops import register_mail_merge_tools
from mcp_documents.tools.bookmark_ops import register_bookmark_tools
from mcp_documents.tools.table_border_ops import register_table_border_tools
from mcp_documents.tools.analysis_ops import register_analysis_tools
from mcp_documents.tools.batch_ops import register_batch_tools
from mcp_documents.tools.excel_ops import register_excel_tools
from mcp_documents.tools.section_ops import register_section_tools
from mcp_documents.tools.metadata_ops import register_metadata_tools
from mcp_documents.tools.security_ops import register_security_tools
from mcp_documents.tools.paragraph_ops import register_paragraph_tools
from mcp_documents.tools.structure_ops import register_structure_tools
from mcp_documents.tools.image_ops import register_image_tools
from mcp_documents.tools.extended_table_ops import register_extended_table_tools

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
    register_header_footer_tools(server, doc_processor)
    register_list_tools(server, doc_processor)
    register_comment_tools(server, doc_processor)
    register_advanced_table_tools(server, doc_processor)
    register_footnote_tools(server, doc_processor)
    register_watermark_tools(server, doc_processor)
    register_hyperlink_tools(server, doc_processor)
    register_document_merge_tools(server, doc_processor)
    register_track_changes_tools(server, doc_processor)
    register_protection_tools(server, doc_processor)
    register_endnote_tools(server, doc_processor)
    register_toc_tools(server, doc_processor)
    register_mail_merge_tools(server, doc_processor)
    register_bookmark_tools(server, doc_processor)
    register_table_border_tools(server, doc_processor)
    register_analysis_tools(server, doc_processor)
    register_batch_tools(server, doc_processor)
    register_excel_tools(server, doc_processor)
    register_section_tools(server, doc_processor)
    register_metadata_tools(server, doc_processor)
    register_security_tools(server, doc_processor)
    register_paragraph_tools(server, doc_processor)
    register_structure_tools(server, doc_processor)
    register_image_tools(server, doc_processor)
    register_extended_table_tools(server, doc_processor)
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
