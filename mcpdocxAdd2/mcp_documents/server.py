"""
MCP Documents Server - TIER 2: INTERMEDIATE
Enhanced productivity features for formatting, structure, and conversions
"""

import argparse
import logging
from mcp.server.fastmcp import FastMCP
from mcp_documents.core import DocumentProcessor

# TIER 2: Intermediate tools
from mcp_documents.tools.document_ops import register_document_tools
from mcp_documents.tools.content_ops import register_content_tools
from mcp_documents.tools.table_ops import register_table_tools
from mcp_documents.tools.paragraph_ops import register_paragraph_tools
from mcp_documents.tools.structure_ops import register_structure_tools
from mcp_documents.tools.header_footer_ops import register_header_footer_tools
from mcp_documents.tools.list_ops import register_list_tools
from mcp_documents.tools.image_ops import register_image_tools
from mcp_documents.tools.style_ops import register_style_tools
from mcp_documents.tools.comment_ops import register_comment_tools
from mcp_documents.tools.bookmark_ops import register_bookmark_tools
from mcp_documents.tools.hyperlink_ops import register_hyperlink_tools
from mcp_documents.tools.advanced_table_ops import register_advanced_table_tools
from mcp_documents.tools.table_border_ops import register_table_border_tools
from mcp_documents.tools.toc_ops import register_toc_tools
from mcp_documents.tools.metadata_ops import register_metadata_tools
from mcp_documents.tools.conversion import register_conversion_tools

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create Tier 2 server with intermediate tools."""
    server = FastMCP("mcp-documents-intermediate")
    doc_processor = DocumentProcessor()
    
    # Register Tier 2 tools
    register_document_tools(server, doc_processor)
    register_content_tools(server, doc_processor)
    register_table_tools(server, doc_processor)
    register_paragraph_tools(server, doc_processor)
    register_structure_tools(server, doc_processor)
    register_header_footer_tools(server, doc_processor)
    register_list_tools(server, doc_processor)
    register_image_tools(server, doc_processor)
    register_style_tools(server, doc_processor)
    register_comment_tools(server, doc_processor)
    register_bookmark_tools(server, doc_processor)
    register_hyperlink_tools(server, doc_processor)
    register_advanced_table_tools(server, doc_processor)
    register_table_border_tools(server, doc_processor)
    register_toc_tools(server, doc_processor)
    register_metadata_tools(server, doc_processor)
    register_conversion_tools(server, doc_processor)
    
    # Partial section tools (basic section operations only)
    @server.tool()
    def add_section_break(break_type: str = "nextPage") -> str:
        """Add a section break."""
        return doc_processor.add_section_break(break_type)
    
    @server.tool()
    def list_sections() -> str:
        """List all sections in the document."""
        return doc_processor.list_sections()
    
    @server.tool()
    def set_section_properties(section_index: int, orientation: str = None, 
                               width_inches: float = None, height_inches: float = None,
                               margin_top: float = None, margin_bottom: float = None,
                               margin_left: float = None, margin_right: float = None) -> str:
        """Set section properties (margins, orientation, size)."""
        return doc_processor.set_section_properties(
            section_index, orientation, width_inches, height_inches,
            margin_top, margin_bottom, margin_left, margin_right
        )
    
    logger.info("MCP Documents TIER 2 (Intermediate) initialized")
    return server, doc_processor


def run_server():
    """Run the Tier 2 MCP server."""
    server, _ = create_server()
    logger.info("Starting MCP Documents TIER 2 server...")
    server.run()


def main():
    parser = argparse.ArgumentParser(description="MCP Documents TIER 2 - Intermediate Operations")
    parser.add_argument("--version", action="version", version="MCP Documents TIER 2 v1.0.0")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    args = parser.parse_args()
    
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    run_server()


if __name__ == "__main__":
    main()
