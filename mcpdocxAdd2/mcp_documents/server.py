"""TIER 2 (Intermediate): 30 Productivity Document Tools"""
import argparse, logging, sys
from pathlib import Path
from mcp.server.fastmcp import FastMCP

# Add parent directory to sys.path
parent_dir = str(Path(__file__).parent.parent.parent)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from mcp_documents.core import DocumentProcessor
from mcp_documents.tools.document_ops import register_document_tools
from mcp_documents.tools.content_ops import register_content_tools
from mcp_documents.tools.table_ops import register_table_tools
from mcp_documents.tools.paragraph_ops import register_paragraph_tools
from mcp_documents.tools.structure_ops import register_structure_tools
from mcp_documents.tools.header_footer_ops import register_header_footer_tools
from mcp_documents.tools.list_ops import register_list_tools
from mcp_documents.tools.image_ops import register_image_tools

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create Tier 2 with 30 productivity tools."""
    server = FastMCP("mcp-documents-intermediate")
    dp = DocumentProcessor()
    
    # Tier 2: Tier 1 (5 modules) + 3 enhancement modules = 28 + 2 inline = 30
    register_document_tools(server, dp)      # 9
    register_content_tools(server, dp)       # 10
    register_table_tools(server, dp)         # 8
    register_paragraph_tools(server, dp)     # 6
    register_structure_tools(server, dp)     # 3
    register_header_footer_tools(server, dp) # 3
    register_list_tools(server, dp)          # 2
    register_image_tools(server, dp)         # 2
    
    # Inline: 2 essentials
    @server.tool()
    def add_section_break(break_type: str = "nextPage") -> str:
        """Add section break."""
        return dp.add_section_break(break_type)
    
    @server.tool()
    def add_comment(para_idx: int, comment_text: str, author: str = "Author") -> str:
        """Add comment to paragraph."""
        return dp.add_comment(para_idx, comment_text, author)
    
    logger.info("TIER 2 (Intermediate - 30 tools) initialized")
    return server, dp

def run_server():
    server, _ = create_server()
    logger.info("Starting TIER 2...")
    server.run()

def main():
    parser = argparse.ArgumentParser(description="TIER 2 - Intermediate Operations")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    if args.debug: logging.getLogger().setLevel(logging.DEBUG)
    run_server()

if __name__ == "__main__": main()
