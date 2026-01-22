"""TIER 1 (Basic): 30 Essential Document Tools"""
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

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create Tier 1 with 30 essential tools."""
    server = FastMCP("mcp-documents-basic")
    dp = DocumentProcessor()
    
    # Tier 1: 5 core modules (~26 tools) + 4 inline = 30 total
    register_document_tools(server, dp)     # 9 tools
    register_content_tools(server, dp)      # 10 tools
    register_table_tools(server, dp)        # 8 tools (consolidated from table_ops)
    register_paragraph_tools(server, dp)    # 6 tools (simplified)
    register_structure_tools(server, dp)    # 3 tools
    
    # Inline: 4 essentials
    @server.tool()
    def apply_style(para_idx: int, style_name: str) -> str:
        """Apply style to paragraph."""
        return dp.apply_style(para_idx, style_name)
    
    @server.tool()
    def list_styles() -> str:
        """List available styles."""
        return dp.list_styles()
    
    @server.tool()
    def convert_to_pdf(output_path: str = None) -> str:
        """Convert to PDF."""
        return dp.convert_to_pdf(output_path)
    
    @server.tool()
    def convert_to_markdown(output_path: str = None) -> str:
        """Convert to Markdown."""
        return dp.convert_to_markdown(output_path)
    
    logger.info("TIER 1 (Basic - 30 tools) initialized")
    return server, dp

def run_server():
    server, _ = create_server()
    logger.info("Starting TIER 1...")
    server.run()

def main():
    parser = argparse.ArgumentParser(description="TIER 1 - Essential Operations")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    if args.debug: logging.getLogger().setLevel(logging.DEBUG)
    run_server()

if __name__ == "__main__": main()
