"""TIER 3 (Advanced): 30 Enterprise Document Tools"""
import argparse, logging, sys
from pathlib import Path
from mcp.server.fastmcp import FastMCP

# Add parent directory to sys.path
parent_dir = str(Path(__file__).parent.parent.parent)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from mcp_documents.core import DocumentProcessor
from mcp_documents.tools.track_changes_ops import register_track_changes_tools
from mcp_documents.tools.protection_ops import register_protection_tools
from mcp_documents.tools.mail_merge_ops import register_mail_merge_tools
from mcp_documents.tools.watermark_ops import register_watermark_tools
from mcp_documents.tools.form_ops import register_form_tools
from mcp_documents.tools.security_ops import register_security_tools
from mcp_documents.tools.document_merge_ops import register_document_merge_tools
from mcp_documents.tools.batch_ops import register_batch_tools
from mcp_documents.tools.analysis_ops import register_analysis_tools
from mcp_documents.tools.collaboration_ops import register_collaboration_tools

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_server() -> tuple[FastMCP, DocumentProcessor]:
    """Create Tier 3 with 30 advanced tools."""
    server = FastMCP("mcp-documents-advanced")
    dp = DocumentProcessor()
    
    # Tier 3: 10 advanced modules + 2 inline = 30
    register_track_changes_tools(server, dp)      # 3
    register_protection_tools(server, dp)         # 3
    register_mail_merge_tools(server, dp)         # 3
    register_watermark_tools(server, dp)          # 2
    register_form_tools(server, dp)               # 3
    register_security_tools(server, dp)           # 3
    register_document_merge_tools(server, dp)     # 2
    register_batch_tools(server, dp)              # 3
    register_analysis_tools(server, dp)           # 3
    register_collaboration_tools(server, dp)      # 3
    
    # Inline: 2 essentials
    @server.tool()
    def list_revisions() -> str:
        """List all tracked changes."""
        return dp.list_revisions()
    
    @server.tool()
    def apply_document_security(password: str = None, read_only: bool = False) -> str:
        """Apply security settings to document."""
        if read_only: dp.set_readonly_mode(True)
        if password: return dp.set_document_password(password)
        return "Security settings applied"
    
    logger.info("TIER 3 (Advanced - 30 tools) initialized")
    return server, dp

def run_server():
    server, _ = create_server()
    logger.info("Starting TIER 3...")
    server.run()

def main():
    parser = argparse.ArgumentParser(description="TIER 3 - Advanced Operations")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    if args.debug: logging.getLogger().setLevel(logging.DEBUG)
    run_server()

if __name__ == "__main__": main()
