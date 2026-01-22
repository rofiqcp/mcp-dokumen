"""
Core module for MCP Documents
"""

from mcp_documents.core.document import DocumentProcessor
from mcp_documents.core.utils import (
    get_temp_dir,
    encode_file_base64,
    decode_base64_to_file,
    validate_file_path,
    parse_color,
    logger,
)
from mcp_documents.core.formatting import (
    apply_paragraph_formatting,
    apply_run_formatting,
)

__all__ = [
    "DocumentProcessor",
    "get_temp_dir",
    "encode_file_base64",
    "decode_base64_to_file",
    "validate_file_path",
    "parse_color",
    "logger",
    "apply_paragraph_formatting",
    "apply_run_formatting",
]
