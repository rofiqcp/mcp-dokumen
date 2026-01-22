"""
Security and redaction operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_security_tools(mcp_server, processor: DocumentProcessor):
    """Register security tools with the MCP server."""

    @mcp_server.tool()
    def redact_text(pattern: str, replacement: str = "[REDACTED]",
                     use_regex: bool = False) -> str:
        """
        Redact text matching pattern throughout the document.
        
        Args:
            pattern: Text or regex pattern to find
            replacement: Replacement text (default: [REDACTED])
            use_regex: Whether to use regex matching
        
        Returns:
            Count of redactions or error message
        """
        return processor.redact_text(
            pattern=pattern,
            replacement=replacement,
            use_regex=use_regex
        )

    @mcp_server.tool()
    def sanitize_external_links() -> str:
        """
        Remove all external hyperlinks from the document.
        Link text is preserved but the hyperlink is removed.
        
        Returns:
            Count of removed links or error message
        """
        return processor.sanitize_external_links()
