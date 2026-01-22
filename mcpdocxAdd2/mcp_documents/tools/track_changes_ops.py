"""
Track changes operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_track_changes_tools(mcp_server, processor: DocumentProcessor):
    """Register track changes tools with the MCP server."""

    @mcp_server.tool()
    def enable_track_changes() -> str:
        """
        Enable track changes in the current document.
        All subsequent edits will be tracked as revisions.
        """
        return processor.enable_track_changes()

    @mcp_server.tool()
    def disable_track_changes() -> str:
        """
        Disable track changes in the current document.
        Future edits will not be tracked.
        """
        return processor.disable_track_changes()

    @mcp_server.tool()
    def accept_all_changes() -> str:
        """
        Accept all tracked changes in the document.
        Insertions are kept, deletions are removed.
        """
        return processor.accept_all_changes()

    @mcp_server.tool()
    def reject_all_changes() -> str:
        """
        Reject all tracked changes in the document.
        Insertions are removed, deletions are restored.
        """
        return processor.reject_all_changes()

    @mcp_server.tool()
    def add_revision_insert(paragraph_index: int, text: str, author: str = "User") -> str:
        """Add an insertion revision (w:ins) to a paragraph."""
        return processor.add_revision_insert(paragraph_index, text, author)

    @mcp_server.tool()
    def add_revision_delete(paragraph_index: int, start: int = 0, end: int = None, author: str = "User") -> str:
        """Add a deletion revision (w:del) wrapping a text span."""
        return processor.add_revision_delete(paragraph_index, start, end, author)
