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
