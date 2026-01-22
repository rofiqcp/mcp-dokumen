"""
Lightweight collaboration operations for MCP Document Server
"""

from mcp_documents.core.document import DocumentProcessor


def register_collaboration_tools(mcp_server, processor: DocumentProcessor):
    """Register collaboration tools."""

    @mcp_server.tool()
    def start_collaboration_session(name: str = None) -> str:
        """
        Start a new collaboration session for tracking changes.
        Returns a session ID that can be used to apply and log edits.
        
        Args:
            name: Optional name for the session
        
        Returns:
            Session ID for subsequent operations
        """
        return processor.start_collaboration_session(name)

    @mcp_server.tool()
    def apply_collaboration_change(session_id: str, paragraph_index: int, new_text: str) -> str:
        """
        Apply a text change to a paragraph within a collaboration session.
        Changes are logged with timestamp and version number.
        
        Args:
            session_id: Active collaboration session ID
            paragraph_index: Index of paragraph to modify
            new_text: New text content for the paragraph
        
        Returns:
            Success message with version number or error
        """
        return processor.apply_collaboration_change(session_id, paragraph_index, new_text)

    @mcp_server.tool()
    def get_collaboration_log(session_id: str) -> str:
        """
        Retrieve the change log for a collaboration session.
        Shows all edits with timestamps and version history.
        
        Args:
            session_id: Collaboration session ID
        
        Returns:
            Formatted log of all changes in the session
        """
        return processor.get_collaboration_log(session_id)
