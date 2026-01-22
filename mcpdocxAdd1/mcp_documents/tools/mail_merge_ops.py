"""
Mail Merge operations for MCP Document Server
"""

from typing import List, Dict
from mcp_documents.core.document import DocumentProcessor


def register_mail_merge_tools(mcp_server, processor: DocumentProcessor):
    """Register mail merge tools with the MCP server."""

    @mcp_server.tool()
    def add_merge_field(field_name: str) -> str:
        """
        Add a mail merge field to the document.
        The field will be replaced with actual data during mail merge.
        
        Args:
            field_name: Name of the merge field (e.g., 'FirstName', 'Address')
        
        Returns:
            Success or error message
        """
        return processor.add_merge_field(field_name)

    @mcp_server.tool()
    def execute_mail_merge(data: List[Dict[str, str]], 
                           output_pattern: str) -> str:
        """
        Execute mail merge with provided data.
        Creates one document per record.
        
        Args:
            data: List of dictionaries with field values
                  Example: [{"FirstName": "John", "LastName": "Doe"}, ...]
            output_pattern: Output filename pattern with {index} placeholder
                           Example: "letter_{index}.docx"
        
        Returns:
            Success message with list of created files or error message
        """
        return processor.execute_mail_merge(data, output_pattern)
