"""
Form control operations for MCP Document Server
"""

from typing import List
from mcp_documents.core.document import DocumentProcessor


def register_form_tools(mcp_server, processor: DocumentProcessor):
    """Register form control tools."""

    @mcp_server.tool()
    def add_checkbox_form(paragraph_index: int, label: str = "Checkbox", checked: bool = False) -> str:
        """Insert a checkbox content control."""
        return processor.add_checkbox_form(paragraph_index, label, checked)

    @mcp_server.tool()
    def add_dropdown_form(paragraph_index: int, options: List[str], placeholder: str = "Select") -> str:
        """Insert a dropdown content control."""
        return processor.add_dropdown_form(paragraph_index, options, placeholder)

    @mcp_server.tool()
    def add_date_picker_form(paragraph_index: int, date_value: str = None, format_string: str = "yyyy-MM-dd") -> str:
        """Insert a date picker content control."""
        return processor.add_date_picker_form(paragraph_index, date_value, format_string)
