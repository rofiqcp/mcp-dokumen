"""
Advanced table border and sizing operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_table_border_tools(mcp_server, processor: DocumentProcessor):
    """Register table border and sizing tools with the MCP server."""

    @mcp_server.tool()
    def set_table_borders(table_index: int,
                          top: Optional[str] = None,
                          bottom: Optional[str] = None,
                          left: Optional[str] = None,
                          right: Optional[str] = None,
                          inside_h: Optional[str] = None,
                          inside_v: Optional[str] = None,
                          size: int = 4,
                          color: str = "000000") -> str:
        """
        Set borders for a table with per-side control.
        
        Args:
            table_index: Index of the table to modify
            top: Top border style
            bottom: Bottom border style
            left: Left border style
            right: Right border style
            inside_h: Inside horizontal border style
            inside_v: Inside vertical border style
            size: Border thickness (default 4)
            color: Border color as hex without # (default "000000")
        
        Border styles: 'single', 'double', 'dotted', 'dashed', 'none'
        
        Returns:
            Success or error message
        """
        return processor.set_table_borders(
            table_index=table_index,
            top=top,
            bottom=bottom,
            left=left,
            right=right,
            inside_h=inside_h,
            inside_v=inside_v,
            size=size,
            color=color
        )

    @mcp_server.tool()
    def set_row_height(table_index: int, 
                       row_index: int, 
                       height_cm: float,
                       rule: str = "exact") -> str:
        """
        Set the height of a specific row in a table.
        
        Args:
            table_index: Index of the table
            row_index: Index of the row to modify
            height_cm: Height in centimeters
            rule: Height rule:
                - 'exact': Exact height
                - 'atLeast': Minimum height
                - 'auto': Automatic height
        
        Returns:
            Success or error message
        """
        return processor.set_row_height(
            table_index=table_index,
            row_index=row_index,
            height_cm=height_cm,
            rule=rule
        )
