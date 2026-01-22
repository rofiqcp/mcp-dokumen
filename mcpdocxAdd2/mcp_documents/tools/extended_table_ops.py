"""
Extended table operations for MCP Document Server
"""

from typing import Optional, List
from mcp_documents.core.document import DocumentProcessor


def register_extended_table_tools(mcp_server, processor: DocumentProcessor):
    """Register extended table tools with the MCP server."""

    @mcp_server.tool()
    def merge_cells(table_index: int, 
                     start_row: int, start_col: int,
                     end_row: int, end_col: int) -> str:
        """
        Merge cells in a rectangular area.
        
        Args:
            table_index: Index of the table
            start_row: Starting row
            start_col: Starting column
            end_row: Ending row
            end_col: Ending column
        
        Returns:
            Success or error message
        """
        return processor.merge_cells(
            table_index=table_index,
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col
        )

    @mcp_server.tool()
    def get_table_data(table_index: int) -> str:
        """
        Get all data from a table as text.
        
        Args:
            table_index: Index of the table
        
        Returns:
            Table data or error message
        """
        return processor.get_table_data(table_index)

    @mcp_server.tool()
    def list_tables() -> str:
        """
        List all tables with their dimensions and styles.
        
        Returns:
            Table listing or error message
        """
        return processor.list_tables()

    @mcp_server.tool()
    def set_table_cell_text(table_index: int, row: int, col: int,
                             text: str) -> str:
        """
        Set text in a specific table cell.
        
        Args:
            table_index: Index of the table
            row: Row index
            col: Column index
            text: Text to set
        
        Returns:
            Success or error message
        """
        return processor.set_table_cell_text(
            table_index=table_index,
            row=row,
            col=col,
            text=text
        )

    @mcp_server.tool()
    def append_table_row(table_index: int, 
                         data: Optional[List[str]] = None) -> str:
        """
        Append a row to a table.
        
        Args:
            table_index: Index of the table
            data: Optional list of cell values for the new row
        
        Returns:
            Success or error message
        """
        return processor.add_table_row(table_index, data)

    @mcp_server.tool()
    def remove_table_row(table_index: int, row_index: int) -> str:
        """
        Remove a row from a table.
        
        Args:
            table_index: Index of the table
            row_index: Index of the row to remove
        
        Returns:
            Success or error message
        """
        return processor.delete_table_row(table_index, row_index)
