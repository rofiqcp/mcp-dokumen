"""
Table operation tools - create, edit, format tables
"""

from typing import Optional, List
from mcp_documents.core import DocumentProcessor


def register_table_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register table operation tools with the MCP server."""
    
    @mcp_server.tool()
    async def add_table(
        rows: int,
        cols: int,
        data: Optional[List[List[str]]] = None,
        style: str = "Table Grid"
    ) -> str:
        """
        Add a table to the current document.
        
        Args:
            rows: Number of rows
            cols: Number of columns
            data: Optional 2D array of cell data
            style: Table style name (default: "Table Grid")
        
        Returns:
            Success message or error
        """
        return doc_processor.add_table(rows, cols, data, style)
    
    @mcp_server.tool()
    async def edit_table_cell(
        table_index: int,
        row_index: int,
        col_index: int,
        text: str
    ) -> str:
        """
        Edit the content of a table cell.
        
        Args:
            table_index: Table index (0-based)
            row_index: Row index (0-based)
            col_index: Column index (0-based)
            text: New cell text
        
        Returns:
            Success message or error
        """
        return doc_processor.edit_table_cell(table_index, row_index, col_index, text)
    
    @mcp_server.tool()
    async def add_table_row(
        table_index: int,
        data: Optional[List[str]] = None
    ) -> str:
        """
        Add a row to an existing table.
        
        Args:
            table_index: Table index (0-based)
            data: Optional list of cell values
        
        Returns:
            Success message or error
        """
        return doc_processor.add_table_row(table_index, data)
    
    @mcp_server.tool()
    async def delete_table_row(table_index: int, row_index: int) -> str:
        """
        Delete a row from a table.
        
        Args:
            table_index: Table index (0-based)
            row_index: Row index to delete (0-based)
        
        Returns:
            Success message or error
        """
        return doc_processor.delete_table_row(table_index, row_index)
    
    @mcp_server.tool()
    async def merge_table_cells(
        table_index: int,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int
    ) -> str:
        """
        Merge multiple table cells.
        
        Args:
            table_index: Table index (0-based)
            start_row: Starting row index
            start_col: Starting column index
            end_row: Ending row index
            end_col: Ending column index
        
        Returns:
            Success message or error
        """
        return doc_processor.merge_table_cells(
            table_index, start_row, start_col, end_row, end_col
        )
    
    @mcp_server.tool()
    async def get_tables_info() -> str:
        """
        Get information about all tables in the document.
        
        Returns:
            Table information including dimensions
        """
        return doc_processor.get_tables_info()
    
    @mcp_server.tool()
    async def get_table_content(table_index: int) -> str:
        """
        Get all content from a specific table.
        
        Args:
            table_index: Table index (0-based)
        
        Returns:
            Table content as formatted text or error
        """
        if not doc_processor.current_document:
            return "Error: No document is open"
        
        tables = doc_processor.current_document.tables
        if table_index >= len(tables):
            return f"Error: Table index out of range (have {len(tables)} tables)"
        
        table = tables[table_index]
        rows_data = []
        
        for row in table.rows:
            cells = [cell.text for cell in row.cells]
            rows_data.append(" | ".join(cells))
        
        return "\n".join(rows_data)
