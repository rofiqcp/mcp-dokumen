"""
Advanced table formatting tools.
"""

from mcp_documents.core import DocumentProcessor


def register_advanced_table_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register advanced table formatting tools."""
    
    @mcp_server.tool()
    async def set_table_cell_shading(
        table_index: int,
        row_index: int,
        col_index: int,
        fill_color: str = "FFFF00"
    ) -> str:
        """Apply background color to a cell."""
        return doc_processor.set_table_cell_shading(table_index, row_index, col_index, fill_color)
    
    @mcp_server.tool()
    async def apply_table_alternating_rows(
        table_index: int,
        color1: str = "FFFFFF",
        color2: str = "F2F2F2"
    ) -> str:
        """Apply alternating row colors to a table."""
        return doc_processor.apply_table_alternating_rows(table_index, color1, color2)
    
    @mcp_server.tool()
    async def highlight_table_header(
        table_index: int,
        header_color: str = "4472C4",
        text_color: str = "FFFFFF"
    ) -> str:
        """Highlight the first row as header."""
        return doc_processor.highlight_table_header(table_index, header_color, text_color)
    
    @mcp_server.tool()
    async def set_cell_alignment(
        table_index: int,
        row_index: int,
        col_index: int,
        horizontal: str = "center",
        vertical: str = "center"
    ) -> str:
        """Set horizontal and vertical alignment for a cell."""
        return doc_processor.set_cell_alignment(table_index, row_index, col_index, horizontal, vertical)
    
    @mcp_server.tool()
    async def set_cell_padding(
        table_index: int,
        row_index: int,
        col_index: int,
        padding_cm: float = 0.1
    ) -> str:
        """Set padding for a cell (all sides)."""
        return doc_processor.set_cell_padding(table_index, row_index, col_index, padding_cm)
    
    @mcp_server.tool()
    async def set_column_width(
        table_index: int,
        col_index: int,
        width_cm: float
    ) -> str:
        """Set width for a specific column across all rows."""
        return doc_processor.set_column_width(table_index, col_index, width_cm)
