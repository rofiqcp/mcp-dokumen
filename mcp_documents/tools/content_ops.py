"""
Content operation tools - paragraphs, headings, images, page breaks
"""

from typing import Optional, List
from mcp_documents.core import DocumentProcessor


def register_content_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register content operation tools with the MCP server."""
    
    @mcp_server.tool()
    async def add_paragraph(
        text: str,
        style: Optional[str] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        color: Optional[str] = None,
        alignment: Optional[str] = None
    ) -> str:
        """
        Add a paragraph to the current document.
        
        Args:
            text: Paragraph text content
            style: Optional Word style name (e.g., "Normal", "Quote")
            bold: Make text bold
            italic: Make text italic
            underline: Underline text
            font_size: Font size in points
            font_name: Font name (e.g., "Arial", "Times New Roman")
            color: Text color as hex (e.g., "#FF0000") or name (e.g., "red")
            alignment: Text alignment ("left", "center", "right", "justify")
        
        Returns:
            Success message or error
        """
        return doc_processor.add_paragraph(
            text=text,
            style=style,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_name=font_name,
            color=color,
            alignment=alignment
        )
    
    @mcp_server.tool()
    async def add_heading(text: str, level: int = 1) -> str:
        """
        Add a heading to the current document.
        
        Args:
            text: Heading text
            level: Heading level (0-9, where 0 is Title)
        
        Returns:
            Success message or error
        """
        return doc_processor.add_heading(text, level)
    
    @mcp_server.tool()
    async def add_image(
        image_path: str,
        width_inches: Optional[float] = None,
        height_inches: Optional[float] = None
    ) -> str:
        """
        Add an image to the current document.
        
        Args:
            image_path: Path to the image file
            width_inches: Optional width in inches
            height_inches: Optional height in inches
        
        Returns:
            Success message or error
        """
        return doc_processor.add_image(image_path, width_inches, height_inches)
    
    @mcp_server.tool()
    async def add_page_break() -> str:
        """
        Add a page break to the current document.
        
        Returns:
            Success message or error
        """
        return doc_processor.add_page_break()
    
    @mcp_server.tool()
    async def search_text(keyword: str) -> str:
        """
        Search for text in the current document.
        
        Args:
            keyword: Text to search for
        
        Returns:
            Search results with locations
        """
        return doc_processor.search_text(keyword)
    
    @mcp_server.tool()
    async def find_and_replace(find_text: str, replace_text: str) -> str:
        """
        Find and replace text in the current document.
        
        Args:
            find_text: Text to find
            replace_text: Text to replace with
        
        Returns:
            Number of replacements made
        """
        return doc_processor.find_and_replace(find_text, replace_text)
    
    @mcp_server.tool()
    async def delete_paragraph(index: int) -> str:
        """
        Delete a paragraph by index.
        
        Args:
            index: Paragraph index (0-based)
        
        Returns:
            Success message or error
        """
        return doc_processor.delete_paragraph(index)
    
    @mcp_server.tool()
    async def get_paragraph(index: int) -> str:
        """
        Get text of a specific paragraph.
        
        Args:
            index: Paragraph index (0-based)
        
        Returns:
            Paragraph text or error
        """
        paragraphs = doc_processor.get_paragraphs()
        if not paragraphs:
            return "No document open or empty document"
        
        if index < 0 or index >= len(paragraphs):
            return f"Error: Index out of range (0-{len(paragraphs)-1})"
        
        return paragraphs[index]
    
    @mcp_server.tool()
    async def count_paragraphs() -> str:
        """
        Get the number of paragraphs in the current document.
        
        Returns:
            Paragraph count or error
        """
        paragraphs = doc_processor.get_paragraphs()
        return f"Document has {len(paragraphs)} paragraphs"
