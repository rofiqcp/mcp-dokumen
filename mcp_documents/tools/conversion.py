"""
Document conversion tools - PDF, image, Pandoc-based conversions
"""

from typing import Optional
from mcp_documents.converters.pdf import docx_to_pdf, pdf_to_docx, html_to_pdf
from mcp_documents.converters.image import convert_image, image_to_base64, base64_to_image, get_image_info
from mcp_documents.converters.pandoc import (
    convert_with_pandoc,
    markdown_to_docx,
    docx_to_markdown,
    markdown_to_pptx,
    markdown_to_html,
    html_to_docx,
    is_pandoc_available,
    get_supported_formats,
)
from mcp_documents.core import DocumentProcessor


def register_conversion_tools(mcp_server, doc_processor: DocumentProcessor):
    """Register document conversion tools with the MCP server."""
    
    # ==================== PDF Conversions ====================
    
    @mcp_server.tool()
    async def convert_docx_to_pdf(
        input_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        Convert a Word document to PDF.
        
        Args:
            input_path: Path to input DOCX file
            output_path: Optional output PDF path (auto-generated if not provided)
        
        Returns:
            Output path or error message
        """
        success, result = docx_to_pdf(input_path, output_path)
        if success:
            return f"PDF created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_pdf_to_docx(
        input_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        Convert a PDF to Word document.
        
        Args:
            input_path: Path to input PDF file
            output_path: Optional output DOCX path (auto-generated if not provided)
        
        Returns:
            Output path or error message
        """
        success, result = pdf_to_docx(input_path, output_path)
        if success:
            return f"DOCX created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_html_to_pdf(
        input_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        Convert HTML to PDF.
        
        Args:
            input_path: Path to input HTML file
            output_path: Optional output PDF path (auto-generated if not provided)
        
        Returns:
            Output path or error message
        """
        success, result = html_to_pdf(input_path, output_path)
        if success:
            return f"PDF created: {result}"
        return f"Error: {result}"
    
    # ==================== Image Conversions ====================
    
    @mcp_server.tool()
    async def convert_image_format(
        input_path: str,
        output_path: str,
        width: Optional[int] = None,
        height: Optional[int] = None,
        quality: int = 85
    ) -> str:
        """
        Convert an image to a different format with optional resizing.
        
        Args:
            input_path: Path to input image
            output_path: Path for output image (format determined by extension)
            width: Optional new width in pixels
            height: Optional new height in pixels
            quality: JPEG quality (1-100)
        
        Returns:
            Output path or error message
        """
        success, result = convert_image(input_path, output_path, width, height, quality)
        if success:
            return f"Image converted: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def get_image_as_base64(image_path: str) -> str:
        """
        Convert an image file to base64 string with data URI prefix.
        
        Args:
            image_path: Path to the image file
        
        Returns:
            Base64 data URI string or error message
        """
        success, result = image_to_base64(image_path)
        if success:
            return result
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def save_base64_as_image(
        base64_data: str,
        output_path: str
    ) -> str:
        """
        Save a base64 encoded image to a file.
        
        Args:
            base64_data: Base64 encoded image data (with or without data URI prefix)
            output_path: Path for output image file
        
        Returns:
            Output path or error message
        """
        success, result = base64_to_image(base64_data, output_path)
        if success:
            return f"Image saved: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def get_image_details(image_path: str) -> str:
        """
        Get information about an image file.
        
        Args:
            image_path: Path to the image file
        
        Returns:
            Image information (format, size, dimensions)
        """
        success, result = get_image_info(image_path)
        if success:
            lines = [f"{k}: {v}" for k, v in result.items()]
            return "\n".join(lines)
        return f"Error: {result.get('error', 'Unknown error')}"
    
    # ==================== Pandoc Conversions ====================
    
    @mcp_server.tool()
    async def convert_markdown_to_docx(
        input_path: str,
        output_path: Optional[str] = None,
        reference_doc: Optional[str] = None
    ) -> str:
        """
        Convert Markdown to Word document.
        
        Args:
            input_path: Path to input Markdown file
            output_path: Optional output DOCX path
            reference_doc: Optional reference DOCX for styling
        
        Returns:
            Output path or error message
        """
        success, result = markdown_to_docx(input_path, output_path, reference_doc)
        if success:
            return f"DOCX created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_docx_to_markdown(
        input_path: str,
        output_path: Optional[str] = None,
        extract_media: bool = False
    ) -> str:
        """
        Convert Word document to Markdown.
        
        Args:
            input_path: Path to input DOCX file
            output_path: Optional output Markdown path
            extract_media: Extract embedded images
        
        Returns:
            Output path or error message
        """
        success, result = docx_to_markdown(input_path, output_path, extract_media)
        if success:
            return f"Markdown created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_markdown_to_pptx(
        input_path: str,
        output_path: Optional[str] = None,
        reference_doc: Optional[str] = None
    ) -> str:
        """
        Convert Markdown to PowerPoint presentation.
        
        Args:
            input_path: Path to input Markdown file
            output_path: Optional output PPTX path
            reference_doc: Optional reference PPTX for styling
        
        Returns:
            Output path or error message
        """
        success, result = markdown_to_pptx(input_path, output_path, reference_doc)
        if success:
            return f"PPTX created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_markdown_to_html(
        input_path: str,
        output_path: Optional[str] = None,
        css: Optional[str] = None
    ) -> str:
        """
        Convert Markdown to HTML.
        
        Args:
            input_path: Path to input Markdown file
            output_path: Optional output HTML path
            css: Optional CSS file to include
        
        Returns:
            Output path or error message
        """
        success, result = markdown_to_html(input_path, output_path, css=css)
        if success:
            return f"HTML created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_html_to_docx(
        input_path: str,
        output_path: Optional[str] = None,
        reference_doc: Optional[str] = None
    ) -> str:
        """
        Convert HTML to Word document.
        
        Args:
            input_path: Path to input HTML file
            output_path: Optional output DOCX path
            reference_doc: Optional reference DOCX for styling
        
        Returns:
            Output path or error message
        """
        success, result = html_to_docx(input_path, output_path, reference_doc)
        if success:
            return f"DOCX created: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def convert_document(
        input_path: str,
        output_path: str,
        from_format: Optional[str] = None,
        to_format: Optional[str] = None
    ) -> str:
        """
        Convert between document formats using Pandoc.
        
        Args:
            input_path: Path to input file
            output_path: Path for output file
            from_format: Optional input format (auto-detected if not specified)
            to_format: Optional output format (auto-detected if not specified)
        
        Returns:
            Output path or error message
        
        Supported formats: markdown, docx, html, latex, epub, odt, rtf, and more.
        """
        success, result = convert_with_pandoc(input_path, output_path, from_format, to_format)
        if success:
            return f"Converted: {result}"
        return f"Error: {result}"
    
    @mcp_server.tool()
    async def check_pandoc() -> str:
        """
        Check if Pandoc is available and get supported formats.
        
        Returns:
            Pandoc status and supported formats
        """
        if not is_pandoc_available():
            return "Pandoc is not installed. Install from https://pandoc.org/"
        
        input_formats, output_formats = get_supported_formats()
        
        return (
            f"Pandoc is available.\n"
            f"Input formats ({len(input_formats)}): {', '.join(input_formats[:20])}...\n"
            f"Output formats ({len(output_formats)}): {', '.join(output_formats[:20])}..."
        )
