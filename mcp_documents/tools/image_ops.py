"""
Image operations for MCP Document Server
"""

from typing import Optional
from mcp_documents.core.document import DocumentProcessor


def register_image_tools(mcp_server, processor: DocumentProcessor):
    """Register image tools with the MCP server."""

    @mcp_server.tool()
    def insert_image(image_path: str, 
                     width_inches: Optional[float] = None,
                     height_inches: Optional[float] = None) -> str:
        """
        Insert an image into the document.
        
        Args:
            image_path: Path to the image file
            width_inches: Optional width in inches
            height_inches: Optional height in inches
        
        Returns:
            Success or error message
        """
        return processor.add_image(
            image_path=image_path,
            width_inches=width_inches,
            height_inches=height_inches
        )

    @mcp_server.tool()
    def add_image_base64(base64_data: str, 
                          width_inches: Optional[float] = None) -> str:
        """
        Add an image from base64 encoded data.
        
        Args:
            base64_data: Base64 encoded image (with or without data URI prefix)
            width_inches: Optional width in inches
        
        Returns:
            Success or error message
        """
        return processor.add_image_base64(
            base64_data=base64_data,
            width_inches=width_inches
        )

    @mcp_server.tool()
    def list_images() -> str:
        """
        List all images in the document with dimensions.
        
        Returns:
            List of images or error message
        """
        return processor.list_images()
