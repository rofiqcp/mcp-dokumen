"""
Image conversion utilities
"""

import os
import base64
import logging
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)

# Supported image formats
IMAGE_FORMATS = {
    "jpg": "JPEG",
    "jpeg": "JPEG",
    "png": "PNG",
    "gif": "GIF",
    "bmp": "BMP",
    "webp": "WebP",
    "tiff": "TIFF",
    "tif": "TIFF",
    "ico": "ICO",
}


def convert_image(
    input_path: str,
    output_path: str,
    width: Optional[int] = None,
    height: Optional[int] = None,
    quality: int = 85
) -> Tuple[bool, str]:
    """
    Convert an image to another format with optional resizing.
    
    Args:
        input_path: Path to input image
        output_path: Path for output image (format determined by extension)
        width: Optional new width (maintains aspect ratio if height not set)
        height: Optional new height (maintains aspect ratio if width not set)
        quality: JPEG quality (1-100)
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if not os.path.exists(input_path):
        return False, f"Input file not found: {input_path}"
    
    try:
        from PIL import Image
        
        # Open image
        img = Image.open(input_path)
        
        # Resize if needed
        if width or height:
            original_width, original_height = img.size
            
            if width and height:
                new_size = (width, height)
            elif width:
                ratio = width / original_width
                new_size = (width, int(original_height * ratio))
            else:
                ratio = height / original_height
                new_size = (int(original_width * ratio), height)
            
            img = img.resize(new_size, Image.Resampling.LANCZOS)
        
        # Determine output format
        output_ext = os.path.splitext(output_path)[1].lower().lstrip('.')
        output_format = IMAGE_FORMATS.get(output_ext)
        
        if not output_format:
            return False, f"Unsupported output format: {output_ext}"
        
        # Handle JPEG conversion (no alpha channel)
        if output_format == "JPEG" and img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = background
        
        # Save
        save_kwargs = {}
        if output_format == "JPEG":
            save_kwargs['quality'] = quality
            save_kwargs['optimize'] = True
        elif output_format == "PNG":
            save_kwargs['optimize'] = True
        
        img.save(output_path, format=output_format, **save_kwargs)
        
        if os.path.exists(output_path):
            return True, output_path
        
        return False, "Failed to save converted image"
    
    except ImportError:
        return False, "Pillow not installed. Run: pip install Pillow"
    except Exception as e:
        return False, f"Image conversion failed: {e}"


def image_to_base64(image_path: str) -> Tuple[bool, str]:
    """
    Convert an image file to base64 string.
    
    Args:
        image_path: Path to the image file
    
    Returns:
        Tuple of (success, base64_string or error_message)
    """
    if not os.path.exists(image_path):
        return False, f"Image not found: {image_path}"
    
    try:
        with open(image_path, 'rb') as f:
            data = base64.b64encode(f.read()).decode('utf-8')
        
        # Add data URI prefix
        ext = os.path.splitext(image_path)[1].lower().lstrip('.')
        mime_types = {
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'png': 'image/png',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'webp': 'image/webp',
            'svg': 'image/svg+xml',
        }
        mime = mime_types.get(ext, 'application/octet-stream')
        
        return True, f"data:{mime};base64,{data}"
    
    except Exception as e:
        return False, f"Failed to encode image: {e}"


def base64_to_image(base64_data: str, output_path: str) -> Tuple[bool, str]:
    """
    Convert a base64 string to an image file.
    
    Args:
        base64_data: Base64 encoded image data (with or without data URI prefix)
        output_path: Path for output image file
    
    Returns:
        Tuple of (success, output_path or error_message)
    """
    try:
        # Remove data URI prefix if present
        if ',' in base64_data:
            base64_data = base64_data.split(',', 1)[1]
        
        # Decode and save
        data = base64.b64decode(base64_data)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'wb') as f:
            f.write(data)
        
        if os.path.exists(output_path):
            return True, output_path
        
        return False, "Failed to save image"
    
    except Exception as e:
        return False, f"Failed to decode base64 image: {e}"


def get_image_info(image_path: str) -> Tuple[bool, dict]:
    """
    Get information about an image.
    
    Args:
        image_path: Path to the image file
    
    Returns:
        Tuple of (success, info_dict or error_message)
    """
    if not os.path.exists(image_path):
        return False, {"error": f"Image not found: {image_path}"}
    
    try:
        from PIL import Image
        
        img = Image.open(image_path)
        
        info = {
            "path": image_path,
            "format": img.format,
            "mode": img.mode,
            "width": img.size[0],
            "height": img.size[1],
            "size_bytes": os.path.getsize(image_path),
        }
        
        # Add EXIF data if available
        if hasattr(img, '_getexif') and img._getexif():
            info["has_exif"] = True
        else:
            info["has_exif"] = False
        
        return True, info
    
    except ImportError:
        return False, {"error": "Pillow not installed"}
    except Exception as e:
        return False, {"error": f"Failed to get image info: {e}"}


def get_supported_formats() -> List[str]:
    """Return list of supported image formats."""
    return list(set(IMAGE_FORMATS.values()))
