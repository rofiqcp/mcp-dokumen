"""
Core utilities for MCP Documents
"""

import os
import base64
import tempfile
import shutil
import logging
from pathlib import Path
from typing import Optional, Union

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("MCPDocuments")


def get_temp_dir() -> str:
    """Get or create a temporary directory for processing."""
    temp_dir = os.path.join(tempfile.gettempdir(), "mcp_documents")
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir


def cleanup_temp_dir(path: str) -> None:
    """Clean up a temporary directory."""
    try:
        if os.path.exists(path) and path.startswith(tempfile.gettempdir()):
            shutil.rmtree(path)
    except Exception as e:
        logger.warning(f"Failed to clean up temp dir: {e}")


def encode_file_base64(file_path: str) -> str:
    """Read a file and return its base64 encoded content."""
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def decode_base64_to_file(content: str, output_path: str) -> str:
    """Decode base64 content and write to file."""
    data = base64.b64decode(content)
    with open(output_path, "wb") as f:
        f.write(data)
    return output_path


def ensure_directory(file_path: str) -> None:
    """Ensure the directory for a file path exists."""
    directory = os.path.dirname(file_path)
    if directory:
        os.makedirs(directory, exist_ok=True)


def validate_file_path(file_path: str, expected_extension: Optional[str] = None) -> str:
    """
    Validate that a file exists and optionally check its extension.
    Returns the absolute file path.
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    
    if expected_extension:
        ext = expected_extension.lower()
        if not ext.startswith('.'):
            ext = '.' + ext
        if not path.suffix.lower() == ext:
            raise ValueError(f"Expected {expected_extension} file, got: {path.suffix}")
    
    return str(path.absolute())


def find_file(filename: str, search_dirs: Optional[list] = None) -> Optional[str]:
    """
    Search for a file in multiple directories.
    Returns the first match or None.
    """
    if search_dirs is None:
        search_dirs = [
            os.getcwd(),
            tempfile.gettempdir(),
            os.path.expanduser("~"),
            os.path.expanduser("~/Documents"),
            os.path.expanduser("~/Downloads"),
        ]
    
    for directory in search_dirs:
        if not os.path.exists(directory):
            continue
        
        # Direct match
        full_path = os.path.join(directory, filename)
        if os.path.exists(full_path):
            return full_path
        
        # Case-insensitive search
        try:
            for f in os.listdir(directory):
                if f.lower() == filename.lower():
                    return os.path.join(directory, f)
        except PermissionError:
            continue
    
    return None


def parse_color(color: str) -> tuple:
    """
    Parse a color string and return (r, g, b) tuple.
    Supports: #RRGGBB, rgb(r,g,b), named colors
    """
    if not color:
        return None
    
    # Hex format
    if color.startswith('#'):
        color = color.lstrip('#')
        if len(color) == 6:
            r = int(color[0:2], 16)
            g = int(color[2:4], 16)
            b = int(color[4:6], 16)
            return (r, g, b)
    
    # RGB format
    if color.lower().startswith('rgb('):
        color = color[4:-1]  # Remove rgb( and )
        parts = color.split(',')
        if len(parts) == 3:
            r = int(parts[0].strip())
            g = int(parts[1].strip())
            b = int(parts[2].strip())
            return (r, g, b)
    
    # Named colors
    named_colors = {
        'red': (255, 0, 0),
        'green': (0, 128, 0),
        'blue': (0, 0, 255),
        'black': (0, 0, 0),
        'white': (255, 255, 255),
        'yellow': (255, 255, 0),
        'orange': (255, 165, 0),
        'purple': (128, 0, 128),
        'gray': (128, 128, 128),
        'grey': (128, 128, 128),
    }
    
    if color.lower() in named_colors:
        return named_colors[color.lower()]
    
    return None


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"
