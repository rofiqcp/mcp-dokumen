"""
Pandoc-based document conversion utilities
"""

import os
import subprocess
import logging
import tempfile
from typing import Optional, Tuple, List, Dict

logger = logging.getLogger(__name__)

# Pandoc format mappings
FORMAT_MAP = {
    # Input formats
    '.md': 'markdown',
    '.markdown': 'markdown',
    '.txt': 'plain',
    '.html': 'html',
    '.htm': 'html',
    '.docx': 'docx',
    '.odt': 'odt',
    '.rtf': 'rtf',
    '.tex': 'latex',
    '.latex': 'latex',
    '.epub': 'epub',
    '.json': 'json',
    '.rst': 'rst',
    '.org': 'org',
    '.mediawiki': 'mediawiki',
    '.wiki': 'mediawiki',
    '.textile': 'textile',
    
    # Output formats
    '.pdf': 'pdf',
    '.pptx': 'pptx',
}


def is_pandoc_available() -> bool:
    """Check if Pandoc is installed and available."""
    try:
        result = subprocess.run(
            ['pandoc', '--version'],
            capture_output=True,
            text=True
        )
        return result.returncode == 0
    except FileNotFoundError:
        return False


def get_pandoc_version() -> Optional[str]:
    """Get Pandoc version string."""
    try:
        result = subprocess.run(
            ['pandoc', '--version'],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            first_line = result.stdout.split('\n')[0]
            return first_line.replace('pandoc ', '')
        return None
    except FileNotFoundError:
        return None


def convert_with_pandoc(
    input_path: str,
    output_path: str,
    from_format: Optional[str] = None,
    to_format: Optional[str] = None,
    extra_args: Optional[List[str]] = None,
    reference_doc: Optional[str] = None,
    template: Optional[str] = None,
    standalone: bool = True,
    metadata: Optional[Dict[str, str]] = None,
    variables: Optional[Dict[str, str]] = None,
    filters: Optional[List[str]] = None,
    defaults_file: Optional[str] = None,
    pdf_engine: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert a document using Pandoc.
    
    Args:
        input_path: Path to input file
        output_path: Path for output file
        from_format: Input format (auto-detected if None)
        to_format: Output format (auto-detected if None)
        extra_args: Additional Pandoc arguments
        reference_doc: Reference document for formatting (e.g., for DOCX/PPTX)
        template: Template file path
        standalone: Generate standalone document
        metadata: Key/value metadata to pass (--metadata key=value)
        variables: Pandoc template variables (--variable key=value)
        filters: List of Pandoc filters to apply (--filter path)
        defaults_file: Path to defaults YAML file (--defaults file)
        pdf_engine: PDF engine (e.g., xelatex, wkhtmltopdf)
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if not os.path.exists(input_path):
        return False, f"Input file not found: {input_path}"
    
    if not is_pandoc_available():
        return False, "Pandoc is not installed. Install from https://pandoc.org/"
    
    # Auto-detect formats from extensions
    if from_format is None:
        ext = os.path.splitext(input_path)[1].lower()
        from_format = FORMAT_MAP.get(ext)
        if from_format is None:
            return False, f"Unknown input format: {ext}"
    
    if to_format is None:
        ext = os.path.splitext(output_path)[1].lower()
        to_format = FORMAT_MAP.get(ext)
        if to_format is None:
            return False, f"Unknown output format: {ext}"
    
    # Build command
    cmd = ['pandoc', input_path, '-o', output_path]
    
    if from_format:
        cmd.extend(['-f', from_format])
    
    if to_format:
        cmd.extend(['-t', to_format])
    
    if standalone:
        cmd.append('-s')
    
    if reference_doc:
        if os.path.exists(reference_doc):
            cmd.extend(['--reference-doc', reference_doc])
    
    if template:
        if os.path.exists(template):
            cmd.extend(['--template', template])

    if defaults_file:
        if os.path.exists(defaults_file):
            cmd.extend(['--defaults', defaults_file])

    if filters:
        for flt in filters:
            cmd.extend(['--filter', flt])

    if metadata:
        for key, val in metadata.items():
            cmd.extend(['--metadata', f"{key}={val}"])

    if variables:
        for key, val in variables.items():
            cmd.extend(['--variable', f"{key}={val}"])

    if pdf_engine:
        cmd.extend(['--pdf-engine', pdf_engine])
    
    if extra_args:
        cmd.extend(extra_args)
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            if os.path.exists(output_path):
                return True, output_path
            return False, "Output file was not created"
        else:
            error_msg = result.stderr or result.stdout or "Unknown error"
            return False, f"Pandoc conversion failed: {error_msg}"
    
    except subprocess.TimeoutExpired:
        return False, "Pandoc conversion timed out"
    except Exception as e:
        return False, f"Pandoc conversion error: {e}"


def markdown_to_docx(
    input_path: str,
    output_path: Optional[str] = None,
    reference_doc: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert Markdown to DOCX.
    
    Args:
        input_path: Path to input Markdown file
        output_path: Path for output DOCX (auto-generated if None)
        reference_doc: Reference DOCX for styling
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    
    return convert_with_pandoc(
        input_path,
        output_path,
        from_format='markdown',
        to_format='docx',
        reference_doc=reference_doc
    )


def docx_to_markdown(
    input_path: str,
    output_path: Optional[str] = None,
    extract_media: bool = False
) -> Tuple[bool, str]:
    """
    Convert DOCX to Markdown.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path for output Markdown (auto-generated if None)
        extract_media: Extract embedded media files
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".md"
    
    extra_args = []
    if extract_media:
        media_dir = os.path.splitext(output_path)[0] + "_media"
        extra_args.extend(['--extract-media', media_dir])
    
    return convert_with_pandoc(
        input_path,
        output_path,
        from_format='docx',
        to_format='markdown',
        extra_args=extra_args if extra_args else None
    )


def markdown_to_pptx(
    input_path: str,
    output_path: Optional[str] = None,
    reference_doc: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert Markdown to PowerPoint (PPTX).
    
    Args:
        input_path: Path to input Markdown file
        output_path: Path for output PPTX (auto-generated if None)
        reference_doc: Reference PPTX for styling
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pptx"
    
    return convert_with_pandoc(
        input_path,
        output_path,
        from_format='markdown',
        to_format='pptx',
        reference_doc=reference_doc
    )


def markdown_to_html(
    input_path: str,
    output_path: Optional[str] = None,
    template: Optional[str] = None,
    css: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert Markdown to HTML.
    
    Args:
        input_path: Path to input Markdown file
        output_path: Path for output HTML (auto-generated if None)
        template: HTML template file
        css: CSS file to include
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".html"
    
    extra_args = []
    if css:
        extra_args.extend(['--css', css])
    
    return convert_with_pandoc(
        input_path,
        output_path,
        from_format='markdown',
        to_format='html',
        template=template,
        extra_args=extra_args if extra_args else None
    )


def html_to_docx(
    input_path: str,
    output_path: Optional[str] = None,
    reference_doc: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Convert HTML to DOCX.
    
    Args:
        input_path: Path to input HTML file
        output_path: Path for output DOCX (auto-generated if None)
        reference_doc: Reference DOCX for styling
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    
    return convert_with_pandoc(
        input_path,
        output_path,
        from_format='html',
        to_format='docx',
        reference_doc=reference_doc
    )


def convert_string(
    content: str,
    from_format: str,
    to_format: str
) -> Tuple[bool, str]:
    """
    Convert a string from one format to another using Pandoc.
    
    Args:
        content: Input content string
        from_format: Input format (e.g., 'markdown')
        to_format: Output format (e.g., 'html')
    
    Returns:
        Tuple of (success, converted_content or error_message)
    """
    if not is_pandoc_available():
        return False, "Pandoc is not installed"
    
    try:
        result = subprocess.run(
            ['pandoc', '-f', from_format, '-t', to_format],
            input=content,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode == 0:
            return True, result.stdout
        else:
            return False, result.stderr or "Conversion failed"
    
    except subprocess.TimeoutExpired:
        return False, "Conversion timed out"
    except Exception as e:
        return False, f"Conversion error: {e}"


def get_supported_formats() -> Tuple[List[str], List[str]]:
    """
    Get lists of supported input and output formats.
    
    Returns:
        Tuple of (input_formats, output_formats)
    """
    try:
        # Get input formats
        result = subprocess.run(
            ['pandoc', '--list-input-formats'],
            capture_output=True,
            text=True
        )
        input_formats = result.stdout.strip().split('\n') if result.returncode == 0 else []
        
        # Get output formats
        result = subprocess.run(
            ['pandoc', '--list-output-formats'],
            capture_output=True,
            text=True
        )
        output_formats = result.stdout.strip().split('\n') if result.returncode == 0 else []
        
        return input_formats, output_formats
    
    except FileNotFoundError:
        return [], []
