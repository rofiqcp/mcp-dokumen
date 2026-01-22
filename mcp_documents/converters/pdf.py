"""
PDF conversion utilities
"""

import os
import subprocess
import logging
from typing import Optional, Tuple

logger = logging.getLogger(__name__)


def docx_to_pdf(input_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
    """
    Convert DOCX to PDF.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path for output PDF (auto-generated if None)
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if not os.path.exists(input_path):
        return False, f"Input file not found: {input_path}"
    
    # Generate output path if not provided
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    
    # Try docx2pdf first (requires MS Word or LibreOffice)
    try:
        from docx2pdf import convert
        convert(input_path, output_path)
        if os.path.exists(output_path):
            return True, output_path
    except ImportError:
        logger.debug("docx2pdf not available")
    except Exception as e:
        logger.debug(f"docx2pdf failed: {e}")
    
    # Fallback: Try LibreOffice
    try:
        output_dir = os.path.dirname(output_path) or "."
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_path
            ],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            # LibreOffice outputs to same basename
            expected_output = os.path.join(
                output_dir,
                os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
            )
            if os.path.exists(expected_output):
                if expected_output != output_path:
                    os.rename(expected_output, output_path)
                return True, output_path
    except FileNotFoundError:
        logger.debug("LibreOffice not found")
    except Exception as e:
        logger.debug(f"LibreOffice conversion failed: {e}")
    
    # Fallback: Try unoconv
    try:
        result = subprocess.run(
            ["unoconv", "-f", "pdf", "-o", output_path, input_path],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0 and os.path.exists(output_path):
            return True, output_path
    except FileNotFoundError:
        logger.debug("unoconv not found")
    except Exception as e:
        logger.debug(f"unoconv conversion failed: {e}")
    
    return False, "No PDF converter available. Install docx2pdf, LibreOffice, or unoconv."


def pdf_to_docx(input_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
    """
    Convert PDF to DOCX.
    
    Args:
        input_path: Path to input PDF file
        output_path: Path for output DOCX (auto-generated if None)
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if not os.path.exists(input_path):
        return False, f"Input file not found: {input_path}"
    
    # Generate output path if not provided
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    
    # Try pdf2docx
    try:
        from pdf2docx import Converter
        
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        
        if os.path.exists(output_path):
            return True, output_path
    except ImportError:
        return False, "pdf2docx not installed. Run: pip install pdf2docx"
    except Exception as e:
        return False, f"PDF to DOCX conversion failed: {e}"
    
    return False, "PDF to DOCX conversion failed"


def html_to_pdf(input_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
    """
    Convert HTML to PDF.
    
    Args:
        input_path: Path to input HTML file
        output_path: Path for output PDF (auto-generated if None)
    
    Returns:
        Tuple of (success, message or output_path)
    """
    if not os.path.exists(input_path):
        return False, f"Input file not found: {input_path}"
    
    # Generate output path if not provided
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    
    # Try weasyprint
    try:
        from weasyprint import HTML
        HTML(input_path).write_pdf(output_path)
        if os.path.exists(output_path):
            return True, output_path
    except ImportError:
        logger.debug("weasyprint not available")
    except Exception as e:
        logger.debug(f"weasyprint failed: {e}")
    
    # Try wkhtmltopdf
    try:
        result = subprocess.run(
            ["wkhtmltopdf", input_path, output_path],
            capture_output=True,
            text=True,
            timeout=120
        )
        if result.returncode == 0 and os.path.exists(output_path):
            return True, output_path
    except FileNotFoundError:
        logger.debug("wkhtmltopdf not found")
    except Exception as e:
        logger.debug(f"wkhtmltopdf failed: {e}")
    
    return False, "No HTML to PDF converter available. Install weasyprint or wkhtmltopdf."
