"""
Core document processor class
"""

import os
import json
import uuid
import hmac
import hashlib
import logging
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from mcp_documents.core.formatting import apply_paragraph_formatting, apply_run_formatting
from mcp_documents.core.utils import parse_color, logger


class DocumentProcessor:
    """
    Unified document processor for Word documents.
    Manages document lifecycle and provides all editing operations.
    """
    
    def __init__(self):
        self.documents: Dict[str, Document] = {}  # Cache of opened documents
        self.current_document: Optional[Document] = None
        self.current_file_path: Optional[str] = None
        self.footnotes: List[str] = []  # Simple footnote store
        self.comments: List[Dict[str, Any]] = []  # Simple comment store
        # Sandbox and security controls
        self.sandbox_enabled: bool = False
        self.sandbox_root: Optional[str] = None
        self.readonly_mode: bool = False
        # Collaboration sessions (in-memory)
        self.collaboration_sessions: Dict[str, Dict[str, Any]] = {}

    # ==================== Internal Helpers ====================

    def _normalize_path(self, path: str) -> str:
        return os.path.abspath(os.path.expanduser(path))

    def _ensure_path_allowed(self, path: str, for_write: bool = False) -> Optional[str]:
        """Validate sandbox and readonly constraints."""
        normalized = self._normalize_path(path)
        if self.sandbox_enabled and self.sandbox_root:
            root = self._normalize_path(self.sandbox_root)
            if not os.path.commonpath([normalized, root]) == root:
                return f"Error: Path not allowed outside sandbox: {normalized}"
        if for_write and self.readonly_mode:
            return "Error: Read-only mode is enabled"
        return None
    
    # ==================== Document Lifecycle ====================
    
    def create_document(self, file_path: str, title: Optional[str] = None) -> str:
        """Create a new Word document."""
        try:
            err = self._ensure_path_allowed(file_path, for_write=True)
            if err:
                return err
            doc = Document()
            
            if title:
                doc.add_heading(title, 0)
            
            doc.save(file_path)
            
            self.current_document = doc
            self.current_file_path = file_path
            self.documents[file_path] = doc
            
            return f"Document created: {file_path}"
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            return f"Error: {e}"
    
    def open_document(self, file_path: str) -> str:
        """Open an existing Word document."""
        try:
            err = self._ensure_path_allowed(file_path, for_write=False)
            if err:
                return err
            if not os.path.exists(file_path):
                return f"Error: File not found: {file_path}"
            
            doc = Document(file_path)
            
            self.current_document = doc
            self.current_file_path = file_path
            self.documents[file_path] = doc
            
            return f"Document opened: {file_path}"
        except Exception as e:
            logger.error(f"Failed to open document: {e}")
            return f"Error: {e}"
    
    def save_document(self, file_path: Optional[str] = None) -> str:
        """Save the current document."""
        try:
            if not self.current_document:
                return "Error: No document is open"
            
            save_path = file_path or self.current_file_path
            if not save_path:
                return "Error: No file path specified"
            err = self._ensure_path_allowed(save_path, for_write=True)
            if err:
                return err
            
            self.current_document.save(save_path)
            
            if file_path:
                self.current_file_path = file_path
                self.documents[file_path] = self.current_document
            
            return f"Document saved: {save_path}"
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return f"Error: {e}"
    
    def close_document(self) -> str:
        """Close the current document."""
        if self.current_file_path and self.current_file_path in self.documents:
            del self.documents[self.current_file_path]
        
        self.current_document = None
        self.current_file_path = None
        
        return "Document closed"

    # Sandbox and readonly controls
    def enable_sandbox(self, base_dir: str) -> str:
        """Restrict file operations to the given base directory."""
        try:
            normalized = self._normalize_path(base_dir)
            if not os.path.isdir(normalized):
                return f"Error: Sandbox root not found: {normalized}"
            self.sandbox_enabled = True
            self.sandbox_root = normalized
            return f"Sandbox enabled at {normalized}"
        except Exception as e:
            logger.error(f"Failed to enable sandbox: {e}")
            return f"Error: {e}"

    def disable_sandbox(self) -> str:
        """Disable sandbox enforcement."""
        self.sandbox_enabled = False
        self.sandbox_root = None
        return "Sandbox disabled"

    def set_readonly_mode(self, enabled: bool = True) -> str:
        """Toggle readonly mode; when enabled, write operations are blocked."""
        self.readonly_mode = bool(enabled)
        return "Readonly mode enabled" if enabled else "Readonly mode disabled"
    
    def get_document_info(self) -> str:
        """Get information about the current document."""
        if not self.current_document:
            return "Error: No document is open"
        
        doc = self.current_document
        
        info = []
        info.append(f"File: {self.current_file_path}")
        info.append(f"Sections: {len(doc.sections)}")
        info.append(f"Paragraphs: {len(doc.paragraphs)}")
        info.append(f"Tables: {len(doc.tables)}")
        
        # Word count
        word_count = sum(len(p.text.split()) for p in doc.paragraphs)
        info.append(f"Word count: {word_count}")
        
        return "\n".join(info)
    
    # ==================== Content Operations ====================
    
    def add_paragraph(
        self,
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
        """Add a paragraph with optional formatting."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            # Add paragraph with optional style
            if style:
                try:
                    paragraph = self.current_document.add_paragraph(text, style=style)
                except KeyError:
                    paragraph = self.current_document.add_paragraph(text)
            else:
                paragraph = self.current_document.add_paragraph(text)
            
            # Apply text formatting to runs
            if paragraph.runs:
                run = paragraph.runs[0]
                run.bold = bold
                run.italic = italic
                run.underline = underline
                
                if font_size:
                    run.font.size = Pt(font_size)
                
                if font_name:
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                
                if color:
                    rgb = parse_color(color)
                    if rgb:
                        run.font.color.rgb = RGBColor(*rgb)
            
            # Apply paragraph alignment
            if alignment:
                alignment_map = {
                    "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
                    "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
                    "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
                    "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                }
                if alignment.lower() in alignment_map:
                    paragraph.alignment = alignment_map[alignment.lower()]
            
            return "Paragraph added"
        except Exception as e:
            logger.error(f"Failed to add paragraph: {e}")
            return f"Error: {e}"
    
    def add_heading(self, text: str, level: int = 1) -> str:
        """Add a heading to the document."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            self.current_document.add_heading(text, level=level)
            return f"Heading level {level} added"
        except Exception as e:
            logger.error(f"Failed to add heading: {e}")
            return f"Error: {e}"
    
    def add_table(
        self,
        rows: int,
        cols: int,
        data: Optional[List[List[str]]] = None,
        style: str = "Table Grid"
    ) -> str:
        """Add a table to the document."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            table = self.current_document.add_table(rows=rows, cols=cols, style=style)
            
            # Fill table data
            if data:
                for i, row_data in enumerate(data):
                    if i < rows:
                        row = table.rows[i]
                        for j, cell_text in enumerate(row_data):
                            if j < cols:
                                row.cells[j].text = str(cell_text)
            
            return f"Table {rows}x{cols} added"
        except Exception as e:
            logger.error(f"Failed to add table: {e}")
            return f"Error: {e}"
    
    def add_image(
        self,
        image_path: str,
        width_inches: Optional[float] = None,
        height_inches: Optional[float] = None
    ) -> str:
        """Add an image to the document."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            if not os.path.exists(image_path):
                return f"Error: Image not found: {image_path}"
            
            width = Inches(width_inches) if width_inches else None
            height = Inches(height_inches) if height_inches else None
            
            self.current_document.add_picture(image_path, width=width, height=height)
            
            return "Image added"
        except Exception as e:
            logger.error(f"Failed to add image: {e}")
            return f"Error: {e}"
    
    def add_page_break(self) -> str:
        """Add a page break."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            self.current_document.add_page_break()
            return "Page break added"
        except Exception as e:
            logger.error(f"Failed to add page break: {e}")
            return f"Error: {e}"

    # ==================== Header & Footer ====================

    def _set_paragraph_alignment(self, paragraph, alignment: str) -> None:
        """Helper to set paragraph alignment safely."""
        align_map = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        }
        if alignment and alignment.lower() in align_map:
            paragraph.alignment = align_map[alignment.lower()]

    def _clear_paragraph(self, paragraph) -> None:
        """Remove all child elements from a paragraph."""
        p = paragraph._p
        for child in list(p):
            p.remove(child)

    def add_header(self, text: str, section_index: int = 0, alignment: str = "center") -> str:
        """Add/update header text for a section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            header = sections[section_index].header
            if header.paragraphs:
                para = header.paragraphs[0]
                self._clear_paragraph(para)
            else:
                para = header.add_paragraph()
            para.add_run(text)
            self._set_paragraph_alignment(para, alignment)
            return f"Header set for section {section_index}"
        except Exception as e:
            logger.error(f"Failed to add header: {e}")
            return f"Error: {e}"

    def add_footer(self, text: str, section_index: int = 0, alignment: str = "center") -> str:
        """Add/update footer text for a section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            footer = sections[section_index].footer
            if footer.paragraphs:
                para = footer.paragraphs[0]
                self._clear_paragraph(para)
            else:
                para = footer.add_paragraph()
            para.add_run(text)
            self._set_paragraph_alignment(para, alignment)
            return f"Footer set for section {section_index}"
        except Exception as e:
            logger.error(f"Failed to add footer: {e}")
            return f"Error: {e}"

    def _add_page_field(self, paragraph) -> None:
        """Insert a PAGE field into a paragraph."""
        fld_char_begin = OxmlElement("w:fldChar")
        fld_char_begin.set(qn("w:fldCharType"), "begin")

        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = "PAGE"

        fld_char_separate = OxmlElement("w:fldChar")
        fld_char_separate.set(qn("w:fldCharType"), "separate")

        fld_char_end = OxmlElement("w:fldChar")
        fld_char_end.set(qn("w:fldCharType"), "end")

        for el in (fld_char_begin, instr_text, fld_char_separate, fld_char_end):
            paragraph._p.append(el)

    def add_page_numbers(
        self,
        position: str = "footer",
        alignment: str = "center",
        format_string: str = "Page {page}",
        section_index: int = 0
    ) -> str:
        """Add simple page numbers to header or footer."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"

            container = sections[section_index].footer if position.lower() == "footer" else sections[section_index].header
            if container.paragraphs:
                para = container.paragraphs[0]
                self._clear_paragraph(para)
            else:
                para = container.add_paragraph()
            self._set_paragraph_alignment(para, alignment)

            # Add format prefix if provided
            prefix = format_string.replace("{page}", "").strip()
            if prefix:
                para.add_run(prefix + " ")
            self._add_page_field(para)
            return f"Page numbers added to {position}"
        except Exception as e:
            logger.error(f"Failed to add page numbers: {e}")
            return f"Error: {e}"

    def remove_header(self, section_index: int = 0) -> str:
        """Remove header content from a section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            header = sections[section_index].header
            for para in list(header.paragraphs):
                p = para._element
                p.getparent().remove(p)
            return f"Header removed from section {section_index}"
        except Exception as e:
            logger.error(f"Failed to remove header: {e}")
            return f"Error: {e}"

    def remove_footer(self, section_index: int = 0) -> str:
        """Remove footer content from a section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            footer = sections[section_index].footer
            for para in list(footer.paragraphs):
                p = para._element
                p.getparent().remove(p)
            return f"Footer removed from section {section_index}"
        except Exception as e:
            logger.error(f"Failed to remove footer: {e}")
            return f"Error: {e}"

    def get_header_text(self, section_index: int = 0) -> str:
        """Get header text."""
        if not self.current_document:
            return "Error: No document is open"
        sections = self.current_document.sections
        if section_index >= len(sections):
            return f"Error: Section index out of range"
        header = sections[section_index].header
        texts = [p.text for p in header.paragraphs if p.text]
        return "\n".join(texts) or "(empty)"

    def get_footer_text(self, section_index: int = 0) -> str:
        """Get footer text."""
        if not self.current_document:
            return "Error: No document is open"
        sections = self.current_document.sections
        if section_index >= len(sections):
            return f"Error: Section index out of range"
        footer = sections[section_index].footer
        texts = [p.text for p in footer.paragraphs if p.text]
        return "\n".join(texts) or "(empty)"

    def set_different_first_page_header(
        self,
        first_page_text: str,
        other_pages_text: str,
        section_index: int = 0
    ) -> str:
        """Set different header for first page and other pages."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            section = sections[section_index]
            section.different_first_page_header_footer = True

            # First page header
            first_header = section.first_page_header
            if first_header.paragraphs:
                para = first_header.paragraphs[0]
                self._clear_paragraph(para)
            else:
                para = first_header.add_paragraph()
            para.add_run(first_page_text)
            self._set_paragraph_alignment(para, "center")

            # Default header
            default_header = section.header
            if default_header.paragraphs:
                para2 = default_header.paragraphs[0]
                self._clear_paragraph(para2)
            else:
                para2 = default_header.add_paragraph()
            para2.add_run(other_pages_text)
            self._set_paragraph_alignment(para2, "center")

            return "Different first-page header set"
        except Exception as e:
            logger.error(f"Failed to set different header: {e}")
            return f"Error: {e}"
    
    # ==================== Table Operations ====================
    
    def edit_table_cell(
        self,
        table_index: int,
        row_index: int,
        col_index: int,
        text: str
    ) -> str:
        """Edit a table cell's content."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range (have {len(tables)} tables)"
            
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Row index out of range"
            if col_index >= len(table.columns):
                return f"Error: Column index out of range"
            
            table.cell(row_index, col_index).text = text
            return f"Cell ({row_index},{col_index}) updated"
        except Exception as e:
            logger.error(f"Failed to edit table cell: {e}")
            return f"Error: {e}"
    
    def add_table_row(self, table_index: int, data: Optional[List[str]] = None) -> str:
        """Add a row to a table."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            
            table = tables[table_index]
            new_row = table.add_row()
            
            if data:
                for i, cell_text in enumerate(data):
                    if i < len(new_row.cells):
                        new_row.cells[i].text = str(cell_text)
            
            return "Row added"
        except Exception as e:
            logger.error(f"Failed to add table row: {e}")
            return f"Error: {e}"
    
    def delete_table_row(self, table_index: int, row_index: int) -> str:
        """Delete a row from a table."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Row index out of range"
            
            row = table.rows[row_index]._tr
            row.getparent().remove(row)
            
            return f"Row {row_index} deleted"
        except Exception as e:
            logger.error(f"Failed to delete table row: {e}")
            return f"Error: {e}"
    
    def merge_table_cells(
        self,
        table_index: int,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int
    ) -> str:
        """Merge table cells."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            
            table = tables[table_index]
            start_cell = table.cell(start_row, start_col)
            end_cell = table.cell(end_row, end_col)
            start_cell.merge(end_cell)
            
            return f"Cells merged from ({start_row},{start_col}) to ({end_row},{end_col})"
        except Exception as e:
            logger.error(f"Failed to merge cells: {e}")
            return f"Error: {e}"
    
    # ==================== Search and Edit ====================
    
    def search_text(self, keyword: str) -> str:
        """Search for text in the document."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            results = []
            
            # Search paragraphs
            for i, para in enumerate(self.current_document.paragraphs):
                if keyword in para.text:
                    results.append(f"Paragraph {i}: {para.text[:100]}...")
            
            # Search tables
            for t_idx, table in enumerate(self.current_document.tables):
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if keyword in cell.text:
                            results.append(f"Table {t_idx} cell ({r_idx},{c_idx}): {cell.text[:50]}...")
            
            if not results:
                return f"'{keyword}' not found"
            
            return f"Found {len(results)} matches:\n" + "\n".join(results)
        except Exception as e:
            logger.error(f"Search failed: {e}")
            return f"Error: {e}"
    
    def find_and_replace(self, find_text: str, replace_text: str) -> str:
        """Find and replace text in the document."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            count = 0
            
            # Replace in paragraphs
            for para in self.current_document.paragraphs:
                if find_text in para.text:
                    occurrences = para.text.count(find_text)
                    para.text = para.text.replace(find_text, replace_text)
                    count += occurrences
            
            # Replace in tables
            for table in self.current_document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if find_text in para.text:
                                occurrences = para.text.count(find_text)
                                para.text = para.text.replace(find_text, replace_text)
                                count += occurrences
            
            return f"Replaced {count} occurrences"
        except Exception as e:
            logger.error(f"Replace failed: {e}")
            return f"Error: {e}"
    
    def delete_paragraph(self, index: int) -> str:
        """Delete a paragraph by index."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            paragraphs = self.current_document.paragraphs
            if index >= len(paragraphs):
                return f"Error: Paragraph index out of range"
            
            p = paragraphs[index]._element
            p.getparent().remove(p)
            
            return f"Paragraph {index} deleted"
        except Exception as e:
            logger.error(f"Failed to delete paragraph: {e}")
            return f"Error: {e}"
    
    # ==================== Page Layout ====================
    
    def set_page_margins(
        self,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
        section_index: int = 0
    ) -> str:
        """Set page margins in centimeters."""
        if not self.current_document:
            return "Error: No document is open"
        
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            
            section = sections[section_index]
            
            if top is not None:
                section.top_margin = Cm(top)
            if bottom is not None:
                section.bottom_margin = Cm(bottom)
            if left is not None:
                section.left_margin = Cm(left)
            if right is not None:
                section.right_margin = Cm(right)
            
            return "Page margins updated"
        except Exception as e:
            logger.error(f"Failed to set margins: {e}")
            return f"Error: {e}"
    
    # ==================== Paragraphs and Sections Access ====================
    
    def get_paragraphs(self) -> List[str]:
        """Get list of all paragraph texts."""
        if not self.current_document:
            return []
        return [p.text for p in self.current_document.paragraphs]
    
    def get_tables_info(self) -> str:
        """Get information about all tables."""
        if not self.current_document:
            return "Error: No document is open"
        
        tables = self.current_document.tables
        if not tables:
            return "No tables in document"
        
        info = []
        for i, table in enumerate(tables):
            rows = len(table.rows)
            cols = len(table.columns)
            info.append(f"Table {i}: {rows} rows x {cols} columns")
        
        return "\n".join(info)

    # ==================== List Operations ====================

    def add_bulleted_list(self, items: List[str], style: str = "List Bullet") -> str:
        """Add a bulleted list."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            for text in items:
                self.current_document.add_paragraph(text, style=style)
            return f"Bulleted list added with {len(items)} items"
        except Exception as e:
            logger.error(f"Failed to add bulleted list: {e}")
            return f"Error: {e}"

    def add_numbered_list(self, items: List[str], style: str = "List Number") -> str:
        """Add a numbered list."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            for text in items:
                self.current_document.add_paragraph(text, style=style)
            return f"Numbered list added with {len(items)} items"
        except Exception as e:
            logger.error(f"Failed to add numbered list: {e}")
            return f"Error: {e}"

    def add_list_item(self, text: str, list_type: str = "bullet") -> str:
        """Add a single list item using bullet or number style."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            style = "List Bullet" if list_type.lower() == "bullet" else "List Number"
            self.current_document.add_paragraph(text, style=style)
            return "List item added"
        except Exception as e:
            logger.error(f"Failed to add list item: {e}")
            return f"Error: {e}"

    def add_multilevel_list(self, items: List[Dict[str, Any]]) -> str:
        """Add a multi-level list based on provided levels."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            for item in items:
                text = item.get("text", "")
                level = int(item.get("level", 0))
                style = "List Bullet" if level == 0 else "List Bullet %d" % (level + 1)
                self.current_document.add_paragraph(text, style=style)
            return f"Multilevel list added with {len(items)} items"
        except Exception as e:
            logger.error(f"Failed to add multilevel list: {e}")
            return f"Error: {e}"

    # ==================== Footnotes ====================

    def add_footnote(self, text: str, paragraph_index: int) -> str:
        """Add a simple footnote (inline reference + footnote section)."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if not paragraphs:
                self.add_paragraph("")
                paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                paragraph_index = len(paragraphs) - 1

            footnote_number = len(self.footnotes) + 1
            target_para = paragraphs[paragraph_index]
            run = target_para.add_run(f"[{footnote_number}]")
            run.font.superscript = True

            # Store footnote and ensure footnote section exists
            self.footnotes.append(text)
            self._ensure_footnote_section()
            return f"Footnote {footnote_number} added"
        except Exception as e:
            logger.error(f"Failed to add footnote: {e}")
            return f"Error: {e}"

    def _ensure_footnote_section(self) -> None:
        """Append/refresh footnote section at document end."""
        # Remove existing custom footnote section if present
        marker = "__MCPDOCX_FOOTNOTES__"
        paras = self.current_document.paragraphs
        existing_index = None
        for idx, para in enumerate(paras):
            if marker in para.text:
                existing_index = idx
                break
        if existing_index is not None:
            # Clear existing footnote content
            for para in list(paras[existing_index:]):
                p = para._element
                p.getparent().remove(p)

        # Recreate section
        self.current_document.add_page_break()
        title_para = self.current_document.add_paragraph("Footnotes")
        title_para.runs[0].bold = True
        title_para.add_run(f" {marker}")
        for i, note in enumerate(self.footnotes, start=1):
            p = self.current_document.add_paragraph()
            p.add_run(f"[{i}] ")
            p.add_run(note)

    # ==================== Comments (lightweight) ====================

    def add_comment(self, comment_text: str, paragraph_index: int, author: str = "MCP Documents", initials: str = "MCP") -> str:
        """Add a lightweight comment marker to a paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if not paragraphs:
                self.add_paragraph("")
                paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                paragraph_index = len(paragraphs) - 1

            marker = f"[{author}: {comment_text}]"
            para = paragraphs[paragraph_index]
            run = para.add_run(f" {marker}")
            run.font.italic = True
            run.font.color.rgb = RGBColor(120, 120, 120)

            self.comments.append({
                "paragraph": paragraph_index,
                "author": author,
                "initials": initials,
                "text": comment_text,
            })
            return "Comment added"
        except Exception as e:
            logger.error(f"Failed to add comment: {e}")
            return f"Error: {e}"

    def get_all_comments(self) -> str:
        if not self.comments:
            return "No comments in document"
        lines = []
        for idx, c in enumerate(self.comments):
            lines.append(f"{idx}: Para {c['paragraph']} by {c['author']} - {c['text']}")
        return "\n".join(lines)

    def get_comments_by_author(self, author: str) -> str:
        filtered = [c for c in self.comments if c.get("author") == author]
        if not filtered:
            return f"No comments by {author}"
        lines = []
        for idx, c in enumerate(filtered):
            lines.append(f"Para {c['paragraph']} - {c['text']}")
        return "\n".join(lines)

    def delete_all_comments(self) -> str:
        self.comments = []
        # Remove comment markers from document text
        if self.current_document:
            for para in self.current_document.paragraphs:
                if "[" in para.text and "]" in para.text:
                    for run in para.runs:
                        if "[" in run.text and "]" in run.text:
                            run.text = run.text.split("[")[0].strip()
        return "All comments removed"

    def get_comment_count(self) -> str:
        return f"Document has {len(self.comments)} comments"

    # ==================== Hyperlinks ====================

    def add_hyperlink(self, paragraph_index: int, text: str, url: str, position: str = "end") -> str:
        """Insert a hyperlink into a paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Paragraph index out of range"

            paragraph = paragraphs[paragraph_index]
            part = paragraph.part
            r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

            hyperlink = OxmlElement('w:hyperlink')
            hyperlink.set(qn('r:id'), r_id)

            new_run = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            # style: hyperlink (blue + underline)
            rStyle = OxmlElement('w:rStyle')
            rStyle.set(qn('w:val'), 'Hyperlink')
            rPr.append(rStyle)
            new_run.append(rPr)

            t = OxmlElement('w:t')
            t.text = text
            new_run.append(t)
            hyperlink.append(new_run)

            if position == "start":
                paragraph._p.insert(0, hyperlink)
            else:
                paragraph._p.append(hyperlink)

            return "Hyperlink added"
        except Exception as e:
            logger.error(f"Failed to add hyperlink: {e}")
            return f"Error: {e}"

    # ==================== Advanced Table Formatting ====================

    def set_table_cell_shading(self, table_index: int, row_index: int, col_index: int, fill_color: str = "FFFF00") -> str:
        """Set a cell's background color."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            cell = table.cell(row_index, col_index)
            tc_pr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), fill_color.replace('#', ''))
            tc_pr.append(shd)
            return "Cell shading applied"
        except Exception as e:
            logger.error(f"Failed to set shading: {e}")
            return f"Error: {e}"

    def apply_table_alternating_rows(self, table_index: int, color1: str = "FFFFFF", color2: str = "F2F2F2") -> str:
        """Apply alternating row colors to a table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            for idx, row in enumerate(table.rows):
                shade = color1 if idx % 2 == 0 else color2
                for cell in row.cells:
                    tc_pr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), shade)
                    tc_pr.append(shd)
            return "Alternating row colors applied"
        except Exception as e:
            logger.error(f"Failed to apply alternating rows: {e}")
            return f"Error: {e}"

    def highlight_table_header(self, table_index: int, header_color: str = "4472C4", text_color: str = "FFFFFF") -> str:
        """Highlight first row as header."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            if not table.rows:
                return "Error: Table has no rows"
            header_row = table.rows[0]
            for cell in header_row.cells:
                tc_pr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:fill'), header_color)
                tc_pr.append(shd)
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        rgb = parse_color(text_color)
                        if rgb:
                            run.font.color.rgb = RGBColor(*rgb)
            return "Header highlighted"
        except Exception as e:
            logger.error(f"Failed to highlight header: {e}")
            return f"Error: {e}"

    def set_cell_alignment(self, table_index: int, row_index: int, col_index: int, horizontal: str = "center", vertical: str = "center") -> str:
        """Set cell horizontal and vertical alignment."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            cell = table.cell(row_index, col_index)
            # Horizontal via paragraph alignment
            for para in cell.paragraphs:
                self._set_paragraph_alignment(para, horizontal)
            # Vertical alignment
            v_map = {
                "top": WD_ALIGN_VERTICAL.TOP,
                "center": WD_ALIGN_VERTICAL.CENTER,
                "bottom": WD_ALIGN_VERTICAL.BOTTOM,
            }
            if vertical.lower() in v_map:
                cell.vertical_alignment = v_map[vertical.lower()]
            return "Cell alignment set"
        except Exception as e:
            logger.error(f"Failed to set cell alignment: {e}")
            return f"Error: {e}"

    def set_cell_padding(self, table_index: int, row_index: int, col_index: int, padding_cm: float = 0.1) -> str:
        """Set padding for a cell."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            cell = table.cell(row_index, col_index)
            tc_pr = cell._tc.get_or_add_tcPr()
            mar = tc_pr.find(qn('w:tcMar'))
            if mar is None:
                mar = OxmlElement('w:tcMar')
                tc_pr.append(mar)
            for side in ['top', 'bottom', 'start', 'end']:
                element = OxmlElement(f'w:{side}')
                element.set(qn('w:w'), str(int(padding_cm * 567)))  # approx cm to twentieths of a point
                element.set(qn('w:type'), 'dxa')
                mar.append(element)
            return "Cell padding set"
        except Exception as e:
            logger.error(f"Failed to set padding: {e}")
            return f"Error: {e}"

    def set_column_width(self, table_index: int, col_index: int, width_cm: float) -> str:
        """Set width for a column across all rows."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Table index out of range"
            table = tables[table_index]
            for row in table.rows:
                if col_index >= len(row.cells):
                    return f"Error: Column index out of range"
                row.cells[col_index].width = Cm(width_cm)
            return f"Column {col_index} width set to {width_cm}cm"
        except Exception as e:
            logger.error(f"Failed to set column width: {e}")
            return f"Error: {e}"

    # ==================== Watermark ====================

    def add_text_watermark(self, text: str = "CONFIDENTIAL", section_index: int = 0) -> str:
        """Add a simple text watermark via header."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            header = sections[section_index].header
            para = header.add_paragraph()
            run = para.add_run(text)
            run.font.size = Pt(48)
            run.font.color.rgb = RGBColor(200, 200, 200)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            return "Watermark added (header-based)"
        except Exception as e:
            logger.error(f"Failed to add watermark: {e}")
            return f"Error: {e}"

    # ==================== Document Merge ====================

    def merge_documents(self, file_paths: List[str]) -> str:
        """Append content from multiple documents into the current document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            count = 0
            for path in file_paths:
                err = self._ensure_path_allowed(path, for_write=False)
                if err:
                    return err
                if not os.path.exists(path):
                    return f"Error: File not found: {path}"
                src = Document(path)
                self._append_document_content(src)
                count += 1
            return f"Merged {count} documents"
        except Exception as e:
            logger.error(f"Failed to merge documents: {e}")
            return f"Error: {e}"

    def insert_document(self, file_path: str) -> str:
        """Insert another document at the end of current document."""
        return self.merge_documents([file_path])

    def _append_document_content(self, src_doc: Document) -> None:
        """Append paragraphs and tables from another document."""
        # Append paragraphs
        for para in src_doc.paragraphs:
            new_para = self.current_document.add_paragraph()
            if para.style:
                new_para.style = para.style
            for run in para.runs:
                new_run = new_para.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                if run.font and run.font.size:
                    new_run.font.size = run.font.size
        # Append tables
        for tbl in src_doc.tables:
            rows = len(tbl.rows)
            cols = len(tbl.columns)
            new_tbl = self.current_document.add_table(rows=rows, cols=cols)
            for r_idx, row in enumerate(tbl.rows):
                for c_idx, cell in enumerate(row.cells):
                    new_tbl.cell(r_idx, c_idx).text = cell.text

    # ==================== Track Changes ====================

    def enable_track_changes(self) -> str:
        """Enable track changes in the document via documentProtection element."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            # Look for existing trackRevisions element
            existing = settings.find(qn('w:trackRevisions'))
            if existing is None:
                track_elem = OxmlElement('w:trackRevisions')
                settings.append(track_elem)
            return "Track changes enabled"
        except Exception as e:
            logger.error(f"Failed to enable track changes: {e}")
            return f"Error: {e}"

    def disable_track_changes(self) -> str:
        """Disable track changes in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            existing = settings.find(qn('w:trackRevisions'))
            if existing is not None:
                settings.remove(existing)
            return "Track changes disabled"
        except Exception as e:
            logger.error(f"Failed to disable track changes: {e}")
            return f"Error: {e}"

    def accept_all_changes(self) -> str:
        """Accept all tracked changes by removing revision marks (simplified)."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            # Find and remove all w:del elements (deletions) and remove w:ins wrappers
            body = self.current_document.element.body
            # Remove deletions
            for del_elem in body.findall('.//'+qn('w:del')):
                parent = del_elem.getparent()
                parent.remove(del_elem)
            # Unwrap insertions (keep content)
            for ins_elem in body.findall('.//'+qn('w:ins')):
                parent = ins_elem.getparent()
                index = list(parent).index(ins_elem)
                for child in ins_elem:
                    parent.insert(index, child)
                    index += 1
                parent.remove(ins_elem)
            return "All changes accepted"
        except Exception as e:
            logger.error(f"Failed to accept changes: {e}")
            return f"Error: {e}"

    def reject_all_changes(self) -> str:
        """Reject all tracked changes (simplified)."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            body = self.current_document.element.body
            # Remove insertions
            for ins_elem in body.findall('.//'+qn('w:ins')):
                parent = ins_elem.getparent()
                parent.remove(ins_elem)
            # Unwrap deletions (restore content)
            for del_elem in body.findall('.//'+qn('w:del')):
                parent = del_elem.getparent()
                index = list(parent).index(del_elem)
                for child in del_elem:
                    parent.insert(index, child)
                    index += 1
                parent.remove(del_elem)
            return "All changes rejected"
        except Exception as e:
            logger.error(f"Failed to reject changes: {e}")
            return f"Error: {e}"

    def add_revision_insert(self, paragraph_index: int, text: str, author: str = "User",
                             timestamp: Optional[str] = None) -> str:
        """Insert text wrapped in w:ins to simulate a tracked insertion."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return "Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            ins = OxmlElement('w:ins')
            ins.set(qn('w:id'), str(uuid.uuid4().int % 65535))
            ins.set(qn('w:author'), author)
            ins.set(qn('w:date'), timestamp or datetime.utcnow().isoformat() + "Z")
            run = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = text
            run.append(t)
            ins.append(run)
            para._p.append(ins)
            return f"Revision insert added to paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add revision insert: {e}")
            return f"Error: {e}"

    def add_revision_delete(self, paragraph_index: int, start: int = 0, end: Optional[int] = None,
                             author: str = "User", timestamp: Optional[str] = None) -> str:
        """Wrap a text slice in w:del to simulate a tracked deletion."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return "Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            text = para.text
            end = end if end is not None else len(text)
            if start < 0 or end > len(text) or start >= end:
                return "Error: Invalid start/end positions"
            before, target, after = text[:start], text[start:end], text[end:]
            for child in list(para._p):
                para._p.remove(child)
            if before:
                para.add_run(before)
            del_elem = OxmlElement('w:del')
            del_elem.set(qn('w:id'), str(uuid.uuid4().int % 65535))
            del_elem.set(qn('w:author'), author)
            del_elem.set(qn('w:date'), timestamp or datetime.utcnow().isoformat() + "Z")
            del_run = OxmlElement('w:r')
            del_text = OxmlElement('w:t')
            del_text.text = target
            del_run.append(del_text)
            del_elem.append(del_run)
            para._p.append(del_elem)
            if after:
                para.add_run(after)
            return f"Revision deletion added to paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add revision deletion: {e}")
            return f"Error: {e}"

    # ==================== Document Protection ====================

    def protect_document(self, password: Optional[str] = None, protection_type: str = "readOnly") -> str:
        """
        Protect document with optional password.
        protection_type: 'readOnly', 'comments', 'trackedChanges', 'forms'
        """
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            # Remove existing protection
            existing = settings.find(qn('w:documentProtection'))
            if existing is not None:
                settings.remove(existing)
            
            prot = OxmlElement('w:documentProtection')
            edit_map = {
                'readOnly': 'readOnly',
                'comments': 'comments',
                'trackedChanges': 'trackedChanges',
                'forms': 'forms'
            }
            prot.set(qn('w:edit'), edit_map.get(protection_type, 'readOnly'))
            prot.set(qn('w:enforcement'), '1')
            
            if password:
                # Simple hash for demo - in production use proper algorithm
                import hashlib
                hash_val = hashlib.sha256(password.encode()).hexdigest()[:16]
                prot.set(qn('w:hash'), hash_val)
                prot.set(qn('w:cryptAlgorithmSid'), '14')  # SHA-512
            
            settings.append(prot)
            return f"Document protected with type: {protection_type}"
        except Exception as e:
            logger.error(f"Failed to protect document: {e}")
            return f"Error: {e}"

    def unprotect_document(self) -> str:
        """Remove document protection."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            existing = settings.find(qn('w:documentProtection'))
            if existing is not None:
                settings.remove(existing)
                return "Document protection removed"
            return "Document was not protected"
        except Exception as e:
            logger.error(f"Failed to unprotect document: {e}")
            return f"Error: {e}"

    # ==================== Digital Signatures (Sidecar) ====================

    def _signature_sidecar_path(self, signature_name: str = "default") -> Optional[str]:
        if not self.current_file_path:
            return None
        base = os.path.splitext(self.current_file_path)[0]
        safe_name = signature_name.replace("/", "_")
        return f"{base}.sig-{safe_name}.json"

    def sign_document(self, signature_name: str = "default", signer: str = "unknown",
                       secret: Optional[str] = None) -> str:
        """Create a detached signature sidecar (SHA-256 with optional HMAC)."""
        if not self.current_file_path or not os.path.exists(self.current_file_path):
            return "Error: No saved document to sign"
        try:
            sidecar = self._signature_sidecar_path(signature_name)
            err = self._ensure_path_allowed(sidecar, for_write=True) if sidecar else None
            if err:
                return err
            with open(self.current_file_path, "rb") as f:
                data = f.read()
            digest = hashlib.sha256(data).hexdigest()
            hmac_digest = None
            if secret:
                hmac_digest = hmac.new(secret.encode(), data, hashlib.sha256).hexdigest()
            payload = {
                "file": os.path.basename(self.current_file_path),
                "signature_name": signature_name,
                "signer": signer,
                "algorithm": "sha256",
                "digest": digest,
                "hmac": bool(secret),
                "hmac_digest": hmac_digest,
                "timestamp": datetime.utcnow().isoformat() + "Z",
            }
            with open(sidecar, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2)
            return f"Signature created: {sidecar}"
        except Exception as e:
            logger.error(f"Failed to sign document: {e}")
            return f"Error: {e}"

    def verify_document_signature(self, signature_name: str = "default", secret: Optional[str] = None) -> str:
        """Verify detached signature sidecar against current document."""
        if not self.current_file_path or not os.path.exists(self.current_file_path):
            return "Error: No saved document to verify"
        try:
            sidecar = self._signature_sidecar_path(signature_name)
            if not sidecar or not os.path.exists(sidecar):
                return f"Error: Signature file not found: {sidecar}"
            with open(sidecar, "r", encoding="utf-8") as f:
                payload = json.load(f)
            with open(self.current_file_path, "rb") as f:
                data = f.read()
            digest = hashlib.sha256(data).hexdigest()
            if digest != payload.get("digest"):
                return "Verification failed: digest mismatch"
            if secret and payload.get("hmac"):
                calc = hmac.new(secret.encode(), data, hashlib.sha256).hexdigest()
                if calc != payload.get("hmac_digest"):
                    return "Verification failed: HMAC mismatch"
            return "Signature verified"
        except Exception as e:
            logger.error(f"Failed to verify signature: {e}")
            return f"Error: {e}"

    # ==================== Native Footnotes/Endnotes ====================

    def add_native_footnote(self, paragraph_index: int, footnote_text: str) -> str:
        """Add a proper Word footnote to a paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            para = paragraphs[paragraph_index]
            
            # Create footnote reference
            run = para.add_run()
            footnote_ref = OxmlElement('w:footnoteReference')
            footnote_id = len(self.footnotes) + 1
            footnote_ref.set(qn('w:id'), str(footnote_id))
            run._r.append(footnote_ref)
            
            # Footnote content needs to be in footnotes.xml part
            # For simplicity, store in our list and add as superscript
            self.footnotes.append(footnote_text)
            super_run = para.add_run(f'[{footnote_id}:{footnote_text}]')
            super_run.font.superscript = True
            super_run.font.size = Pt(8)
            
            return f"Footnote {footnote_id} added"
        except Exception as e:
            logger.error(f"Failed to add native footnote: {e}")
            return f"Error: {e}"

    def add_endnote(self, paragraph_index: int, endnote_text: str) -> str:
        """Add an endnote marker and store for end of document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            para = paragraphs[paragraph_index]
            endnote_id = len(self.footnotes) + 100  # Endnotes start at 100
            
            run = para.add_run(f'[EN{endnote_id - 99}]')
            run.font.superscript = True
            run.font.size = Pt(8)
            
            self.footnotes.append(f"ENDNOTE:{endnote_text}")
            return f"Endnote {endnote_id - 99} added"
        except Exception as e:
            logger.error(f"Failed to add endnote: {e}")
            return f"Error: {e}"

    def render_endnotes(self) -> str:
        """Add endnotes section at end of document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            endnotes = [fn for fn in self.footnotes if fn.startswith("ENDNOTE:")]
            if not endnotes:
                return "No endnotes to render"
            
            self.current_document.add_paragraph()  # Space
            self.current_document.add_heading("Endnotes", level=2)
            
            for i, en in enumerate(endnotes, 1):
                text = en.replace("ENDNOTE:", "")
                self.current_document.add_paragraph(f"{i}. {text}")
            
            return f"Rendered {len(endnotes)} endnotes"
        except Exception as e:
            logger.error(f"Failed to render endnotes: {e}")
            return f"Error: {e}"

    # ==================== Table of Contents ====================

    def add_table_of_contents(self, title: str = "Table of Contents", 
                               heading_levels: int = 3) -> str:
        """Add a Table of Contents field to the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            # Add TOC title
            toc_title = self.current_document.add_paragraph(title)
            toc_title.style = 'Heading 1'
            
            # Create TOC field paragraph
            para = self.current_document.add_paragraph()
            run = para.add_run()
            
            # Add TOC field code
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = f' TOC \\o "1-{heading_levels}" \\h \\z \\u '
            
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'separate')
            
            fldChar3 = OxmlElement('w:fldChar')
            fldChar3.set(qn('w:fldCharType'), 'end')
            
            run._r.append(fldChar1)
            run._r.append(instrText)
            run._r.append(fldChar2)
            run2 = para.add_run("Right-click and select 'Update Field' to generate TOC")
            run2.italic = True
            run2.font.color.rgb = RGBColor(128, 128, 128)
            run3 = para.add_run()
            run3._r.append(fldChar3)
            
            # Add page break after TOC
            self.current_document.add_page_break()
            
            return "Table of Contents added (update in Word to populate)"
        except Exception as e:
            logger.error(f"Failed to add TOC: {e}")
            return f"Error: {e}"

    def update_toc(self) -> str:
        """Return instructions for updating TOC (must be done in Word)."""
        return "To update TOC: Open in Word, right-click the TOC, select 'Update Field'"

    # ==================== Mail Merge ====================

    def add_merge_field(self, field_name: str) -> str:
        """Add a mail merge field to the current document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            para = self.current_document.add_paragraph()
            run = para.add_run()
            
            # MERGEFIELD field code
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = f' MERGEFIELD {field_name} '
            
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'separate')
            
            fldChar3 = OxmlElement('w:fldChar')
            fldChar3.set(qn('w:fldCharType'), 'end')
            
            run._r.append(fldChar1)
            run._r.append(instrText)
            run._r.append(fldChar2)
            run2 = para.add_run(f'«{field_name}»')
            run3 = para.add_run()
            run3._r.append(fldChar3)
            
            return f"Merge field '{field_name}' added"
        except Exception as e:
            logger.error(f"Failed to add merge field: {e}")
            return f"Error: {e}"

    def execute_mail_merge(self, data: List[Dict[str, str]], output_pattern: str) -> str:
        """
        Execute mail merge with data, creating one document per record.
        output_pattern should include {index} e.g. 'letter_{index}.docx'
        """
        if not self.current_document:
            return "Error: No document is open"
        if not self.current_file_path:
            return "Error: No file path set"
        try:
            import re
            
            def _evaluate_if_field(instr: str, record_lookup: Dict[str, Any]) -> Optional[str]:
                """Evaluate simple IF MERGEFIELD expressions."""
                import re
                pattern = re.compile(
                    r"IF\s+MERGEFIELD\s+\"?([\w\-.]+)\"?\s*(=|<>|>|<)\s*\"?([^\"]+)\"?\s+\"([^\"]*)\"(?:\s+\"([^\"]*)\")?",
                    flags=re.IGNORECASE,
                )
                m = pattern.search(instr)
                if not m:
                    return None
                field, op, comparator, true_text, false_text = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5) or ""
                value = record_lookup.get(field.lower())
                lhs = str(value) if value is not None else ""
                rhs = comparator
                result = False
                try:
                    if op == "=":
                        result = lhs == rhs
                    elif op == "<>":
                        result = lhs != rhs
                    elif op == ">":
                        result = lhs > rhs
                    elif op == "<":
                        result = lhs < rhs
                except Exception:
                    result = False
                return true_text if result else false_text

            def _replace_mergefields_in_paragraph(para, record_map) -> int:
                """Replace MERGEFIELD placeholders (guillemets and field codes) in a paragraph."""
                replaced = 0
                record_lookup = {str(k).lower(): v for k, v in record_map.items()}

                # 1) Guillemets placeholder «Field»
                for field_name, value in record_map.items():
                    placeholder = f'«{field_name}»'
                    if placeholder in para.text:
                        for run in para.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(value))
                                replaced += 1

                # 2) fldSimple elements (w:fldSimple instr="MERGEFIELD field")
                for fld_simple in para._p.findall('.//'+qn('w:fldSimple')):
                    instr = fld_simple.get(qn('w:instr'))
                    if not instr:
                        continue
                    if instr.strip().upper().startswith("IF"):
                        evaluated = _evaluate_if_field(instr, record_lookup)
                        if evaluated is not None:
                            for t in fld_simple.findall('.//'+qn('w:t')):
                                t.text = str(evaluated)
                            replaced += 1
                            continue
                    m = re.search(r'MERGEFIELD\s+"?([\w\-.]+)"?', instr, flags=re.IGNORECASE)
                    if not m:
                        continue
                    name = m.group(1).strip()
                    value = record_lookup.get(name.lower())
                    if value is not None:
                        for t in fld_simple.findall('.//'+qn('w:t')):
                            t.text = str(value)
                        replaced += 1

                # 3) Complex field codes (w:fldChar begin/separate/end with w:instrText)
                runs = list(para.runs)
                i_run = 0
                while i_run < len(runs):
                    run = runs[i_run]
                    # Look for instrText in this run
                    found_name = None
                    for child in run._r:
                        if child.tag == qn('w:instrText') and child.text:
                            if child.text.strip().upper().startswith("IF"):
                                evaluated = _evaluate_if_field(child.text, record_lookup)
                                if evaluated is not None:
                                    found_name = "__IF__"
                                    value = evaluated
                                    break
                            m = re.search(r'MERGEFIELD\s+"?([\w\-.]+)"?', child.text, flags=re.IGNORECASE)
                            if m:
                                found_name = m.group(1).strip()
                                value = record_lookup.get(found_name.lower())
                                break
                    if found_name:
                        if found_name != "__IF__" and value is None:
                            i_run += 1
                            continue
                        j = i_run + 1
                        while j < len(runs):
                            end_field = False
                            for child in runs[j]._r:
                                if child.tag == qn('w:fldChar') and child.get(qn('w:fldCharType')) == 'end':
                                    end_field = True
                            if runs[j].text:
                                runs[j].text = str(value)
                            j += 1
                            if end_field:
                                break
                        replaced += 1
                        i_run = j
                        continue
                    i_run += 1
                return replaced

            created = []
            for i, record in enumerate(data):
                new_doc = Document(self.current_file_path)

                # Paragraphs
                for para in new_doc.paragraphs:
                    _replace_mergefields_in_paragraph(para, record)

                # Tables
                for table in new_doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                _replace_mergefields_in_paragraph(para, record)

                output_path = output_pattern.format(index=i)
                err = self._ensure_path_allowed(output_path, for_write=True)
                if err:
                    return err
                new_doc.save(output_path)
                created.append(output_path)
            
            return f"Created {len(created)} documents: {', '.join(created)}"
        except Exception as e:
            logger.error(f"Failed to execute mail merge: {e}")
            return f"Error: {e}"

    # ==================== Bookmarks ====================

    def add_bookmark(self, paragraph_index: int, bookmark_name: str) -> str:
        """Add a bookmark to a paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            para = paragraphs[paragraph_index]
            
            # Create bookmark start and end
            bookmark_id = str(paragraph_index + 1)
            
            bookmark_start = OxmlElement('w:bookmarkStart')
            bookmark_start.set(qn('w:id'), bookmark_id)
            bookmark_start.set(qn('w:name'), bookmark_name)
            
            bookmark_end = OxmlElement('w:bookmarkEnd')
            bookmark_end.set(qn('w:id'), bookmark_id)
            
            para._p.insert(0, bookmark_start)
            para._p.append(bookmark_end)
            
            return f"Bookmark '{bookmark_name}' added at paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add bookmark: {e}")
            return f"Error: {e}"

    def add_hyperlink_to_bookmark(self, text: str, bookmark_name: str) -> str:
        """Add a hyperlink that jumps to a bookmark."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            para = self.current_document.add_paragraph()
            
            hyperlink = OxmlElement('w:hyperlink')
            hyperlink.set(qn('w:anchor'), bookmark_name)
            
            new_run = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            
            # Blue underlined style
            color = OxmlElement('w:color')
            color.set(qn('w:val'), '0000FF')
            rPr.append(color)
            
            u = OxmlElement('w:u')
            u.set(qn('w:val'), 'single')
            rPr.append(u)
            
            new_run.append(rPr)
            new_run.text = text
            
            t = OxmlElement('w:t')
            t.text = text
            new_run.append(t)
            
            hyperlink.append(new_run)
            para._p.append(hyperlink)
            
            return f"Internal link to bookmark '{bookmark_name}' added"
        except Exception as e:
            logger.error(f"Failed to add bookmark link: {e}")
            return f"Error: {e}"

    def list_bookmarks(self) -> str:
        """List all bookmarks in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            bookmarks = []
            body = self.current_document.element.body
            for elem in body.iter():
                if elem.tag == qn('w:bookmarkStart'):
                    name = elem.get(qn('w:name'))
                    if name and not name.startswith('_'):
                        bookmarks.append(name)
            
            if not bookmarks:
                return "No bookmarks found"
            return f"Bookmarks: {', '.join(bookmarks)}"
        except Exception as e:
            logger.error(f"Failed to list bookmarks: {e}")
            return f"Error: {e}"

    # ==================== Advanced Table Borders ====================

    def set_table_borders(self, table_index: int, 
                          top: Optional[str] = None, 
                          bottom: Optional[str] = None,
                          left: Optional[str] = None, 
                          right: Optional[str] = None,
                          inside_h: Optional[str] = None,
                          inside_v: Optional[str] = None,
                          size: int = 4, color: str = "000000") -> str:
        """
        Set table borders per side.
        Border style values: 'single', 'double', 'dotted', 'dashed', 'none'
        """
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            tbl = table._tbl
            
            # Get or create tblPr
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)
            
            # Get or create tblBorders
            tblBorders = tblPr.find(qn('w:tblBorders'))
            if tblBorders is None:
                tblBorders = OxmlElement('w:tblBorders')
                tblPr.append(tblBorders)
            
            def set_border(border_name: str, style: Optional[str]):
                if style is None:
                    return
                existing = tblBorders.find(qn(f'w:{border_name}'))
                if existing is not None:
                    tblBorders.remove(existing)
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), style if style != 'none' else 'nil')
                border.set(qn('w:sz'), str(size))
                border.set(qn('w:color'), color)
                tblBorders.append(border)
            
            set_border('top', top)
            set_border('bottom', bottom)
            set_border('left', left)
            set_border('right', right)
            set_border('insideH', inside_h)
            set_border('insideV', inside_v)
            
            return "Table borders updated"
        except Exception as e:
            logger.error(f"Failed to set table borders: {e}")
            return f"Error: {e}"

    def set_row_height(self, table_index: int, row_index: int, 
                        height_cm: float, rule: str = "exact") -> str:
        """
        Set row height.
        rule: 'exact', 'atLeast', 'auto'
        """
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Invalid row index"
            
            row = table.rows[row_index]
            tr = row._tr
            
            # Get or create trPr
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                tr.insert(0, trPr)
            
            # Remove existing height
            existing = trPr.find(qn('w:trHeight'))
            if existing is not None:
                trPr.remove(existing)
            
            trHeight = OxmlElement('w:trHeight')
            # Convert cm to twips (1 cm = 567 twips)
            twips = int(height_cm * 567)
            trHeight.set(qn('w:val'), str(twips))
            rule_map = {'exact': 'exact', 'atLeast': 'atLeast', 'auto': 'auto'}
            trHeight.set(qn('w:hRule'), rule_map.get(rule, 'exact'))
            trPr.append(trHeight)
            
            return f"Row {row_index} height set to {height_cm}cm"
        except Exception as e:
            logger.error(f"Failed to set row height: {e}")
            return f"Error: {e}"

    # ==================== Image Watermark ====================

    def add_image_watermark(self, image_path: str, width_inches: float = 4.0,
                             opacity: float = 0.3, section_index: int = 0) -> str:
        """Add an image as a watermark via header."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            if not os.path.exists(image_path):
                return f"Error: Image not found: {image_path}"
            
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            
            header = sections[section_index].header
            para = header.add_paragraph()
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.add_run()
            run.add_picture(image_path, width=Inches(width_inches))

            # Attempt to convert inline to anchored behind text for true watermark effect
            inline_elems = run._r.xpath('.//wp:inline')
            if inline_elems:
                inline = inline_elems[0]
                inline.tag = qn('wp:anchor')
                inline.set('behindDoc', '1')
                inline.set('layoutInCell', '1')
                inline.set('locked', '0')
                inline.set('allowOverlap', '1')
            return f"Image watermark added (anchor-based, {width_inches} inches wide)"
        except Exception as e:
            logger.error(f"Failed to add image watermark: {e}")
            return f"Error: {e}"

    # ==================== Document Analysis ====================

    def get_document_outline(self) -> str:
        """Get document outline based on headings."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            outline = []
            for para in self.current_document.paragraphs:
                style_name = para.style.name if para.style else ""
                if style_name.startswith("Heading"):
                    level = style_name.replace("Heading ", "")
                    indent = "  " * (int(level) - 1) if level.isdigit() else ""
                    outline.append(f"{indent}{para.text[:50]}")
            
            if not outline:
                return "No headings found"
            return "Document Outline:\n" + "\n".join(outline)
        except Exception as e:
            logger.error(f"Failed to get outline: {e}")
            return f"Error: {e}"

    def get_document_statistics(self) -> str:
        """Get comprehensive document statistics."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = len(self.current_document.paragraphs)
            tables = len(self.current_document.tables)
            sections = len(self.current_document.sections)
            
            word_count = 0
            char_count = 0
            for para in self.current_document.paragraphs:
                text = para.text
                word_count += len(text.split())
                char_count += len(text)
            
            # Count headings by level
            heading_counts = {}
            for para in self.current_document.paragraphs:
                style_name = para.style.name if para.style else ""
                if style_name.startswith("Heading"):
                    heading_counts[style_name] = heading_counts.get(style_name, 0) + 1
            
            stats = f"""Document Statistics:
- Paragraphs: {paragraphs}
- Tables: {tables}
- Sections: {sections}
- Words: {word_count}
- Characters: {char_count}
- Headings: {sum(heading_counts.values())}"""
            
            for h, count in sorted(heading_counts.items()):
                stats += f"\n  - {h}: {count}"
            
            return stats
        except Exception as e:
            logger.error(f"Failed to get statistics: {e}")
            return f"Error: {e}"

    def get_formatting_report(self) -> str:
        """Get a report of fonts and styles used in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            fonts_used = set()
            sizes_used = set()
            styles_used = set()
            
            for para in self.current_document.paragraphs:
                if para.style:
                    styles_used.add(para.style.name)
                for run in para.runs:
                    if run.font.name:
                        fonts_used.add(run.font.name)
                    if run.font.size:
                        sizes_used.add(str(run.font.size.pt) + "pt")
            
            report = f"""Formatting Report:
- Fonts used: {', '.join(sorted(fonts_used)) or 'Default only'}
- Font sizes: {', '.join(sorted(sizes_used)) or 'Default only'}
- Styles used: {', '.join(sorted(styles_used)) or 'None'}"""
            
            return report
        except Exception as e:
            logger.error(f"Failed to get formatting report: {e}")
            return f"Error: {e}"

    # ==================== Batch Operations ====================

    def batch_replace(self, replacements: Dict[str, str]) -> str:
        """Replace multiple strings in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            count = 0
            for para in self.current_document.paragraphs:
                for old, new in replacements.items():
                    if old in para.text:
                        for run in para.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, new)
                                count += 1
            
            # Also in tables
            for table in self.current_document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for old, new in replacements.items():
                                if old in para.text:
                                    for run in para.runs:
                                        if old in run.text:
                                            run.text = run.text.replace(old, new)
                                            count += 1
            
            return f"Made {count} replacements"
        except Exception as e:
            logger.error(f"Failed to batch replace: {e}")
            return f"Error: {e}"

    def batch_format_paragraphs(self, style_name: str, 
                                 start_index: int = 0, 
                                 end_index: Optional[int] = None) -> str:
        """Apply a style to a range of paragraphs."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            end = end_index if end_index is not None else len(paragraphs)
            
            count = 0
            for i in range(start_index, min(end, len(paragraphs))):
                paragraphs[i].style = style_name
                count += 1
            
            return f"Applied '{style_name}' to {count} paragraphs"
        except Exception as e:
            logger.error(f"Failed to batch format: {e}")
            return f"Error: {e}"

    # ==================== Section Operations ====================

    def add_section(self, section_type: str = "NEW_PAGE") -> str:
        """
        Add a new section to the document.
        section_type: 'NEW_PAGE', 'CONTINUOUS', 'ODD_PAGE', 'EVEN_PAGE'
        """
        if not self.current_document:
            return "Error: No document is open"
        try:
            from docx.enum.section import WD_SECTION
            
            section_map = {
                'NEW_PAGE': WD_SECTION.NEW_PAGE,
                'CONTINUOUS': WD_SECTION.CONTINUOUS,
                'ODD_PAGE': WD_SECTION.ODD_PAGE,
                'EVEN_PAGE': WD_SECTION.EVEN_PAGE,
            }
            
            sec_type = section_map.get(section_type.upper(), WD_SECTION.NEW_PAGE)
            
            # Add paragraph then set section break
            para = self.current_document.add_paragraph()
            new_section = self.current_document.add_section(sec_type)
            
            return f"Section added with type: {section_type}"
        except Exception as e:
            logger.error(f"Failed to add section: {e}")
            return f"Error: {e}"

    def list_sections(self) -> str:
        """List all sections with their properties."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            result = [f"Total sections: {len(sections)}"]
            
            for i, section in enumerate(sections):
                width = section.page_width.inches if section.page_width else "N/A"
                height = section.page_height.inches if section.page_height else "N/A"
                orient = "Landscape" if section.orientation == 1 else "Portrait"
                
                result.append(f"\nSection {i}:")
                result.append(f"  Orientation: {orient}")
                result.append(f"  Page size: {width:.2f}\" x {height:.2f}\"")
                result.append(f"  Margins: T={section.top_margin.inches:.2f}\", "
                            f"B={section.bottom_margin.inches:.2f}\", "
                            f"L={section.left_margin.inches:.2f}\", "
                            f"R={section.right_margin.inches:.2f}\"")
            
            return "\n".join(result)
        except Exception as e:
            logger.error(f"Failed to list sections: {e}")
            return f"Error: {e}"

    def set_section_properties(self, section_index: int = 0,
                                orientation: Optional[str] = None,
                                width_inches: Optional[float] = None,
                                height_inches: Optional[float] = None,
                                top_margin: Optional[float] = None,
                                bottom_margin: Optional[float] = None,
                                left_margin: Optional[float] = None,
                                right_margin: Optional[float] = None) -> str:
        """Set properties for a specific section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            from docx.enum.section import WD_ORIENT
            
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Invalid section index"
            
            section = sections[section_index]
            
            if orientation:
                if orientation.upper() == "LANDSCAPE":
                    section.orientation = WD_ORIENT.LANDSCAPE
                    # Swap dimensions
                    new_width = section.page_height
                    new_height = section.page_width
                    section.page_width = new_width
                    section.page_height = new_height
                elif orientation.upper() == "PORTRAIT":
                    section.orientation = WD_ORIENT.PORTRAIT
            
            if width_inches:
                section.page_width = Inches(width_inches)
            if height_inches:
                section.page_height = Inches(height_inches)
            if top_margin is not None:
                section.top_margin = Inches(top_margin)
            if bottom_margin is not None:
                section.bottom_margin = Inches(bottom_margin)
            if left_margin is not None:
                section.left_margin = Inches(left_margin)
            if right_margin is not None:
                section.right_margin = Inches(right_margin)
            
            return f"Section {section_index} properties updated"
        except Exception as e:
            logger.error(f"Failed to set section properties: {e}")
            return f"Error: {e}"

    # ==================== Zoned Headers/Footers ====================

    def add_zoned_header(self, section_index: int = 0,
                          left_text: str = "",
                          center_text: str = "",
                          right_text: str = "") -> str:
        """Add header with left, center, right zones."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Invalid section index"
            
            header = sections[section_index].header
            header.is_linked_to_previous = False
            
            # Clear existing
            for para in header.paragraphs:
                for run in para.runs:
                    run.text = ""
            
            # Create table for zones
            if header.paragraphs:
                para = header.paragraphs[0]
            else:
                para = header.add_paragraph()
            
            # Use tabs for positioning
            para.clear()
            
            # Add left text
            run = para.add_run(left_text)
            
            # Add center with tab
            para.add_run("\t")
            para.add_run(center_text)
            
            # Add right with tab
            para.add_run("\t")
            para.add_run(right_text)
            
            para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
            # Set tab stops
            from docx.shared import Twips
            pPr = para._p.get_or_add_pPr()
            tabs = OxmlElement('w:tabs')
            
            # Center tab
            tab_center = OxmlElement('w:tab')
            tab_center.set(qn('w:val'), 'center')
            tab_center.set(qn('w:pos'), '4680')  # ~3.25 inches
            tabs.append(tab_center)
            
            # Right tab
            tab_right = OxmlElement('w:tab')
            tab_right.set(qn('w:val'), 'right')
            tab_right.set(qn('w:pos'), '9360')  # ~6.5 inches
            tabs.append(tab_right)
            
            pPr.append(tabs)
            
            return "Zoned header added"
        except Exception as e:
            logger.error(f"Failed to add zoned header: {e}")
            return f"Error: {e}"

    def add_zoned_footer(self, section_index: int = 0,
                          left_text: str = "",
                          center_text: str = "",
                          right_text: str = "") -> str:
        """Add footer with left, center, right zones."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Invalid section index"
            
            footer = sections[section_index].footer
            footer.is_linked_to_previous = False
            
            # Clear and create new
            if footer.paragraphs:
                para = footer.paragraphs[0]
            else:
                para = footer.add_paragraph()
            
            para.clear()
            
            # Add zones with tabs
            para.add_run(left_text)
            para.add_run("\t")
            para.add_run(center_text)
            para.add_run("\t")
            para.add_run(right_text)
            
            # Set tab stops
            pPr = para._p.get_or_add_pPr()
            tabs = OxmlElement('w:tabs')
            
            tab_center = OxmlElement('w:tab')
            tab_center.set(qn('w:val'), 'center')
            tab_center.set(qn('w:pos'), '4680')
            tabs.append(tab_center)
            
            tab_right = OxmlElement('w:tab')
            tab_right.set(qn('w:val'), 'right')
            tab_right.set(qn('w:pos'), '9360')
            tabs.append(tab_right)
            
            pPr.append(tabs)
            
            return "Zoned footer added"
        except Exception as e:
            logger.error(f"Failed to add zoned footer: {e}")
            return f"Error: {e}"

    # ==================== Metadata Operations ====================

    def get_metadata(self) -> str:
        """Get document metadata/properties."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            props = self.current_document.core_properties
            
            result = ["Document Metadata:"]
            result.append(f"  Title: {props.title or 'Not set'}")
            result.append(f"  Author: {props.author or 'Not set'}")
            result.append(f"  Subject: {props.subject or 'Not set'}")
            result.append(f"  Keywords: {props.keywords or 'Not set'}")
            result.append(f"  Category: {props.category or 'Not set'}")
            result.append(f"  Comments: {props.comments or 'Not set'}")
            result.append(f"  Created: {props.created or 'Not set'}")
            result.append(f"  Modified: {props.modified or 'Not set'}")
            result.append(f"  Last modified by: {props.last_modified_by or 'Not set'}")
            result.append(f"  Revision: {props.revision or 'Not set'}")
            
            return "\n".join(result)
        except Exception as e:
            logger.error(f"Failed to get metadata: {e}")
            return f"Error: {e}"

    def set_metadata(self, title: Optional[str] = None,
                      author: Optional[str] = None,
                      subject: Optional[str] = None,
                      keywords: Optional[str] = None,
                      category: Optional[str] = None,
                      comments: Optional[str] = None) -> str:
        """Set document metadata/properties."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            props = self.current_document.core_properties
            
            if title is not None:
                props.title = title
            if author is not None:
                props.author = author
            if subject is not None:
                props.subject = subject
            if keywords is not None:
                props.keywords = keywords
            if category is not None:
                props.category = category
            if comments is not None:
                props.comments = comments
            
            return "Metadata updated"
        except Exception as e:
            logger.error(f"Failed to set metadata: {e}")
            return f"Error: {e}"

    def strip_personal_info(self) -> str:
        """Remove personal information from document metadata."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            props = self.current_document.core_properties
            
            props.author = ""
            props.last_modified_by = ""
            props.comments = ""
            
            return "Personal information stripped from metadata"
        except Exception as e:
            logger.error(f"Failed to strip personal info: {e}")
            return f"Error: {e}"

    # ==================== Security/Redaction ====================

    def redact_text(self, pattern: str, replacement: str = "[REDACTED]",
                     use_regex: bool = False) -> str:
        """Redact text matching pattern throughout the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            import re
            count = 0
            
            for para in self.current_document.paragraphs:
                for run in para.runs:
                    if use_regex:
                        if re.search(pattern, run.text):
                            run.text = re.sub(pattern, replacement, run.text)
                            count += 1
                    else:
                        if pattern in run.text:
                            run.text = run.text.replace(pattern, replacement)
                            count += 1
            
            # Also in tables
            for table in self.current_document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if use_regex:
                                    if re.search(pattern, run.text):
                                        run.text = re.sub(pattern, replacement, run.text)
                                        count += 1
                                else:
                                    if pattern in run.text:
                                        run.text = run.text.replace(pattern, replacement)
                                        count += 1
            
            return f"Redacted {count} occurrences"
        except Exception as e:
            logger.error(f"Failed to redact text: {e}")
            return f"Error: {e}"

    def sanitize_external_links(self) -> str:
        """Remove or neutralize external hyperlinks."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            body = self.current_document.element.body
            count = 0
            
            # Find all hyperlink elements
            for hyperlink in body.findall('.//'+qn('w:hyperlink')):
                # Check if external (has r:id)
                r_id = hyperlink.get(qn('r:id'))
                if r_id:
                    # Remove the hyperlink but keep the text
                    parent = hyperlink.getparent()
                    index = list(parent).index(hyperlink)
                    for child in hyperlink:
                        parent.insert(index, child)
                        index += 1
                    parent.remove(hyperlink)
                    count += 1
            
            return f"Removed {count} external hyperlinks"
        except Exception as e:
            logger.error(f"Failed to sanitize links: {e}")
            return f"Error: {e}"

    # ==================== Paragraph Text Operations ====================

    def get_paragraph_text(self, index: int) -> str:
        """Get text from a specific paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index >= len(paragraphs):
                return f"Error: Invalid paragraph index (max: {len(paragraphs) - 1})"
            
            return paragraphs[index].text
        except Exception as e:
            logger.error(f"Failed to get paragraph text: {e}")
            return f"Error: {e}"

    def set_paragraph_text(self, index: int, text: str) -> str:
        """Replace text of a specific paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            para = paragraphs[index]
            # Clear existing runs
            for run in para.runs:
                run.text = ""
            
            # Add new text
            if para.runs:
                para.runs[0].text = text
            else:
                para.add_run(text)
            
            return f"Paragraph {index} text updated"
        except Exception as e:
            logger.error(f"Failed to set paragraph text: {e}")
            return f"Error: {e}"

    def delete_paragraph(self, index: int) -> str:
        """Delete a paragraph by index."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            para = paragraphs[index]
            p = para._element
            p.getparent().remove(p)
            
            return f"Paragraph {index} deleted"
        except Exception as e:
            logger.error(f"Failed to delete paragraph: {e}")
            return f"Error: {e}"

    def insert_paragraph_after(self, index: int, text: str,
                                style: Optional[str] = None) -> str:
        """Insert a paragraph after the specified index."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            
            # Create new paragraph element
            new_para = OxmlElement('w:p')
            new_run = OxmlElement('w:r')
            new_text = OxmlElement('w:t')
            new_text.text = text
            new_run.append(new_text)
            new_para.append(new_run)
            
            # Insert after target
            target = paragraphs[index]._element
            target.addnext(new_para)
            
            return f"Paragraph inserted after index {index}"
        except Exception as e:
            logger.error(f"Failed to insert paragraph: {e}")
            return f"Error: {e}"

    def find_text(self, search_text: str, case_sensitive: bool = False) -> str:
        """Find all occurrences of text in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            results = []
            
            for i, para in enumerate(self.current_document.paragraphs):
                text = para.text
                if not case_sensitive:
                    if search_text.lower() in text.lower():
                        results.append(f"Paragraph {i}: ...{text[:100]}...")
                else:
                    if search_text in text:
                        results.append(f"Paragraph {i}: ...{text[:100]}...")
            
            if not results:
                return f"No occurrences of '{search_text}' found"
            
            return f"Found {len(results)} occurrences:\n" + "\n".join(results)
        except Exception as e:
            logger.error(f"Failed to find text: {e}")
            return f"Error: {e}"

    # ==================== Document Structure ====================

    def get_document_structure(self) -> str:
        """Get structural overview of the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            result = ["Document Structure:"]
            result.append(f"\nSections: {len(self.current_document.sections)}")
            result.append(f"Paragraphs: {len(self.current_document.paragraphs)}")
            result.append(f"Tables: {len(self.current_document.tables)}")
            
            # List headings
            result.append("\nHeadings:")
            for i, para in enumerate(self.current_document.paragraphs):
                style = para.style.name if para.style else ""
                if style.startswith("Heading"):
                    level = style.replace("Heading ", "")
                    indent = "  " * (int(level) if level.isdigit() else 1)
                    result.append(f"{indent}[{i}] {para.text[:60]}")
            
            # List tables
            if self.current_document.tables:
                result.append("\nTables:")
                for i, table in enumerate(self.current_document.tables):
                    result.append(f"  Table {i}: {len(table.rows)} rows x {len(table.columns)} cols")
            
            return "\n".join(result)
        except Exception as e:
            logger.error(f"Failed to get structure: {e}")
            return f"Error: {e}"

    def get_raw_xml(self) -> str:
        """Get the raw XML structure of the document body."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            from lxml import etree
            body = self.current_document.element.body
            xml_str = etree.tostring(body, pretty_print=True, encoding='unicode')
            
            # Truncate if too long
            if len(xml_str) > 10000:
                return xml_str[:10000] + "\n... (truncated)"
            return xml_str
        except Exception as e:
            logger.error(f"Failed to get raw XML: {e}")
            return f"Error: {e}"

    # ==================== Advanced Table Operations ====================

    def merge_cells(self, table_index: int, 
                     start_row: int, start_col: int,
                     end_row: int, end_col: int) -> str:
        """Merge cells in a rectangular area."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            
            # Get cell references
            start_cell = table.cell(start_row, start_col)
            end_cell = table.cell(end_row, end_col)
            
            # Merge
            start_cell.merge(end_cell)
            
            return f"Merged cells from ({start_row},{start_col}) to ({end_row},{end_col})"
        except Exception as e:
            logger.error(f"Failed to merge cells: {e}")
            return f"Error: {e}"

    def get_table_data(self, table_index: int) -> str:
        """Get all data from a table as text."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            result = []
            
            for i, row in enumerate(table.rows):
                row_data = [cell.text for cell in row.cells]
                result.append(f"Row {i}: {row_data}")
            
            return "\n".join(result)
        except Exception as e:
            logger.error(f"Failed to get table data: {e}")
            return f"Error: {e}"

    def list_tables(self) -> str:
        """List all tables with their dimensions."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if not tables:
                return "No tables in document"
            
            result = [f"Total tables: {len(tables)}"]
            for i, table in enumerate(tables):
                rows = len(table.rows)
                cols = len(table.columns)
                style = table.style.name if table.style else "No style"
                result.append(f"  Table {i}: {rows} rows x {cols} cols, style: {style}")
            
            return "\n".join(result)
        except Exception as e:
            logger.error(f"Failed to list tables: {e}")
            return f"Error: {e}"

    def set_table_cell_text(self, table_index: int, row: int, col: int, 
                             text: str) -> str:
        """Set text in a specific table cell."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            if row >= len(table.rows):
                return f"Error: Invalid row index"
            if col >= len(table.columns):
                return f"Error: Invalid column index"
            
            table.cell(row, col).text = text
            return f"Cell ({row},{col}) text set"
        except Exception as e:
            logger.error(f"Failed to set cell text: {e}")
            return f"Error: {e}"

    def add_table_row(self, table_index: int, data: Optional[List[str]] = None) -> str:
        """Add a row to a table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            row = table.add_row()
            
            if data:
                for i, text in enumerate(data):
                    if i < len(row.cells):
                        row.cells[i].text = text
            
            return f"Row added to table {table_index}"
        except Exception as e:
            logger.error(f"Failed to add table row: {e}")
            return f"Error: {e}"

    def delete_table_row(self, table_index: int, row_index: int) -> str:
        """Delete a row from a table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Invalid row index"
            
            row = table.rows[row_index]
            tr = row._tr
            tr.getparent().remove(tr)
            
            return f"Row {row_index} deleted from table {table_index}"
        except Exception as e:
            logger.error(f"Failed to delete table row: {e}")
            return f"Error: {e}"

    # ==================== Document Copy ====================

    def copy_document(self, dest_path: str) -> str:
        """Create a copy of the current document."""
        if not self.current_document:
            return "Error: No document is open"
        if not self.current_file_path:
            return "Error: No file path set"
        try:
            err = self._ensure_path_allowed(dest_path, for_write=True)
            if err:
                return err
            import shutil
            shutil.copy2(self.current_file_path, dest_path)
            return f"Document copied to: {dest_path}"
        except Exception as e:
            logger.error(f"Failed to copy document: {e}")
            return f"Error: {e}"

    # ==================== List Files ====================

    def list_docx_files(self, directory: str) -> str:
        """List all .docx files in a directory."""
        try:
            err = self._ensure_path_allowed(directory, for_write=False)
            if err:
                return err
            if not os.path.isdir(directory):
                return f"Error: Directory not found: {directory}"
            
            files = []
            for f in os.listdir(directory):
                if f.endswith('.docx') and not f.startswith('~$'):
                    full_path = os.path.join(directory, f)
                    size = os.path.getsize(full_path)
                    files.append(f"{f} ({size} bytes)")
            
            if not files:
                return f"No .docx files found in {directory}"
            
            return f"Found {len(files)} files:\n" + "\n".join(files)
        except Exception as e:
            logger.error(f"Failed to list files: {e}")
            return f"Error: {e}"

    # ==================== Image Operations ====================

    def add_image(self, image_path: str, width_inches: Optional[float] = None,
                   height_inches: Optional[float] = None) -> str:
        """Add an image to the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            if not os.path.exists(image_path):
                return f"Error: Image not found: {image_path}"
            
            para = self.current_document.add_paragraph()
            run = para.add_run()
            
            if width_inches and height_inches:
                run.add_picture(image_path, 
                               width=Inches(width_inches), 
                               height=Inches(height_inches))
            elif width_inches:
                run.add_picture(image_path, width=Inches(width_inches))
            elif height_inches:
                run.add_picture(image_path, height=Inches(height_inches))
            else:
                run.add_picture(image_path)
            
            return f"Image added: {image_path}"
        except Exception as e:
            logger.error(f"Failed to add image: {e}")
            return f"Error: {e}"

    def add_image_base64(self, base64_data: str, 
                          width_inches: Optional[float] = None) -> str:
        """Add an image from base64 data."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            import base64
            import io
            
            # Remove data URI prefix if present
            if ',' in base64_data:
                base64_data = base64_data.split(',')[1]
            
            image_bytes = base64.b64decode(base64_data)
            image_stream = io.BytesIO(image_bytes)
            
            para = self.current_document.add_paragraph()
            run = para.add_run()
            
            if width_inches:
                run.add_picture(image_stream, width=Inches(width_inches))
            else:
                run.add_picture(image_stream)
            
            return "Image added from base64"
        except Exception as e:
            logger.error(f"Failed to add base64 image: {e}")
            return f"Error: {e}"

    def list_images(self) -> str:
        """List all images in the document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            # Access inline shapes
            images = []
            for i, shape in enumerate(self.current_document.inline_shapes):
                width = shape.width.inches if shape.width else "N/A"
                height = shape.height.inches if shape.height else "N/A"
                images.append(f"Image {i}: {width:.2f}\" x {height:.2f}\"")
            
            if not images:
                return "No images found in document"
            
            return f"Found {len(images)} images:\n" + "\n".join(images)
        except Exception as e:
            logger.error(f"Failed to list images: {e}")
            return f"Error: {e}"

    # ==================== Word Count ====================

    def get_word_count(self) -> str:
        """Get detailed word count statistics."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            word_count = 0
            char_count = 0
            char_no_space = 0
            para_count = len(self.current_document.paragraphs)
            line_count = 0
            
            for para in self.current_document.paragraphs:
                text = para.text
                words = text.split()
                word_count += len(words)
                char_count += len(text)
                char_no_space += len(text.replace(" ", ""))
                line_count += text.count('\n') + 1 if text else 0
            
            return f"""Word Count Statistics:
- Words: {word_count}
- Characters (with spaces): {char_count}
- Characters (no spaces): {char_no_space}
- Paragraphs: {para_count}
- Lines (approx): {line_count}"""
        except Exception as e:
            logger.error(f"Failed to get word count: {e}")
            return f"Error: {e}"

    # ==================== Page Layout Extensions ====================

    def set_page_columns(self, section_index: int = 0, num_columns: int = 2,
                          space_cm: float = 0.5, separator_line: bool = False) -> str:
        """Configure multi-column layout for a section."""
        if not self.current_document:
            return "Error: No document is open"
        if num_columns < 1:
            return "Error: num_columns must be >= 1"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return "Error: Section index out of range"
            sectPr = sections[section_index]._sectPr
            cols = sectPr.find(qn('w:cols'))
            if cols is None:
                cols = OxmlElement('w:cols')
                sectPr.append(cols)
            cols.set(qn('w:num'), str(num_columns))
            cols.set(qn('w:space'), str(int(space_cm * 567)))
            if separator_line:
                cols.set(qn('w:sep'), '1')
            elif cols.get(qn('w:sep')):
                cols.attrib.pop(qn('w:sep'))
            return f"Columns set to {num_columns} with {space_cm}cm spacing"
        except Exception as e:
            logger.error(f"Failed to set columns: {e}")
            return f"Error: {e}"

    # ==================== Editable Ranges ====================

    def add_editable_range(self, start_paragraph: int, end_paragraph: int, editor: str = "everyone") -> str:
        """Mark a range of paragraphs as editable using permStart/permEnd."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if start_paragraph < 0 or end_paragraph >= len(paragraphs) or start_paragraph > end_paragraph:
                return "Error: Invalid paragraph range"
            body = self.current_document.element.body
            start_p = paragraphs[start_paragraph]._p
            end_p = paragraphs[end_paragraph]._p
            body_children = list(body)
            start_idx = body_children.index(start_p)
            end_idx = body_children.index(end_p)
            range_id = str(uuid.uuid4().int % 65535)
            perm_start = OxmlElement('w:permStart')
            perm_start.set(qn('w:id'), range_id)
            perm_start.set(qn('w:ed'), editor)
            perm_end = OxmlElement('w:permEnd')
            perm_end.set(qn('w:id'), range_id)
            body.insert(start_idx, perm_start)
            body.insert(end_idx + 2, perm_end)
            return f"Editable range added for paragraphs {start_paragraph}-{end_paragraph}"
        except Exception as e:
            logger.error(f"Failed to add editable range: {e}")
            return f"Error: {e}"

    # ==================== Form Controls ====================

    def add_checkbox_form(self, paragraph_index: int, label: str = "Checkbox", checked: bool = False) -> str:
        """Insert a checkbox content control (sdt) at the paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return "Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            sdt = OxmlElement('w:sdt')
            sdtPr = OxmlElement('w:sdtPr')
            alias = OxmlElement('w:alias')
            alias.set(qn('w:val'), label)
            sdtPr.append(alias)
            tag = OxmlElement('w:tag')
            tag.set(qn('w:val'), f"checkbox:{label}")
            sdtPr.append(tag)
            checkbox = OxmlElement(qn('w14:checkbox'))
            checked_elem = OxmlElement(qn('w14:checked'))
            checked_elem.set(qn('w14:val'), '1' if checked else '0')
            checkbox.append(checked_elem)
            sdtPr.append(checkbox)
            sdtContent = OxmlElement('w:sdtContent')
            run = OxmlElement('w:r')
            text = OxmlElement('w:t')
            text.text = label
            run.append(text)
            sdtContent.append(run)
            sdt.append(sdtPr)
            sdt.append(sdtContent)
            para._p.append(sdt)
            return f"Checkbox control added at paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add checkbox: {e}")
            return f"Error: {e}"

    def add_dropdown_form(self, paragraph_index: int, options: List[str], placeholder: str = "Select") -> str:
        """Insert a dropdown (comboBox) content control."""
        if not self.current_document:
            return "Error: No document is open"
        if not options:
            return "Error: Options list cannot be empty"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return "Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            sdt = OxmlElement('w:sdt')
            sdtPr = OxmlElement('w:sdtPr')
            alias = OxmlElement('w:alias')
            alias.set(qn('w:val'), "Dropdown")
            sdtPr.append(alias)
            dd = OxmlElement('w:dropDownList')
            for opt in options:
                item = OxmlElement('w:listItem')
                item.set(qn('w:displayText'), opt)
                item.set(qn('w:value'), opt)
                dd.append(item)
            sdtPr.append(dd)
            sdtContent = OxmlElement('w:sdtContent')
            run = OxmlElement('w:r')
            text = OxmlElement('w:t')
            text.text = placeholder
            run.append(text)
            sdtContent.append(run)
            sdt.append(sdtPr)
            sdt.append(sdtContent)
            para._p.append(sdt)
            return f"Dropdown control added at paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add dropdown: {e}")
            return f"Error: {e}"

    def add_date_picker_form(self, paragraph_index: int, date_value: Optional[str] = None, format_string: str = "yyyy-MM-dd") -> str:
        """Insert a date picker content control."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return "Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            sdt = OxmlElement('w:sdt')
            sdtPr = OxmlElement('w:sdtPr')
            alias = OxmlElement('w:alias')
            alias.set(qn('w:val'), "DatePicker")
            sdtPr.append(alias)
            date_el = OxmlElement('w:date')
            date_el.set(qn('w:fullDate'), date_value or datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'))
            date_el.set(qn('w:format'), format_string)
            sdtPr.append(date_el)
            sdtContent = OxmlElement('w:sdtContent')
            run = OxmlElement('w:r')
            text = OxmlElement('w:t')
            text.text = date_value or datetime.utcnow().strftime('%Y-%m-%d')
            run.append(text)
            sdtContent.append(run)
            sdt.append(sdtPr)
            sdt.append(sdtContent)
            para._p.append(sdt)
            return f"Date picker added at paragraph {paragraph_index}"
        except Exception as e:
            logger.error(f"Failed to add date picker: {e}")
            return f"Error: {e}"

    # ==================== Linting / Spell-check (lightweight) ====================

    def lint_document(self) -> str:
        """Run lightweight lint: double spaces, repeated words, trailing spaces, optional grammar."""
        if not self.current_document:
            return "Error: No document is open"
        import re
        issues = []
        for i, para in enumerate(self.current_document.paragraphs):
            text = para.text
            if '  ' in text:
                issues.append(f"Paragraph {i}: contains double spaces")
            if re.search(r"\b(\w+)\s+\1\b", text, flags=re.IGNORECASE):
                issues.append(f"Paragraph {i}: repeated word detected")
            if text.rstrip() != text:
                issues.append(f"Paragraph {i}: trailing whitespace")
        # Optional grammar/spell if language_tool_python is available
        try:
            import language_tool_python
            sample = "\n".join(p.text for p in self.current_document.paragraphs)
            tool = language_tool_python.LanguageTool('en-US')
            matches = tool.check(sample[:4000])
            if matches:
                issues.append(f"Grammar/Spell suggestions: {len(matches)} issues (showing first 5)")
                for m in matches[:5]:
                    issues.append(f" - {m.context}:{m.message}")
        except ImportError:
            issues.append("Grammar check skipped (language_tool_python not installed)")
        if not issues:
            return "Document lint clean"
        return "Lint report:\n" + "\n".join(issues)

    # ==================== Collaboration (lightweight) ====================

    def start_collaboration_session(self, name: Optional[str] = None) -> str:
        session_id = uuid.uuid4().hex[:8]
        self.collaboration_sessions[session_id] = {
            "name": name or session_id,
            "log": [],
            "version": 0,
        }
        return f"Session started: {session_id}"

    def apply_collaboration_change(self, session_id: str, paragraph_index: int, new_text: str) -> str:
        if session_id not in self.collaboration_sessions:
            return "Error: Session not found"
        if not self.current_document:
            return "Error: No document is open"
        paragraphs = self.current_document.paragraphs
        if paragraph_index >= len(paragraphs):
            return "Error: Invalid paragraph index"
        paragraphs[paragraph_index].text = new_text
        session = self.collaboration_sessions[session_id]
        session["version"] += 1
        session["log"].append({
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "paragraph": paragraph_index,
            "text": new_text,
            "version": session["version"],
        })
        return f"Change applied in session {session_id}, version {session['version']}"

    def get_collaboration_log(self, session_id: str) -> str:
        if session_id not in self.collaboration_sessions:
            return "Error: Session not found"
        log = self.collaboration_sessions[session_id]["log"]
        if not log:
            return "No changes recorded"
        lines = [f"v{entry['version']} @ {entry['timestamp']} -> paragraph {entry['paragraph']}" for entry in log]
        return "\n".join(lines)
