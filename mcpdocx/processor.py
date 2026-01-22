"""
Core Document Processor for DOCX operations.
Handles all document manipulation logic.
"""

import os
import re
import logging
import hashlib
from typing import Optional, Dict, Any, List

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Twips
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logger = logging.getLogger(__name__)


def parse_color(color_str: str) -> Optional[tuple]:
    """Parse color string to RGB tuple."""
    if not color_str:
        return None
    color_str = color_str.strip().upper()
    
    # Named colors
    color_map = {
        "RED": (255, 0, 0),
        "GREEN": (0, 255, 0),
        "BLUE": (0, 0, 255),
        "BLACK": (0, 0, 0),
        "WHITE": (255, 255, 255),
        "YELLOW": (255, 255, 0),
        "ORANGE": (255, 165, 0),
        "PURPLE": (128, 0, 128),
        "GRAY": (128, 128, 128),
        "GREY": (128, 128, 128),
    }
    
    if color_str in color_map:
        return color_map[color_str]
    
    # Hex color
    if color_str.startswith("#"):
        color_str = color_str[1:]
    
    if len(color_str) == 6:
        try:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            return (r, g, b)
        except ValueError:
            pass
    
    return None


class DocumentProcessor:
    """
    Unified document processor for Word documents.
    Manages document lifecycle and provides all editing operations.
    """
    
    def __init__(self):
        self.documents: Dict[str, Document] = {}
        self.current_document: Optional[Document] = None
        self.current_file_path: Optional[str] = None
        self.footnotes: List[str] = []
        self.comments: List[Dict[str, Any]] = []
        self.collaboration_sessions: Dict[str, Dict] = {}
    
    # ==================== Document Lifecycle ====================
    
    def create_document(self, file_path: str, title: Optional[str] = None) -> str:
        """Create a new Word document."""
        try:
            doc = Document()
            if title:
                doc.add_heading(title, 0)
            doc.save(file_path)
            self.current_document = doc
            self.current_file_path = file_path
            self.documents[file_path] = doc
            self.footnotes = []
            self.comments = []
            return f"Document created: {file_path}"
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            return f"Error: {e}"
    
    def open_document(self, file_path: str) -> str:
        """Open an existing Word document."""
        try:
            if not os.path.exists(file_path):
                return f"Error: File not found: {file_path}"
            doc = Document(file_path)
            self.current_document = doc
            self.current_file_path = file_path
            self.documents[file_path] = doc
            self.footnotes = []
            self.comments = []
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
        self.footnotes = []
        self.comments = []
        return "Document closed"
    
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
        word_count = sum(len(p.text.split()) for p in doc.paragraphs)
        info.append(f"Word count: {word_count}")
        return "\n".join(info)
    
    def get_document_text(self) -> str:
        """Get all text from the document."""
        if not self.current_document:
            return "Error: No document is open"
        paragraphs = [p.text for p in self.current_document.paragraphs]
        return "\n\n".join(paragraphs) if paragraphs else "Document is empty"
    
    def list_open_documents(self) -> str:
        """List all open documents."""
        if not self.documents:
            return "No documents open"
        docs = list(self.documents.keys())
        lines = []
        for doc in docs:
            marker = " (current)" if doc == self.current_file_path else ""
            lines.append(f"- {doc}{marker}")
        return "Open documents:\n" + "\n".join(lines)
    
    def switch_document(self, file_path: str) -> str:
        """Switch to a different open document."""
        if file_path not in self.documents:
            return f"Error: Document not open: {file_path}"
        self.current_document = self.documents[file_path]
        self.current_file_path = file_path
        return f"Switched to: {file_path}"
    
    # ==================== Content Operations ====================
    
    def _set_paragraph_alignment(self, paragraph, alignment: str) -> None:
        """Set paragraph alignment."""
        align_map = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        }
        if alignment and alignment.lower() in align_map:
            paragraph.alignment = align_map[alignment.lower()]
    
    def _clear_paragraph(self, paragraph) -> None:
        """Clear paragraph content."""
        p = paragraph._p
        for child in list(p):
            p.remove(child)
    
    def add_paragraph(self, text: str, style: Optional[str] = None, bold: bool = False,
                      italic: bool = False, underline: bool = False,
                      font_size: Optional[int] = None, font_name: Optional[str] = None,
                      color: Optional[str] = None, alignment: Optional[str] = None) -> str:
        """Add a paragraph with formatting."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            if style:
                try:
                    paragraph = self.current_document.add_paragraph(text, style=style)
                except KeyError:
                    paragraph = self.current_document.add_paragraph(text)
            else:
                paragraph = self.current_document.add_paragraph(text)
            
            if paragraph.runs:
                run = paragraph.runs[0]
                run.bold = bold
                run.italic = italic
                run.underline = underline
                if font_size:
                    run.font.size = Pt(font_size)
                if font_name:
                    run.font.name = font_name
                if color:
                    rgb = parse_color(color)
                    if rgb:
                        run.font.color.rgb = RGBColor(*rgb)
            
            if alignment:
                self._set_paragraph_alignment(paragraph, alignment)
            
            return "Paragraph added"
        except Exception as e:
            logger.error(f"Failed to add paragraph: {e}")
            return f"Error: {e}"
    
    def add_heading(self, text: str, level: int = 1) -> str:
        """Add a heading."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            self.current_document.add_heading(text, level=level)
            return f"Heading level {level} added"
        except Exception as e:
            logger.error(f"Failed to add heading: {e}")
            return f"Error: {e}"
    
    def add_page_break(self) -> str:
        """Add a page break."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            self.current_document.add_page_break()
            return "Page break added"
        except Exception as e:
            return f"Error: {e}"
    
    def get_paragraph_text(self, index: int) -> str:
        """Get paragraph text by index."""
        if not self.current_document:
            return "Error: No document is open"
        paragraphs = self.current_document.paragraphs
        if index < 0 or index >= len(paragraphs):
            return f"Error: Index out of range (0-{len(paragraphs)-1})"
        return paragraphs[index].text
    
    def set_paragraph_text(self, index: int, text: str) -> str:
        """Set paragraph text by index."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index < 0 or index >= len(paragraphs):
                return f"Error: Index out of range"
            para = paragraphs[index]
            for run in para.runs:
                run.text = ""
            if para.runs:
                para.runs[0].text = text
            else:
                para.add_run(text)
            return f"Paragraph {index} updated"
        except Exception as e:
            return f"Error: {e}"
    
    def delete_paragraph(self, index: int) -> str:
        """Delete paragraph by index."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index < 0 or index >= len(paragraphs):
                return f"Error: Index out of range"
            p = paragraphs[index]._element
            p.getparent().remove(p)
            return f"Paragraph {index} deleted"
        except Exception as e:
            return f"Error: {e}"
    
    def insert_paragraph_after(self, index: int, text: str, style: Optional[str] = None) -> str:
        """Insert paragraph after index."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if index < 0 or index >= len(paragraphs):
                return f"Error: Index out of range"
            new_para = OxmlElement('w:p')
            new_run = OxmlElement('w:r')
            new_text = OxmlElement('w:t')
            new_text.text = text
            new_run.append(new_text)
            new_para.append(new_run)
            paragraphs[index]._element.addnext(new_para)
            return f"Paragraph inserted after index {index}"
        except Exception as e:
            return f"Error: {e}"
    
    def count_paragraphs(self) -> str:
        """Count paragraphs."""
        if not self.current_document:
            return "Error: No document is open"
        return f"Document has {len(self.current_document.paragraphs)} paragraphs"
    
    # ==================== Search & Replace ====================
    
    def search_text(self, keyword: str) -> str:
        """Search for text in the document."""
        if not self.current_document:
            return "Error: No document is open"
        results = []
        for i, para in enumerate(self.current_document.paragraphs):
            if keyword.lower() in para.text.lower():
                results.append(f"Paragraph {i}: {para.text[:80]}...")
        for t_idx, table in enumerate(self.current_document.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    if keyword.lower() in cell.text.lower():
                        results.append(f"Table {t_idx} cell ({r_idx},{c_idx}): {cell.text[:50]}...")
        if not results:
            return f"'{keyword}' not found"
        return f"Found {len(results)} matches:\n" + "\n".join(results)
    
    def find_and_replace(self, find_text: str, replace_text: str) -> str:
        """Find and replace text."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            count = 0
            for para in self.current_document.paragraphs:
                if find_text in para.text:
                    for run in para.runs:
                        if find_text in run.text:
                            run.text = run.text.replace(find_text, replace_text)
                            count += 1
            for table in self.current_document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if find_text in para.text:
                                for run in para.runs:
                                    if find_text in run.text:
                                        run.text = run.text.replace(find_text, replace_text)
                                        count += 1
            return f"Replaced {count} occurrences"
        except Exception as e:
            return f"Error: {e}"
    
    def batch_replace(self, replacements: Dict[str, str]) -> str:
        """Batch replace multiple strings."""
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
            return f"Error: {e}"
    
    def redact_text(self, pattern: str, replacement: str = "[REDACTED]", use_regex: bool = False) -> str:
        """Redact text matching pattern."""
        if not self.current_document:
            return "Error: No document is open"
        try:
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
            return f"Error: {e}"
    
    # ==================== Table Operations ====================
    
    def add_table(self, rows: int, cols: int, data: Optional[List[List[str]]] = None,
                  style: str = "Table Grid") -> str:
        """Add a table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            table = self.current_document.add_table(rows=rows, cols=cols, style=style)
            if data:
                for i, row_data in enumerate(data):
                    if i < rows:
                        row = table.rows[i]
                        for j, cell_text in enumerate(row_data):
                            if j < cols:
                                row.cells[j].text = str(cell_text)
            return f"Table {rows}x{cols} added"
        except Exception as e:
            return f"Error: {e}"
    
    def get_tables_info(self) -> str:
        """Get info about all tables."""
        if not self.current_document:
            return "Error: No document is open"
        tables = self.current_document.tables
        if not tables:
            return "No tables in document"
        info = []
        for i, table in enumerate(tables):
            style = table.style.name if table.style else "No style"
            info.append(f"Table {i}: {len(table.rows)} rows x {len(table.columns)} cols, style: {style}")
        return "\n".join(info)
    
    def get_table_data(self, table_index: int) -> str:
        """Get table content."""
        if not self.current_document:
            return "Error: No document is open"
        tables = self.current_document.tables
        if table_index < 0 or table_index >= len(tables):
            return f"Error: Invalid table index"
        table = tables[table_index]
        result = []
        for i, row in enumerate(table.rows):
            row_data = [cell.text for cell in row.cells]
            result.append(" | ".join(row_data))
        return "\n".join(result)
    
    def edit_table_cell(self, table_index: int, row_index: int, col_index: int, text: str) -> str:
        """Edit table cell."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Row index out of range"
            if col_index >= len(table.columns):
                return f"Error: Column index out of range"
            table.cell(row_index, col_index).text = text
            return f"Cell ({row_index},{col_index}) updated"
        except Exception as e:
            return f"Error: {e}"
    
    def add_table_row(self, table_index: int, data: Optional[List[str]] = None) -> str:
        """Add row to table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            new_row = table.add_row()
            if data:
                for i, cell_text in enumerate(data):
                    if i < len(new_row.cells):
                        new_row.cells[i].text = str(cell_text)
            return "Row added"
        except Exception as e:
            return f"Error: {e}"
    
    def delete_table_row(self, table_index: int, row_index: int) -> str:
        """Delete row from table."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            if row_index >= len(table.rows):
                return f"Error: Row index out of range"
            row = table.rows[row_index]._tr
            row.getparent().remove(row)
            return f"Row {row_index} deleted"
        except Exception as e:
            return f"Error: {e}"
    
    def merge_table_cells(self, table_index: int, start_row: int, start_col: int,
                          end_row: int, end_col: int) -> str:
        """Merge table cells."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            start_cell = table.cell(start_row, start_col)
            end_cell = table.cell(end_row, end_col)
            start_cell.merge(end_cell)
            return f"Cells merged from ({start_row},{start_col}) to ({end_row},{end_col})"
        except Exception as e:
            return f"Error: {e}"
    
    def set_table_cell_shading(self, table_index: int, row_index: int, 
                                col_index: int, fill_color: str = "FFFF00") -> str:
        """Set cell background color."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            cell = tables[table_index].cell(row_index, col_index)
            tc_pr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), fill_color.replace('#', ''))
            tc_pr.append(shd)
            return "Cell shading applied"
        except Exception as e:
            return f"Error: {e}"
    
    def apply_table_alternating_rows(self, table_index: int, 
                                      color1: str = "FFFFFF", color2: str = "F2F2F2") -> str:
        """Apply alternating row colors."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
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
            return f"Error: {e}"
    
    def highlight_table_header(self, table_index: int, header_color: str = "4472C4",
                                text_color: str = "FFFFFF") -> str:
        """Highlight table header row."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            if not table.rows:
                return "Error: Table has no rows"
            for cell in table.rows[0].cells:
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
            return f"Error: {e}"
    
    def set_column_width(self, table_index: int, col_index: int, width_cm: float) -> str:
        """Set column width."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            tables = self.current_document.tables
            if table_index >= len(tables):
                return f"Error: Invalid table index"
            table = tables[table_index]
            for row in table.rows:
                if col_index < len(row.cells):
                    row.cells[col_index].width = Cm(width_cm)
            return f"Column {col_index} width set to {width_cm}cm"
        except Exception as e:
            return f"Error: {e}"
    
    def set_row_height(self, table_index: int, row_index: int, 
                       height_cm: float, rule: str = "exact") -> str:
        """Set row height."""
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
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                tr.insert(0, trPr)
            existing = trPr.find(qn('w:trHeight'))
            if existing is not None:
                trPr.remove(existing)
            trHeight = OxmlElement('w:trHeight')
            twips = int(height_cm * 567)
            trHeight.set(qn('w:val'), str(twips))
            trHeight.set(qn('w:hRule'), rule)
            trPr.append(trHeight)
            return f"Row {row_index} height set to {height_cm}cm"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Header & Footer ====================
    
    def add_header(self, text: str, section_index: int = 0, alignment: str = "center") -> str:
        """Add header text."""
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
            return f"Error: {e}"
    
    def add_footer(self, text: str, section_index: int = 0, alignment: str = "center") -> str:
        """Add footer text."""
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
            return f"Error: {e}"
    
    def _add_page_field(self, paragraph) -> None:
        """Insert PAGE field."""
        for tag, typ in [("w:fldChar", "begin"), ("w:instrText", None), 
                        ("w:fldChar", "separate"), ("w:fldChar", "end")]:
            el = OxmlElement(tag)
            if typ:
                el.set(qn("w:fldCharType"), typ)
            elif tag == "w:instrText":
                el.set(qn("xml:space"), "preserve")
                el.text = "PAGE"
            paragraph._p.append(el)
    
    def add_page_numbers(self, position: str = "footer", alignment: str = "center",
                         format_string: str = "Page {page}", section_index: int = 0) -> str:
        """Add page numbers."""
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
            prefix = format_string.replace("{page}", "").strip()
            if prefix:
                para.add_run(prefix + " ")
            self._add_page_field(para)
            return f"Page numbers added to {position}"
        except Exception as e:
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
    
    def remove_header(self, section_index: int = 0) -> str:
        """Remove header."""
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
            return f"Error: {e}"
    
    def remove_footer(self, section_index: int = 0) -> str:
        """Remove footer."""
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
            return f"Error: {e}"
    
    # ==================== Image Operations ====================
    
    def add_image(self, image_path: str, width_inches: Optional[float] = None,
                  height_inches: Optional[float] = None) -> str:
        """Add image to document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            if not os.path.exists(image_path):
                return f"Error: Image not found: {image_path}"
            para = self.current_document.add_paragraph()
            run = para.add_run()
            if width_inches and height_inches:
                run.add_picture(image_path, width=Inches(width_inches), height=Inches(height_inches))
            elif width_inches:
                run.add_picture(image_path, width=Inches(width_inches))
            elif height_inches:
                run.add_picture(image_path, height=Inches(height_inches))
            else:
                run.add_picture(image_path)
            return f"Image added: {image_path}"
        except Exception as e:
            return f"Error: {e}"
    
    def add_image_base64(self, base64_data: str, width_inches: Optional[float] = None) -> str:
        """Add image from base64 data."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            import base64
            import io
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
            return f"Error: {e}"
    
    def list_images(self) -> str:
        """List all images."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            images = []
            for i, shape in enumerate(self.current_document.inline_shapes):
                width = shape.width.inches if shape.width else "N/A"
                height = shape.height.inches if shape.height else "N/A"
                images.append(f"Image {i}: {width:.2f}\" x {height:.2f}\"")
            if not images:
                return "No images found"
            return f"Found {len(images)} images:\n" + "\n".join(images)
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== List Operations ====================
    
    def add_bulleted_list(self, items: List[str], style: str = "List Bullet") -> str:
        """Add bulleted list."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            for text in items:
                self.current_document.add_paragraph(text, style=style)
            return f"Bulleted list added with {len(items)} items"
        except Exception as e:
            return f"Error: {e}"
    
    def add_numbered_list(self, items: List[str], style: str = "List Number") -> str:
        """Add numbered list."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            for text in items:
                self.current_document.add_paragraph(text, style=style)
            return f"Numbered list added with {len(items)} items"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Page Layout ====================
    
    def set_page_margins(self, top: Optional[float] = None, bottom: Optional[float] = None,
                         left: Optional[float] = None, right: Optional[float] = None,
                         section_index: int = 0) -> str:
        """Set page margins in cm."""
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
            return f"Error: {e}"
    
    def set_page_size(self, width: float, height: float, section_index: int = 0) -> str:
        """Set page size in cm."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            section = sections[section_index]
            section.page_width = Cm(width)
            section.page_height = Cm(height)
            return f"Page size set to {width}cm x {height}cm"
        except Exception as e:
            return f"Error: {e}"
    
    def set_page_orientation(self, orientation: str, section_index: int = 0) -> str:
        """Set page orientation."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            section = sections[section_index]
            if orientation.lower() == "landscape":
                section.orientation = WD_ORIENT.LANDSCAPE
                new_width = section.page_height
                new_height = section.page_width
                section.page_width = new_width
                section.page_height = new_height
            elif orientation.lower() == "portrait":
                section.orientation = WD_ORIENT.PORTRAIT
                if section.page_width > section.page_height:
                    new_width = section.page_height
                    new_height = section.page_width
                    section.page_width = new_width
                    section.page_height = new_height
            else:
                return f"Error: Invalid orientation. Use 'portrait' or 'landscape'"
            return f"Page orientation set to {orientation}"
        except Exception as e:
            return f"Error: {e}"
    
    def get_page_info(self, section_index: int = 0) -> str:
        """Get page layout info."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            if section_index >= len(sections):
                return f"Error: Section index out of range"
            section = sections[section_index]
            def to_cm(emu):
                return round(emu / 360000, 2) if emu else 0
            info = []
            info.append(f"Section {section_index}:")
            info.append(f"  Page size: {to_cm(section.page_width)}cm x {to_cm(section.page_height)}cm")
            info.append(f"  Orientation: {'Landscape' if section.orientation == 1 else 'Portrait'}")
            info.append(f"  Margins: T={to_cm(section.top_margin)}cm, B={to_cm(section.bottom_margin)}cm, L={to_cm(section.left_margin)}cm, R={to_cm(section.right_margin)}cm")
            return "\n".join(info)
        except Exception as e:
            return f"Error: {e}"
    
    def add_section(self, section_type: str = "NEW_PAGE") -> str:
        """Add a new section."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            section_map = {
                'NEW_PAGE': WD_SECTION.NEW_PAGE,
                'CONTINUOUS': WD_SECTION.CONTINUOUS,
                'ODD_PAGE': WD_SECTION.ODD_PAGE,
                'EVEN_PAGE': WD_SECTION.EVEN_PAGE,
            }
            sec_type = section_map.get(section_type.upper(), WD_SECTION.NEW_PAGE)
            self.current_document.add_section(sec_type)
            return f"Section added with type: {section_type}"
        except Exception as e:
            return f"Error: {e}"
    
    def list_sections(self) -> str:
        """List all sections."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            sections = self.current_document.sections
            result = [f"Total sections: {len(sections)}"]
            for i, section in enumerate(sections):
                width = section.page_width.inches if section.page_width else "N/A"
                height = section.page_height.inches if section.page_height else "N/A"
                orient = "Landscape" if section.orientation == 1 else "Portrait"
                result.append(f"\nSection {i}: {orient}, {width:.2f}\" x {height:.2f}\"")
            return "\n".join(result)
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Styles ====================
    
    def list_styles(self) -> str:
        """List available paragraph styles."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            styles = self.current_document.styles
            paragraph_styles = []
            for style in styles:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    paragraph_styles.append(style.name)
            return "Paragraph styles:\n" + "\n".join(f"  - {s}" for s in sorted(paragraph_styles))
        except Exception as e:
            return f"Error: {e}"
    
    def apply_style(self, paragraph_index: int, style_name: str) -> str:
        """Apply style to paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            paragraphs[paragraph_index].style = style_name
            return f"Style '{style_name}' applied to paragraph {paragraph_index}"
        except Exception as e:
            return f"Error: {e}"
    
    def batch_format_paragraphs(self, style_name: str, start_index: int = 0, 
                                 end_index: Optional[int] = None) -> str:
        """Apply style to range of paragraphs."""
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
            return f"Error: {e}"
    
    # ==================== Metadata ====================
    
    def get_metadata(self) -> str:
        """Get document metadata."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            props = self.current_document.core_properties
            result = ["Document Metadata:"]
            result.append(f"  Title: {props.title or 'Not set'}")
            result.append(f"  Author: {props.author or 'Not set'}")
            result.append(f"  Subject: {props.subject or 'Not set'}")
            result.append(f"  Keywords: {props.keywords or 'Not set'}")
            result.append(f"  Created: {props.created or 'Not set'}")
            result.append(f"  Modified: {props.modified or 'Not set'}")
            return "\n".join(result)
        except Exception as e:
            return f"Error: {e}"
    
    def set_metadata(self, title: Optional[str] = None, author: Optional[str] = None,
                     subject: Optional[str] = None, keywords: Optional[str] = None) -> str:
        """Set document metadata."""
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
            return "Metadata updated"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Comments ====================
    
    def add_comment(self, comment_text: str, paragraph_index: int, 
                    author: str = "MCP", initials: str = "MCP") -> str:
        """Add a comment marker."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                paragraph_index = len(paragraphs) - 1
            para = paragraphs[paragraph_index]
            marker = f"[{author}: {comment_text}]"
            run = para.add_run(f" {marker}")
            run.font.italic = True
            run.font.color.rgb = RGBColor(120, 120, 120)
            self.comments.append({
                "paragraph": paragraph_index,
                "author": author,
                "text": comment_text,
            })
            return "Comment added"
        except Exception as e:
            return f"Error: {e}"
    
    def get_all_comments(self) -> str:
        """Get all comments."""
        if not self.comments:
            return "No comments in document"
        lines = []
        for idx, c in enumerate(self.comments):
            lines.append(f"{idx}: Para {c['paragraph']} by {c['author']} - {c['text']}")
        return "\n".join(lines)
    
    def delete_all_comments(self) -> str:
        """Delete all comments."""
        self.comments = []
        return "All comments removed"
    
    # ==================== Hyperlinks & Bookmarks ====================
    
    def add_hyperlink(self, paragraph_index: int, text: str, url: str) -> str:
        """Add hyperlink to paragraph."""
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
            rStyle = OxmlElement('w:rStyle')
            rStyle.set(qn('w:val'), 'Hyperlink')
            rPr.append(rStyle)
            new_run.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            new_run.append(t)
            hyperlink.append(new_run)
            paragraph._p.append(hyperlink)
            return "Hyperlink added"
        except Exception as e:
            return f"Error: {e}"
    
    def add_bookmark(self, paragraph_index: int, bookmark_name: str) -> str:
        """Add bookmark to paragraph."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                return f"Error: Invalid paragraph index"
            para = paragraphs[paragraph_index]
            bookmark_id = str(paragraph_index + 1)
            bookmark_start = OxmlElement('w:bookmarkStart')
            bookmark_start.set(qn('w:id'), bookmark_id)
            bookmark_start.set(qn('w:name'), bookmark_name)
            bookmark_end = OxmlElement('w:bookmarkEnd')
            bookmark_end.set(qn('w:id'), bookmark_id)
            para._p.insert(0, bookmark_start)
            para._p.append(bookmark_end)
            return f"Bookmark '{bookmark_name}' added"
        except Exception as e:
            return f"Error: {e}"
    
    def list_bookmarks(self) -> str:
        """List all bookmarks."""
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
            return f"Error: {e}"
    
    # ==================== Watermark ====================
    
    def add_text_watermark(self, text: str = "CONFIDENTIAL", section_index: int = 0) -> str:
        """Add text watermark."""
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
            return "Watermark added"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Document Merge ====================
    
    def merge_documents(self, file_paths: List[str]) -> str:
        """Merge multiple documents into current."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            count = 0
            for path in file_paths:
                if not os.path.exists(path):
                    return f"Error: File not found: {path}"
                src = Document(path)
                for para in src.paragraphs:
                    new_para = self.current_document.add_paragraph()
                    if para.style:
                        new_para.style = para.style
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                for tbl in src.tables:
                    rows = len(tbl.rows)
                    cols = len(tbl.columns)
                    new_tbl = self.current_document.add_table(rows=rows, cols=cols)
                    for r_idx, row in enumerate(tbl.rows):
                        for c_idx, cell in enumerate(row.cells):
                            new_tbl.cell(r_idx, c_idx).text = cell.text
                count += 1
            return f"Merged {count} documents"
        except Exception as e:
            return f"Error: {e}"
    
    def copy_document(self, dest_path: str) -> str:
        """Copy current document."""
        if not self.current_document:
            return "Error: No document is open"
        if not self.current_file_path:
            return "Error: No file path set"
        try:
            import shutil
            shutil.copy2(self.current_file_path, dest_path)
            return f"Document copied to: {dest_path}"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== TOC ====================
    
    def add_table_of_contents(self, title: str = "Table of Contents", heading_levels: int = 3) -> str:
        """Add Table of Contents."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            toc_title = self.current_document.add_paragraph(title)
            toc_title.style = 'Heading 1'
            para = self.current_document.add_paragraph()
            run = para.add_run()
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
            self.current_document.add_page_break()
            return "Table of Contents added"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Track Changes ====================
    
    def enable_track_changes(self) -> str:
        """Enable track changes."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            existing = settings.find(qn('w:trackRevisions'))
            if existing is None:
                track_elem = OxmlElement('w:trackRevisions')
                settings.append(track_elem)
            return "Track changes enabled"
        except Exception as e:
            return f"Error: {e}"
    
    def disable_track_changes(self) -> str:
        """Disable track changes."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            existing = settings.find(qn('w:trackRevisions'))
            if existing is not None:
                settings.remove(existing)
            return "Track changes disabled"
        except Exception as e:
            return f"Error: {e}"
    
    def accept_all_changes(self) -> str:
        """Accept all tracked changes."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            body = self.current_document.element.body
            for del_elem in body.findall('.//'+qn('w:del')):
                parent = del_elem.getparent()
                parent.remove(del_elem)
            for ins_elem in body.findall('.//'+qn('w:ins')):
                parent = ins_elem.getparent()
                index = list(parent).index(ins_elem)
                for child in ins_elem:
                    parent.insert(index, child)
                    index += 1
                parent.remove(ins_elem)
            return "All changes accepted"
        except Exception as e:
            return f"Error: {e}"
    
    def reject_all_changes(self) -> str:
        """Reject all tracked changes."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            body = self.current_document.element.body
            for ins_elem in body.findall('.//'+qn('w:ins')):
                parent = ins_elem.getparent()
                parent.remove(ins_elem)
            for del_elem in body.findall('.//'+qn('w:del')):
                parent = del_elem.getparent()
                index = list(parent).index(del_elem)
                for child in del_elem:
                    parent.insert(index, child)
                    index += 1
                parent.remove(del_elem)
            return "All changes rejected"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Protection ====================
    
    def protect_document(self, password: Optional[str] = None, 
                         protection_type: str = "readOnly") -> str:
        """Protect document."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            settings = self.current_document.settings.element
            existing = settings.find(qn('w:documentProtection'))
            if existing is not None:
                settings.remove(existing)
            prot = OxmlElement('w:documentProtection')
            edit_map = {'readOnly': 'readOnly', 'comments': 'comments', 
                       'trackedChanges': 'trackedChanges', 'forms': 'forms'}
            prot.set(qn('w:edit'), edit_map.get(protection_type, 'readOnly'))
            prot.set(qn('w:enforcement'), '1')
            if password:
                hash_val = hashlib.sha256(password.encode()).hexdigest()[:16]
                prot.set(qn('w:hash'), hash_val)
            settings.append(prot)
            return f"Document protected with type: {protection_type}"
        except Exception as e:
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
            return f"Error: {e}"
    
    # ==================== Footnotes ====================
    
    def add_footnote(self, text: str, paragraph_index: int) -> str:
        """Add a footnote."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = self.current_document.paragraphs
            if paragraph_index >= len(paragraphs):
                paragraph_index = len(paragraphs) - 1
            footnote_number = len(self.footnotes) + 1
            target_para = paragraphs[paragraph_index]
            run = target_para.add_run(f"[{footnote_number}]")
            run.font.superscript = True
            self.footnotes.append(text)
            return f"Footnote {footnote_number} added"
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== Statistics ====================
    
    def get_word_count(self) -> str:
        """Get word count statistics."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            word_count = 0
            char_count = 0
            para_count = len(self.current_document.paragraphs)
            for para in self.current_document.paragraphs:
                text = para.text
                word_count += len(text.split())
                char_count += len(text)
            return f"Words: {word_count}, Characters: {char_count}, Paragraphs: {para_count}"
        except Exception as e:
            return f"Error: {e}"
    
    def get_document_statistics(self) -> str:
        """Get comprehensive statistics."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            paragraphs = len(self.current_document.paragraphs)
            tables = len(self.current_document.tables)
            sections = len(self.current_document.sections)
            word_count = sum(len(p.text.split()) for p in self.current_document.paragraphs)
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
- Headings: {sum(heading_counts.values())}"""
            return stats
        except Exception as e:
            return f"Error: {e}"
    
    def get_document_outline(self) -> str:
        """Get document outline from headings."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            outline = []
            for para in self.current_document.paragraphs:
                style_name = para.style.name if para.style else ""
                if style_name.startswith("Heading"):
                    level = style_name.replace("Heading ", "")
                    indent = "  " * (int(level) - 1) if level.isdigit() else ""
                    outline.append(f"{indent}{para.text[:60]}")
            if not outline:
                return "No headings found"
            return "Document Outline:\n" + "\n".join(outline)
        except Exception as e:
            return f"Error: {e}"
    
    def get_document_structure(self) -> str:
        """Get structural overview."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            result = ["Document Structure:"]
            result.append(f"Sections: {len(self.current_document.sections)}")
            result.append(f"Paragraphs: {len(self.current_document.paragraphs)}")
            result.append(f"Tables: {len(self.current_document.tables)}")
            result.append("\nHeadings:")
            for i, para in enumerate(self.current_document.paragraphs):
                style = para.style.name if para.style else ""
                if style.startswith("Heading"):
                    level = style.replace("Heading ", "")
                    indent = "  " * (int(level) if level.isdigit() else 1)
                    result.append(f"{indent}[{i}] {para.text[:50]}")
            return "\n".join(result)
        except Exception as e:
            return f"Error: {e}"
    
    # ==================== File Operations ====================
    
    def list_docx_files(self, directory: str) -> str:
        """List docx files in directory."""
        try:
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
            return f"Error: {e}"
    
    # ==================== Mail Merge ====================
    
    def add_merge_field(self, field_name: str) -> str:
        """Add mail merge field."""
        if not self.current_document:
            return "Error: No document is open"
        try:
            para = self.current_document.add_paragraph()
            run = para.add_run()
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
            para.add_run(f'«{field_name}»')
            run3 = para.add_run()
            run3._r.append(fldChar3)
            return f"Merge field '{field_name}' added"
        except Exception as e:
            return f"Error: {e}"
    
    def execute_mail_merge(self, data: List[Dict[str, str]], output_pattern: str) -> str:
        """Execute mail merge."""
        if not self.current_document:
            return "Error: No document is open"
        if not self.current_file_path:
            return "Error: No file path set"
        try:
            created = []
            for i, record in enumerate(data):
                new_doc = Document(self.current_file_path)
                for para in new_doc.paragraphs:
                    for field_name, value in record.items():
                        placeholder = f'«{field_name}»'
                        if placeholder in para.text:
                            for run in para.runs:
                                if placeholder in run.text:
                                    run.text = run.text.replace(placeholder, str(value))
                for table in new_doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                for field_name, value in record.items():
                                    placeholder = f'«{field_name}»'
                                    if placeholder in para.text:
                                        for run in para.runs:
                                            if placeholder in run.text:
                                                run.text = run.text.replace(placeholder, str(value))
                output_path = output_pattern.format(index=i)
                new_doc.save(output_path)
                created.append(output_path)
            return f"Created {len(created)} documents"
        except Exception as e:
            return f"Error: {e}"
