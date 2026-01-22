"""
Extended feature tests for MCP Documents.
Run with: python -m pytest -q test_extended_features.py
"""

import os
import tempfile
from mcp_documents.core import DocumentProcessor


def test_headers_footers_and_page_numbers():
    doc = DocumentProcessor()
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "doc.docx")
        doc.create_document(path, title="Test")
        assert "Header" in doc.add_header("Header Text")
        assert "Footer" in doc.add_footer("Footer Text")
        assert "Page numbers" in doc.add_page_numbers()
        doc.save_document()
        assert os.path.exists(path)


def test_lists_and_hyperlink():
    doc = DocumentProcessor()
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "lists.docx")
        doc.create_document(path)
        doc.add_bulleted_list(["A", "B", "C"])
        doc.add_numbered_list(["One", "Two"])
        doc.add_list_item("Extra", list_type="number")
        doc.add_paragraph("Link here")
        doc.add_hyperlink(paragraph_index=1, text="OpenAI", url="https://openai.com")
        doc.save_document()
        assert os.path.exists(path)


def test_footnotes_and_comments():
    doc = DocumentProcessor()
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "notes.docx")
        doc.create_document(path)
        doc.add_paragraph("Paragraph for note")
        doc.add_footnote("This is a footnote", paragraph_index=1)
        doc.add_comment("Needs review", paragraph_index=1, author="QA", initials="QA")
        comments_summary = doc.get_all_comments()
        assert "QA" in comments_summary and "Needs review" in comments_summary
        doc.save_document()
        assert os.path.exists(path)


def test_advanced_table_formatting():
    doc = DocumentProcessor()
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "table.docx")
        doc.create_document(path)
        doc.add_table(3, 3, data=[["H1", "H2", "H3"], ["a", "b", "c"], ["d", "e", "f"]])
        doc.highlight_table_header(0)
        doc.apply_table_alternating_rows(0)
        doc.set_table_cell_shading(0, 1, 1, "FFEEEE")
        doc.set_cell_alignment(0, 1, 1, "center", "center")
        doc.set_cell_padding(0, 1, 1, 0.2)
        doc.set_column_width(0, 0, 3.0)
        doc.save_document()
        assert os.path.exists(path)


def test_watermark_and_merge():
    doc = DocumentProcessor()
    with tempfile.TemporaryDirectory() as tmpdir:
        base = os.path.join(tmpdir, "base.docx")
        other = os.path.join(tmpdir, "other.docx")
        doc.create_document(base, title="Base")
        doc.add_paragraph("Base content")
        doc.save_document()

        # create other doc
        other_doc = DocumentProcessor()
        other_doc.create_document(other, title="Other")
        other_doc.add_paragraph("Other content")
        other_doc.save_document()

        # reopen base and merge
        doc.open_document(base)
        doc.add_text_watermark("CONFIDENTIAL")
        doc.merge_documents([other])
        doc.save_document()
        assert os.path.exists(base)


if __name__ == "__main__":
    # Quick manual run
    test_headers_footers_and_page_numbers()
    test_lists_and_hyperlink()
    test_footnotes_and_comments()
    test_advanced_table_formatting()
    test_watermark_and_merge()
    print("All extended feature tests passed")
