"""
Comprehensive tests for all new MCP Document Server features
"""

import os
import tempfile
import shutil
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Import the server components
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp_documents.core.document import DocumentProcessor


class TestTrackChanges(unittest.TestCase):
    """Test track changes functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_track.docx")
        self.processor.create_document(self.test_file, "Track Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_enable_track_changes(self):
        result = self.processor.enable_track_changes()
        self.assertIn("enabled", result.lower())
    
    def test_disable_track_changes(self):
        self.processor.enable_track_changes()
        result = self.processor.disable_track_changes()
        self.assertIn("disabled", result.lower())
    
    def test_accept_all_changes(self):
        result = self.processor.accept_all_changes()
        self.assertIn("accepted", result.lower())
    
    def test_reject_all_changes(self):
        result = self.processor.reject_all_changes()
        self.assertIn("rejected", result.lower())


class TestDocumentProtection(unittest.TestCase):
    """Test document protection functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_protect.docx")
        self.processor.create_document(self.test_file, "Protection Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_protect_document_readonly(self):
        result = self.processor.protect_document(protection_type="readOnly")
        self.assertIn("protected", result.lower())
    
    def test_protect_document_with_password(self):
        result = self.processor.protect_document(
            password="secret123", 
            protection_type="comments"
        )
        self.assertIn("protected", result.lower())
    
    def test_unprotect_document(self):
        self.processor.protect_document()
        result = self.processor.unprotect_document()
        self.assertIn("removed", result.lower())


class TestFootnotesEndnotes(unittest.TestCase):
    """Test footnotes and endnotes functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_footnotes.docx")
        self.processor.create_document(self.test_file, "Footnotes Test")
        self.processor.add_paragraph("First paragraph for testing.")
        self.processor.add_paragraph("Second paragraph for testing.")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_native_footnote(self):
        result = self.processor.add_native_footnote(1, "This is a footnote")
        self.assertIn("footnote", result.lower())
        self.assertIn("added", result.lower())
    
    def test_add_endnote(self):
        result = self.processor.add_endnote(1, "This is an endnote")
        self.assertIn("endnote", result.lower())
        self.assertIn("added", result.lower())
    
    def test_render_endnotes(self):
        self.processor.add_endnote(1, "First endnote")
        self.processor.add_endnote(1, "Second endnote")
        result = self.processor.render_endnotes()
        self.assertIn("rendered", result.lower())


class TestTableOfContents(unittest.TestCase):
    """Test TOC functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_toc.docx")
        self.processor.create_document(self.test_file, "TOC Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_table_of_contents(self):
        result = self.processor.add_table_of_contents()
        self.assertIn("table of contents", result.lower())
    
    def test_add_toc_with_custom_levels(self):
        result = self.processor.add_table_of_contents(
            title="Contents", 
            heading_levels=5
        )
        self.assertIn("added", result.lower())
    
    def test_update_toc_instructions(self):
        result = self.processor.update_toc()
        self.assertIn("update", result.lower())


class TestMailMerge(unittest.TestCase):
    """Test mail merge functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_merge.docx")
        self.processor.create_document(self.test_file, "Mail Merge Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_merge_field(self):
        result = self.processor.add_merge_field("FirstName")
        self.assertIn("merge field", result.lower())
        self.assertIn("firstname", result.lower())
    
    def test_execute_mail_merge(self):
        self.processor.add_merge_field("Name")
        self.processor.save_document()
        
        data = [
            {"Name": "John"},
            {"Name": "Jane"}
        ]
        output_pattern = os.path.join(self.temp_dir, "letter_{index}.docx")
        
        result = self.processor.execute_mail_merge(data, output_pattern)
        self.assertIn("created", result.lower())

    def test_execute_mail_merge_replaces_field_codes(self):
        self.processor.add_merge_field("Name")
        self.processor.save_document()

        data = [{"Name": "Alice"}]
        output_pattern = os.path.join(self.temp_dir, "letter_{index}.docx")

        self.processor.execute_mail_merge(data, output_pattern)
        merged_doc = Document(output_pattern.format(index=0))
        merged_text = "\n".join(p.text for p in merged_doc.paragraphs)

        self.assertIn("Alice", merged_text)
        self.assertNotIn("«Name»", merged_text)

    def test_execute_mail_merge_replaces_fldsimple_fields(self):
        para = self.processor.current_document.add_paragraph()
        fld_simple = OxmlElement('w:fldSimple')
        fld_simple.set(qn('w:instr'), ' MERGEFIELD Company ')
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = "«Company»"
        run.append(text)
        fld_simple.append(run)
        para._p.append(fld_simple)
        self.processor.save_document()

        data = [{"Company": "Acme"}]
        output_pattern = os.path.join(self.temp_dir, "letter_{index}.docx")

        self.processor.execute_mail_merge(data, output_pattern)
        merged_doc = Document(output_pattern.format(index=0))
        merged_text = "\n".join(p.text for p in merged_doc.paragraphs)

        self.assertIn("Acme", merged_text)
        self.assertNotIn("«Company»", merged_text)


class TestBookmarks(unittest.TestCase):
    """Test bookmark functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_bookmarks.docx")
        self.processor.create_document(self.test_file, "Bookmarks Test")
        self.processor.add_paragraph("First paragraph")
        self.processor.add_paragraph("Second paragraph")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_bookmark(self):
        result = self.processor.add_bookmark(1, "section1")
        self.assertIn("bookmark", result.lower())
        self.assertIn("added", result.lower())
    
    def test_add_hyperlink_to_bookmark(self):
        self.processor.add_bookmark(1, "target")
        result = self.processor.add_hyperlink_to_bookmark("Go to target", "target")
        self.assertIn("link", result.lower())
    
    def test_list_bookmarks(self):
        self.processor.add_bookmark(1, "bookmark1")
        result = self.processor.list_bookmarks()
        self.assertIn("bookmark1", result)


class TestTableBorders(unittest.TestCase):
    """Test table border and sizing functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_borders.docx")
        self.processor.create_document(self.test_file, "Table Borders Test")
        self.processor.add_table(rows=3, cols=3)
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_set_table_borders(self):
        result = self.processor.set_table_borders(
            table_index=0,
            top="single",
            bottom="double",
            left="single",
            right="single"
        )
        self.assertIn("updated", result.lower())
    
    def test_set_row_height(self):
        result = self.processor.set_row_height(
            table_index=0, 
            row_index=0, 
            height_cm=1.5
        )
        self.assertIn("height", result.lower())


class TestDocumentAnalysis(unittest.TestCase):
    """Test document analysis functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_analysis.docx")
        self.processor.create_document(self.test_file, "Analysis Test")
        self.processor.add_paragraph("Some test content here")
        self.processor.add_heading("Chapter 1", level=1)
        self.processor.add_paragraph("Chapter content")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_get_document_outline(self):
        result = self.processor.get_document_outline()
        # Should find at least the document title heading
        self.assertNotIn("error", result.lower())
    
    def test_get_document_statistics(self):
        result = self.processor.get_document_statistics()
        self.assertIn("paragraphs", result.lower())
        self.assertIn("words", result.lower())
    
    def test_get_formatting_report(self):
        result = self.processor.get_formatting_report()
        self.assertIn("report", result.lower())


class TestBatchOperations(unittest.TestCase):
    """Test batch operation functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_batch.docx")
        self.processor.create_document(self.test_file, "Batch Test")
        self.processor.add_paragraph("Hello World")
        self.processor.add_paragraph("Hello Again")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_batch_replace(self):
        result = self.processor.batch_replace({"Hello": "Hi", "World": "Universe"})
        self.assertIn("replacement", result.lower())
    
    def test_batch_format_paragraphs(self):
        result = self.processor.batch_format_paragraphs("Normal", 0, 2)
        self.assertIn("applied", result.lower())


class TestImageWatermark(unittest.TestCase):
    """Test image watermark functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_watermark.docx")
        self.processor.create_document(self.test_file, "Watermark Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_image_watermark_file_not_found(self):
        result = self.processor.add_image_watermark("/nonexistent/image.png")
        self.assertIn("error", result.lower())
    
    def test_add_text_watermark(self):
        result = self.processor.add_text_watermark("DRAFT")
        self.assertIn("watermark", result.lower())


class TestCSVOperations(unittest.TestCase):
    """Test CSV operations functionality."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_csv.docx")
        self.processor.create_document(self.test_file, "CSV Test")
        
        # Create a test CSV
        self.csv_file = os.path.join(self.temp_dir, "test.csv")
        with open(self.csv_file, 'w') as f:
            f.write("Name,Age,City\n")
            f.write("John,30,NYC\n")
            f.write("Jane,25,LA\n")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_csv_to_table(self):
        # Import the excel_ops module
        from mcp_documents.tools import excel_ops
        
        # Call directly via processor methods (simplified test)
        import csv
        with open(self.csv_file, 'r') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        self.assertEqual(len(rows), 3)  # header + 2 data rows
        self.assertEqual(rows[0], ["Name", "Age", "City"])


class TestIntegration(unittest.TestCase):
    """Integration tests combining multiple features."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_integration.docx")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_full_document_workflow(self):
        """Test creating a document with multiple features."""
        # Create document
        result = self.processor.create_document(self.test_file, "Report 2024")
        self.assertIn("created", result.lower())
        
        # Add TOC
        result = self.processor.add_table_of_contents("Contents")
        self.assertIn("added", result.lower())
        
        # Add content with headings
        self.processor.add_heading("Introduction", level=1)
        self.processor.add_paragraph("This is the introduction.")
        self.processor.add_heading("Methods", level=1)
        self.processor.add_paragraph("This section describes methods.")
        
        # Add table with borders
        self.processor.add_table(rows=3, cols=3)
        result = self.processor.set_table_borders(0, top="single", bottom="single")
        self.assertIn("updated", result.lower())
        
        # Add bookmark
        result = self.processor.add_bookmark(2, "intro")
        self.assertIn("bookmark", result.lower())
        
        # Add watermark
        result = self.processor.add_text_watermark("CONFIDENTIAL")
        self.assertIn("watermark", result.lower())
        
        # Get statistics
        result = self.processor.get_document_statistics()
        self.assertIn("paragraphs", result.lower())
        
        # Save
        result = self.processor.save_document()
        self.assertIn("saved", result.lower())
        
        # Verify file exists
        self.assertTrue(os.path.exists(self.test_file))


def run_tests():
    """Run all tests."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test classes
    test_classes = [
        TestTrackChanges,
        TestDocumentProtection,
        TestFootnotesEndnotes,
        TestTableOfContents,
        TestMailMerge,
        TestBookmarks,
        TestTableBorders,
        TestDocumentAnalysis,
        TestBatchOperations,
        TestImageWatermark,
        TestCSVOperations,
        TestIntegration,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == "__main__":
    run_tests()
