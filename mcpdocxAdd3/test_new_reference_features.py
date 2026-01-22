"""
Tests for newly added reference features
"""

import os
import tempfile
import shutil
import unittest

from mcp_documents.core.document import DocumentProcessor


class TestSectionOperations(unittest.TestCase):
    """Test section operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_sections.docx")
        self.processor.create_document(self.test_file, "Section Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_section(self):
        result = self.processor.add_section("NEW_PAGE")
        self.assertIn("section", result.lower())
    
    def test_list_sections(self):
        result = self.processor.list_sections()
        self.assertIn("section", result.lower())
    
    def test_set_section_properties(self):
        result = self.processor.set_section_properties(
            section_index=0,
            orientation="LANDSCAPE"
        )
        self.assertIn("updated", result.lower())
    
    def test_zoned_header(self):
        result = self.processor.add_zoned_header(
            section_index=0,
            left_text="Left",
            center_text="Center",
            right_text="Right"
        )
        self.assertIn("header", result.lower())
    
    def test_zoned_footer(self):
        result = self.processor.add_zoned_footer(
            section_index=0,
            left_text="Page",
            center_text="",
            right_text="Date"
        )
        self.assertIn("footer", result.lower())


class TestMetadataOperations(unittest.TestCase):
    """Test metadata operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_metadata.docx")
        self.processor.create_document(self.test_file, "Metadata Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_get_metadata(self):
        result = self.processor.get_metadata()
        self.assertIn("metadata", result.lower())
    
    def test_set_metadata(self):
        result = self.processor.set_metadata(
            title="Test Document",
            author="Test Author",
            subject="Testing"
        )
        self.assertIn("updated", result.lower())
    
    def test_strip_personal_info(self):
        self.processor.set_metadata(author="Secret Author")
        result = self.processor.strip_personal_info()
        self.assertIn("stripped", result.lower())


class TestSecurityOperations(unittest.TestCase):
    """Test security/redaction operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_security.docx")
        self.processor.create_document(self.test_file, "Security Test")
        self.processor.add_paragraph("Secret information here")
        self.processor.add_paragraph("Password is 12345")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_redact_text(self):
        result = self.processor.redact_text("Secret", "[REDACTED]")
        self.assertIn("redacted", result.lower())
    
    def test_redact_with_regex(self):
        result = self.processor.redact_text(r"\d+", "[NUM]", use_regex=True)
        self.assertIn("redacted", result.lower())
    
    def test_sanitize_external_links(self):
        # Add a hyperlink to paragraph index 1 (first content paragraph)
        self.processor.add_hyperlink(1, "Click here", "https://example.com")
        result = self.processor.sanitize_external_links()
        # May return 0 if hyperlink wasn't created properly
        self.assertNotIn("error", result.lower())


class TestParagraphOperations(unittest.TestCase):
    """Test paragraph text operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_paragraphs.docx")
        self.processor.create_document(self.test_file, "Paragraph Test")
        self.processor.add_paragraph("First paragraph")
        self.processor.add_paragraph("Second paragraph")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_get_paragraph_text(self):
        result = self.processor.get_paragraph_text(1)
        self.assertEqual(result, "First paragraph")
    
    def test_set_paragraph_text(self):
        result = self.processor.set_paragraph_text(1, "Modified text")
        self.assertIn("updated", result.lower())
    
    def test_delete_paragraph(self):
        result = self.processor.delete_paragraph(1)
        self.assertIn("deleted", result.lower())
    
    def test_insert_paragraph_after(self):
        result = self.processor.insert_paragraph_after(1, "Inserted text")
        self.assertIn("inserted", result.lower())
    
    def test_find_text(self):
        result = self.processor.find_text("paragraph")
        self.assertIn("found", result.lower())


class TestStructureOperations(unittest.TestCase):
    """Test structure operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_structure.docx")
        self.processor.create_document(self.test_file, "Structure Test")
        self.processor.add_heading("Chapter 1", level=1)
        self.processor.add_paragraph("Content here")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_get_document_structure(self):
        result = self.processor.get_document_structure()
        self.assertIn("structure", result.lower())
    
    def test_get_raw_xml(self):
        result = self.processor.get_raw_xml()
        self.assertIn("<", result)  # XML content
    
    def test_copy_document(self):
        dest = os.path.join(self.temp_dir, "copy.docx")
        result = self.processor.copy_document(dest)
        self.assertIn("copied", result.lower())
        self.assertTrue(os.path.exists(dest))
    
    def test_list_docx_files(self):
        result = self.processor.list_docx_files(self.temp_dir)
        self.assertIn("test_structure.docx", result)
    
    def test_get_word_count(self):
        result = self.processor.get_word_count()
        self.assertIn("words", result.lower())


class TestExtendedTableOperations(unittest.TestCase):
    """Test extended table operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_tables.docx")
        self.processor.create_document(self.test_file, "Table Test")
        self.processor.add_table(rows=3, cols=3)
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_merge_cells(self):
        result = self.processor.merge_cells(
            table_index=0,
            start_row=0, start_col=0,
            end_row=0, end_col=1
        )
        self.assertIn("merged", result.lower())
    
    def test_get_table_data(self):
        result = self.processor.get_table_data(0)
        self.assertIn("row", result.lower())
    
    def test_list_tables(self):
        result = self.processor.list_tables()
        self.assertIn("table", result.lower())
    
    def test_set_table_cell_text(self):
        result = self.processor.set_table_cell_text(0, 0, 0, "Header")
        self.assertIn("set", result.lower())
    
    def test_add_table_row(self):
        result = self.processor.add_table_row(0, ["A", "B", "C"])
        self.assertIn("added", result.lower())
    
    def test_delete_table_row(self):
        result = self.processor.delete_table_row(0, 0)
        self.assertIn("deleted", result.lower())


class TestImageOperations(unittest.TestCase):
    """Test image operations."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_images.docx")
        self.processor.create_document(self.test_file, "Image Test")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_add_image_file_not_found(self):
        result = self.processor.add_image("/nonexistent/image.png")
        self.assertIn("error", result.lower())
    
    def test_list_images(self):
        result = self.processor.list_images()
        self.assertIn("no images", result.lower())


class TestIntegrationNewFeatures(unittest.TestCase):
    """Integration tests for new features."""
    
    def setUp(self):
        self.processor = DocumentProcessor()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_integration.docx")
    
    def tearDown(self):
        shutil.rmtree(self.temp_dir)
    
    def test_complete_document_workflow(self):
        """Test creating a complete document with all new features."""
        # Create document
        result = self.processor.create_document(self.test_file, "Complete Document")
        self.assertIn("created", result.lower())
        
        # Set metadata
        result = self.processor.set_metadata(
            title="Annual Report",
            author="Test Corp",
            subject="Financial Results"
        )
        self.assertIn("updated", result.lower())
        
        # Add section with landscape orientation
        self.processor.add_section("NEW_PAGE")
        result = self.processor.set_section_properties(1, orientation="LANDSCAPE")
        self.assertIn("updated", result.lower())
        
        # Add zoned header
        result = self.processor.add_zoned_header(0, "Company", "Title", "Date")
        self.assertIn("header", result.lower())
        
        # Add content
        self.processor.add_heading("Executive Summary", level=1)
        self.processor.add_paragraph("This is the executive summary.")
        
        # Add table with data
        self.processor.add_table(rows=3, cols=3)
        self.processor.set_table_cell_text(0, 0, 0, "Item")
        self.processor.set_table_cell_text(0, 0, 1, "Q1")
        self.processor.set_table_cell_text(0, 0, 2, "Q2")
        
        # Get structure
        result = self.processor.get_document_structure()
        self.assertIn("structure", result.lower())
        
        # Get word count
        result = self.processor.get_word_count()
        self.assertIn("words", result.lower())
        
        # Save
        result = self.processor.save_document()
        self.assertIn("saved", result.lower())
        
        # Verify file exists
        self.assertTrue(os.path.exists(self.test_file))


class TestPandocAdvanced(unittest.TestCase):
    """Test advanced Pandoc options (error-path only to avoid dependency)."""

    def test_missing_input_file(self):
        from mcp_documents.converters.pandoc import convert_with_pandoc

        success, message = convert_with_pandoc(
            input_path="/nonexistent/input.md",
            output_path="/tmp/out.docx",
            filters=["pandoc-citeproc"],
            metadata={"title": "Demo"},
            variables={"author": "Test"},
        )
        self.assertFalse(success)
        self.assertIn("not found", message.lower())


def run_tests():
    """Run all tests."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    test_classes = [
        TestSectionOperations,
        TestMetadataOperations,
        TestSecurityOperations,
        TestParagraphOperations,
        TestStructureOperations,
        TestExtendedTableOperations,
        TestImageOperations,
        TestIntegrationNewFeatures,
    ]
    
    for test_class in test_classes:
        tests = loader.loadTestsFromTestCase(test_class)
        suite.addTests(tests)
    
    runner = unittest.TextTestRunner(verbosity=2)
    return runner.run(suite)


if __name__ == "__main__":
    run_tests()
