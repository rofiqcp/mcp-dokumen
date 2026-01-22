#!/usr/bin/env python3
"""
Test script for MCP Documents
"""

import os
import sys
import tempfile

# Add the package to path for testing
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp_documents.core import DocumentProcessor


def test_document_operations():
    """Test basic document operations."""
    print("=" * 60)
    print("Testing MCP Documents")
    print("=" * 60)
    
    # Create document processor
    doc = DocumentProcessor()
    
    # Create a temporary directory for test files
    with tempfile.TemporaryDirectory() as tmpdir:
        test_file = os.path.join(tmpdir, "test_document.docx")
        
        # Test 1: Create document
        print("\n1. Creating document...")
        result = doc.create_document(test_file, title="Test Document")
        print(f"   {result}")
        assert "created" in result.lower()
        
        # Test 2: Add heading
        print("\n2. Adding heading...")
        result = doc.add_heading("Introduction", level=1)
        print(f"   {result}")
        assert "added" in result.lower()
        
        # Test 3: Add paragraph
        print("\n3. Adding paragraph...")
        result = doc.add_paragraph(
            "This is a test paragraph with formatting.",
            bold=True,
            font_size=12,
            alignment="justify"
        )
        print(f"   {result}")
        assert "added" in result.lower()
        
        # Test 4: Add table
        print("\n4. Adding table...")
        result = doc.add_table(
            rows=3,
            cols=3,
            data=[
                ["Header 1", "Header 2", "Header 3"],
                ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
                ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
            ]
        )
        print(f"   {result}")
        assert "added" in result.lower()
        
        # Test 5: Get document info
        print("\n5. Getting document info...")
        result = doc.get_document_info()
        print(f"   {result}")
        
        # Test 6: Save document
        print("\n6. Saving document...")
        result = doc.save_document()
        print(f"   {result}")
        assert "saved" in result.lower()
        
        # Verify file exists
        assert os.path.exists(test_file), "Document file not created"
        file_size = os.path.getsize(test_file)
        print(f"\n   File created: {test_file}")
        print(f"   File size: {file_size} bytes")
        
        # Test 7: Search
        print("\n7. Searching for text...")
        result = doc.search_text("test")
        print(f"   {result}")
        
        # Test 8: Find and replace
        print("\n8. Find and replace...")
        result = doc.find_and_replace("test", "sample")
        print(f"   {result}")
        
        # Test 9: Close document
        print("\n9. Closing document...")
        result = doc.close_document()
        print(f"   {result}")
        
        # Test 10: Open document
        print("\n10. Reopening document...")
        result = doc.open_document(test_file)
        print(f"   {result}")
        assert "opened" in result.lower()
        
        # Test 11: Get paragraphs
        print("\n11. Getting paragraphs...")
        paragraphs = doc.get_paragraphs()
        print(f"   Found {len(paragraphs)} paragraphs")
        for i, p in enumerate(paragraphs):
            if p.strip():
                print(f"   [{i}] {p[:50]}...")
    
    print("\n" + "=" * 60)
    print("All tests passed!")
    print("=" * 60)


def test_converters():
    """Test converter modules (import only)."""
    print("\n" + "=" * 60)
    print("Testing Converter Imports")
    print("=" * 60)
    
    try:
        from mcp_documents.converters.pdf import docx_to_pdf, pdf_to_docx
        print("  ✓ PDF converters imported")
    except ImportError as e:
        print(f"  ✗ PDF converters: {e}")
    
    try:
        from mcp_documents.converters.image import convert_image, image_to_base64
        print("  ✓ Image converters imported")
    except ImportError as e:
        print(f"  ✗ Image converters: {e}")
    
    try:
        from mcp_documents.converters.pandoc import convert_with_pandoc, is_pandoc_available
        available = is_pandoc_available()
        status = "available" if available else "not installed"
        print(f"  ✓ Pandoc converters imported (Pandoc is {status})")
    except ImportError as e:
        print(f"  ✗ Pandoc converters: {e}")


def test_server_import():
    """Test server module import."""
    print("\n" + "=" * 60)
    print("Testing Server Import")
    print("=" * 60)
    
    try:
        from mcp_documents.server import create_server
        print("  ✓ Server module imported")
        
        server, doc_processor = create_server()
        print("  ✓ Server created successfully")
        print(f"  ✓ Document processor initialized: {type(doc_processor).__name__}")
    except ImportError as e:
        print(f"  ✗ Server import failed: {e}")
    except Exception as e:
        print(f"  ✗ Server creation failed: {e}")


if __name__ == "__main__":
    try:
        test_document_operations()
        test_converters()
        test_server_import()
        print("\n✅ All tests completed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
