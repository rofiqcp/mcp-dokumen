# MCP Documents

**Unified MCP Server for Document Processing**

A comprehensive Model Context Protocol (MCP) server that combines the best features from multiple document processing libraries into a single, efficient Python 3.10 compatible engine.

## Features

### 📄 Word Document Operations (DOCX)
- **Create/Open/Save** - Full document lifecycle management
- **Content** - Paragraphs, headings, tables, images
- **Formatting** - Font styles, colors, alignment, spacing
- **Tables** - Create, edit, merge cells, add/delete rows
- **Sections** - Page layout, margins, orientation
- **Headers/Footers** - Add, modify, zone-based headers
- **Styles** - Built-in and custom styles management
- **Search/Replace** - Find and replace with preview

### 🔄 Format Conversion
- **DOCX ↔ PDF** - Bidirectional conversion
- **Markdown → DOCX/PDF** - With styling support
- **Markdown → PPTX** - PowerPoint presentations
- **Image formats** - PNG, JPG, WebP, GIF, BMP, TIFF
- **Excel → CSV** - Spreadsheet conversion
- **HTML → PDF** - Web page to document

### 🎨 Advanced Features
- **Template support** - Reference documents for styling
- **Pandoc integration** - Advanced format conversion
- **Base64 I/O** - Direct content handling without files
- **Batch operations** - Multiple edits in one call

## Installation

```bash
cd MCP-Documents
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

### Optional Dependencies

For PDF conversion:
```bash
pip install -e ".[pdf]"
```

For all features:
```bash
pip install -e ".[full]"
```

### External Tools

Some features require external tools:
- **Pandoc** - For advanced format conversion: `sudo apt install pandoc`
- **LibreOffice** - For DOCX→PDF on Linux: `sudo apt install libreoffice`
- **wkhtmltopdf** - For HTML→PDF: `sudo apt install wkhtmltopdf`

## Configuration

### Claude Desktop / VS Code

Add to your MCP configuration:

```json
{
  "mcpServers": {
    "mcp-documents": {
      "command": "/path/to/MCP-Documents/.venv/bin/python",
      "args": ["-m", "mcp_documents.server"]
    }
  }
}
```

## Usage Examples

### Create a Document
```
"Create a new Word document called report.docx with title 'Q4 Report'"
```

### Add Content
```
"Add a heading 'Summary' at level 2, then add a paragraph with the quarterly results"
```

### Format Text
```
"Make the title bold and center-aligned, set font size to 24pt"
```

### Add Tables
```
"Insert a 3x4 table with sales data, merge the header row"
```

### Convert Formats
```
"Convert report.docx to PDF"
"Convert this markdown to PowerPoint and save as presentation.pptx"
```

## Available Tools

### Document Management
| Tool | Description |
|------|-------------|
| `create_document` | Create new Word document |
| `open_document` | Open existing document |
| `save_document` | Save current document |
| `save_as_document` | Save as new file |
| `get_document_info` | Get document statistics |

### Content Operations
| Tool | Description |
|------|-------------|
| `add_paragraph` | Add paragraph with formatting |
| `add_heading` | Add heading (level 1-9) |
| `add_table` | Create table with data |
| `add_image` | Insert image |
| `add_page_break` | Insert page break |

### Table Operations
| Tool | Description |
|------|-------------|
| `edit_table_cell` | Modify cell content |
| `add_table_row` | Add row to table |
| `delete_table_row` | Remove row |
| `merge_table_cells` | Merge cell range |

### Format Conversion
| Tool | Description |
|------|-------------|
| `convert_docx_to_pdf` | DOCX → PDF |
| `convert_pdf_to_docx` | PDF → DOCX |
| `convert_markdown` | MD → DOCX/PDF/PPTX |
| `convert_image` | Image format conversion |

### Search & Edit
| Tool | Description |
|------|-------------|
| `search_text` | Find text in document |
| `find_and_replace` | Replace text |
| `delete_paragraph` | Remove paragraph |

## Architecture

```
mcp_documents/
├── __init__.py          # Package exports
├── server.py            # MCP server entry point
├── core/
│   ├── document.py      # Document processor class
│   ├── formatting.py    # Text/paragraph formatting
│   └── utils.py         # Helper functions
├── tools/
│   ├── document_ops.py  # Document CRUD operations
│   ├── content_ops.py   # Content manipulation
│   ├── table_ops.py     # Table operations
│   ├── style_ops.py     # Style management
│   ├── section_ops.py   # Section/layout operations
│   ├── header_footer.py # Headers and footers
│   └── conversion.py    # Format conversion tools
└── converters/
    ├── pdf.py           # PDF conversion
    ├── image.py         # Image conversion
    └── pandoc.py        # Pandoc wrapper
```

## License

MIT License
