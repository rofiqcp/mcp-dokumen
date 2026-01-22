# MCP DOCX Unified Server

🔧 **90 Tools for Reading, Generating, and Editing Word Documents**

A comprehensive Model Context Protocol (MCP) server that provides 90 specialized tools for working with Microsoft Word documents (.docx).

## Features

- **Document Lifecycle**: Create, open, save, close, and manage multiple documents
- **Content Operations**: Add/edit paragraphs, headings, lists, and formatting
- **Table Operations**: Create, edit, style, and format tables
- **Page Layout**: Margins, page size, orientation, sections
- **Headers & Footers**: Add/remove headers, footers, page numbers, watermarks
- **Images**: Insert images from file or base64
- **Search & Replace**: Find, replace, batch replace, redact
- **Metadata**: Read/write document properties
- **Comments & Annotations**: Add comments and footnotes
- **Track Changes**: Enable/disable, accept/reject changes
- **Protection**: Protect and unprotect documents
- **Mail Merge**: Create merge fields and execute mail merges
- **Statistics**: Word count, document structure, outline

## Installation

```bash
# Install from source
pip install -e /home/sirobo/Documents/mcp/mcpdocx-unified

# Or install dependencies manually
pip install mcp python-docx lxml
```

## MCP Configuration

Add to your MCP configuration file (e.g., `~/.config/Code/User/mcp.json`):

```json
{
  "mcpServers": {
    "mcpdocx": {
      "command": "python",
      "args": ["-m", "mcpdocx"],
      "cwd": "/home/sirobo/Documents/mcp/mcpdocx-unified"
    }
  }
}
```

Or with uv:

```json
{
  "mcpServers": {
    "mcpdocx": {
      "command": "uv",
      "args": ["run", "--directory", "/home/sirobo/Documents/mcp/mcpdocx-unified", "mcpdocx"]
    }
  }
}
```

## 90 Tools Reference

### Document Lifecycle (8 tools)
| Tool | Description |
|------|-------------|
| `create_document` | Create a new Word document |
| `open_document` | Open an existing document |
| `save_document` | Save the current document |
| `close_document` | Close without saving |
| `get_document_info` | Get document metadata |
| `get_document_text` | Get all text content |
| `list_open_documents` | List all open documents |
| `switch_document` | Switch between documents |

### Paragraph & Content (12 tools)
| Tool | Description |
|------|-------------|
| `add_paragraph` | Add formatted paragraph |
| `add_heading` | Add heading (level 0-9) |
| `add_page_break` | Insert page break |
| `get_paragraph_text` | Get paragraph by index |
| `set_paragraph_text` | Update paragraph text |
| `delete_paragraph` | Remove paragraph |
| `insert_paragraph_after` | Insert after index |
| `count_paragraphs` | Count total paragraphs |
| `add_bulleted_list` | Add bullet list |
| `add_numbered_list` | Add numbered list |
| `apply_style` | Apply style to paragraph |
| `batch_format_paragraphs` | Style range of paragraphs |

### Search & Replace (5 tools)
| Tool | Description |
|------|-------------|
| `search_text` | Search in document |
| `find_and_replace` | Find and replace text |
| `batch_replace` | Replace multiple strings |
| `redact_text` | Redact sensitive text |
| `sanitize_external_links` | Remove external links |

### Table Operations (14 tools)
| Tool | Description |
|------|-------------|
| `add_table` | Create new table |
| `get_tables_info` | Get all tables info |
| `get_table_data` | Get table content |
| `edit_table_cell` | Edit cell content |
| `add_table_row` | Add row to table |
| `delete_table_row` | Remove row |
| `merge_table_cells` | Merge cells |
| `set_table_cell_shading` | Cell background color |
| `apply_table_alternating_rows` | Zebra striping |
| `highlight_table_header` | Style header row |
| `set_column_width` | Set column width |
| `set_row_height` | Set row height |
| `set_table_borders` | Configure borders |
| `import_csv_as_table` | Import CSV as table |

### Header & Footer (8 tools)
| Tool | Description |
|------|-------------|
| `add_header` | Add/update header |
| `add_footer` | Add/update footer |
| `add_page_numbers` | Insert page numbers |
| `get_header_text` | Get header content |
| `get_footer_text` | Get footer content |
| `remove_header` | Remove header |
| `remove_footer` | Remove footer |
| `add_text_watermark` | Add text watermark |

### Image Operations (4 tools)
| Tool | Description |
|------|-------------|
| `add_image` | Insert image from file |
| `add_image_base64` | Insert from base64 |
| `list_images` | List all images |
| `add_image_watermark` | Image watermark |

### Page Layout (8 tools)
| Tool | Description |
|------|-------------|
| `set_page_margins` | Set margins (cm) |
| `set_page_size` | Set page dimensions |
| `set_page_orientation` | Portrait/landscape |
| `get_page_info` | Get layout info |
| `add_section` | Add section break |
| `list_sections` | List all sections |
| `list_styles` | Available styles |
| `get_formatting_report` | Fonts & styles used |

### Metadata (3 tools)
| Tool | Description |
|------|-------------|
| `get_metadata` | Get properties |
| `set_metadata` | Set properties |
| `strip_personal_info` | Remove personal data |

### Comments & Annotations (4 tools)
| Tool | Description |
|------|-------------|
| `add_comment` | Add comment marker |
| `get_all_comments` | List all comments |
| `delete_all_comments` | Remove all comments |
| `add_footnote` | Add footnote |

### Hyperlinks & Bookmarks (4 tools)
| Tool | Description |
|------|-------------|
| `add_hyperlink` | External link |
| `add_bookmark` | Add bookmark |
| `list_bookmarks` | List bookmarks |
| `add_internal_link` | Link to bookmark |

### Document Merge (3 tools)
| Tool | Description |
|------|-------------|
| `merge_documents` | Merge multiple docs |
| `copy_document` | Copy document |
| `insert_document` | Insert document |

### TOC & Structure (4 tools)
| Tool | Description |
|------|-------------|
| `add_table_of_contents` | Add TOC |
| `get_document_outline` | Heading outline |
| `get_document_structure` | Structural overview |
| `get_raw_xml` | Raw XML (debug) |

### Track Changes & Protection (6 tools)
| Tool | Description |
|------|-------------|
| `enable_track_changes` | Enable tracking |
| `disable_track_changes` | Disable tracking |
| `accept_all_changes` | Accept all changes |
| `reject_all_changes` | Reject all changes |
| `protect_document` | Add protection |
| `unprotect_document` | Remove protection |

### Statistics (3 tools)
| Tool | Description |
|------|-------------|
| `get_word_count` | Word count stats |
| `get_document_statistics` | Full statistics |
| `count_tables` | Table count |

### File Operations (2 tools)
| Tool | Description |
|------|-------------|
| `list_docx_files` | List .docx in directory |
| `compare_documents` | Compare two documents |

### Mail Merge (2 tools)
| Tool | Description |
|------|-------------|
| `add_merge_field` | Add merge field |
| `execute_mail_merge` | Execute merge |

## Usage Examples

### Create a Document with Content

```
1. create_document("/path/to/doc.docx", "My Document Title")
2. add_heading("Introduction", 1)
3. add_paragraph("This is the first paragraph.", bold=true)
4. add_bulleted_list(["Item 1", "Item 2", "Item 3"])
5. add_table(3, 3, [["A", "B", "C"], ["1", "2", "3"], ["X", "Y", "Z"]])
6. save_document()
```

### Read and Analyze a Document

```
1. open_document("/path/to/existing.docx")
2. get_document_info()
3. get_document_outline()
4. get_word_count()
5. search_text("keyword")
```

### Edit an Existing Document

```
1. open_document("/path/to/doc.docx")
2. find_and_replace("old text", "new text")
3. set_paragraph_text(0, "Updated first paragraph")
4. add_header("Document Header")
5. save_document()
```

## Requirements

- Python 3.10+
- python-docx >= 1.1.0
- mcp >= 1.0.0
- lxml >= 4.9.0

## License

MIT License
