# MCP Documents Server - TIER 2: INTERMEDIATE

**Enhanced productivity features for formatting, structure, and conversions**

Total Tools: **113**

## Overview

Tier 2 extends Tier 1 with professional document features including headers/footers, images, advanced styling, comments, bookmarks, hyperlinks, and comprehensive format conversion capabilities. Ideal for business documents and content-rich automation.

## Tool Categories

### Core Operations (39 tools from Tier 1)
All Tier 1 features plus:

### Header & Footer Operations (9 tools)
- Add/edit headers and footers
- Different first page headers
- Section-specific headers
- Page numbering

### List Operations (5 tools)
- Create bullet lists
- Create numbered lists
- Multi-level lists
- List formatting

### Image Operations (4 tools)
- Insert images
- Resize and position images
- Image formatting
- Picture management

### Style Operations (8 tools)
- Apply character and paragraph styles
- Create custom styles
- Style management
- Advanced formatting

### Comment Operations (6 tools)
- Add comments
- Reply to comments
- Delete comments
- Comment management

### Bookmark Operations (4 tools)
- Create bookmarks
- Navigate to bookmarks
- Manage bookmark references

### Hyperlink Operations (2 tools)
- Add hyperlinks
- Manage links

### Advanced Table Operations (7 tools)
- Merge/split cells
- Table styles
- Advanced table formatting
- Cell borders and shading

### Table Border Operations (3 tools)
- Custom border styles
- Border colors
- Border width control

### Table of Contents (3 tools)
- Generate TOC
- Update TOC
- Custom TOC styles

### Metadata Operations (4 tools)
- Set document properties
- Author, title, subject
- Custom properties
- Document statistics

### Conversion Operations (16 tools)
- PDF conversion (with options)
- Markdown conversion
- HTML conversion
- Text extraction
- DOCX ↔ other formats
- Image conversion
- Multiple format support

### Section Operations (3 inline tools)
- Add section breaks
- List sections
- Set page orientation, margins, size

## Use Cases

- ✅ Business documents with headers/footers
- ✅ Reports with images and tables of contents
- ✅ Documents with cross-references and hyperlinks
- ✅ Styled corporate templates
- ✅ Comprehensive format conversions
- ✅ Review workflows (with comments)
- ✅ Multi-section documents
- ✅ Professional presentations converted to Word

## Not Included (See Tier 3)

- ❌ Track changes (revisions)
- ❌ Document protection
- ❌ Mail merge
- ❌ Watermarks
- ❌ Form controls
- ❌ Digital signatures and security
- ❌ Batch operations
- ❌ Collaboration features

## Running the Server

```bash
cd mcpdocxAdd1
python -m mcp_documents.server
```

## Configuration Example

```json
{
  "mcpServers": {
    "mcp-documents-intermediate": {
      "command": "python",
      "args": ["-m", "mcp_documents.server"],
      "cwd": "/path/to/mcpdocxAdd1"
    }
  }
}
```
