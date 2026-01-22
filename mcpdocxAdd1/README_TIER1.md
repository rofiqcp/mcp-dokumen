# MCP Documents Server - TIER 1: BASIC

**Essential document operations for basic CRUD and simple formatting**

Total Tools: **44**

## Overview

Tier 1 provides core functionality for creating, editing, and managing Word documents with basic formatting and simple conversions. Perfect for lightweight document automation and simple use cases.

## Tool Categories

### Document Operations (9 tools)
- Create, open, save documents
- Get document information
- Basic document management

### Content Operations (10 tools)
- Add text, paragraphs
- Insert page breaks
- Find and replace text
- Basic content manipulation

### Table Operations (8 tools)
- Create tables
- Add rows and columns
- Set cell values
- Basic table formatting

### Paragraph Operations (6 tools)
- Format paragraphs
- Set alignment, spacing, indentation
- Basic paragraph styling

### Structure Operations (6 tools)
- Manage document structure
- Add sections
- Basic navigation

### Style & Conversion (5 inline tools)
- Apply paragraph styles
- List available styles
- Convert to PDF
- Convert to Markdown
- Convert to HTML

## Use Cases

- ✅ Simple document generation
- ✅ Basic text editing and formatting
- ✅ Table creation and manipulation
- ✅ PDF/Markdown/HTML conversion
- ✅ Template-based document creation
- ✅ Basic automation scripts

## Not Included (See Tier 2 & 3)

- ❌ Headers/Footers
- ❌ Images and graphics
- ❌ Comments and track changes
- ❌ Advanced formatting (watermarks, bookmarks)
- ❌ Mail merge
- ❌ Security features
- ❌ Batch operations

## Installation

Same as main mcpdocx installation - this is just a reduced feature set using the same codebase.

## Running the Server

```bash
cd mcpdocx
python -m mcp_documents.server
```

## Configuration Example

```json
{
  "mcpServers": {
    "mcp-documents-basic": {
      "command": "python",
      "args": ["-m", "mcp_documents.server"],
      "cwd": "/path/to/mcpdocx"
    }
  }
}
```
