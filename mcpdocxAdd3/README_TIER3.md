# MCP Documents Server - TIER 3: ADVANCED

**Enterprise features: security, collaboration, track changes, mail merge, batch operations**

Total Tools: **72**

## Overview

Tier 3 provides enterprise-grade document features including track changes with real revision markup, document protection, digital signatures, mail merge, form controls, watermarks, batch operations, Excel integration, advanced analysis, and collaborative editing sessions.

## Tool Categories

### Track Changes Operations (7 tools)
- Enable/disable track changes
- Accept/reject revisions
- Add revision inserts (w:ins)
- Add revision deletes (w:del)
- Real OOXML revision markup
- Author tracking

### Protection Operations (4 tools)
- Protect document with password
- Unprotect document
- Add editable ranges (permStart/permEnd)
- Range-based permissions

### Mail Merge Operations (3 tools)
- Insert MERGEFIELD placeholders
- Execute mail merge with data
- Advanced IF/COMPARE field evaluation

### Watermark Operations (2 tools)
- Add text watermarks
- Add image watermarks (anchored behind text)

### Form Operations (4 tools)
- Add checkbox form controls (sdt)
- Add dropdown form controls (sdt)
- Add date picker form controls (sdt)
- Content control management

### Security Operations (8 tools)
- Redact sensitive content
- Sanitize metadata
- Enable filesystem sandbox
- Disable sandbox
- Set readonly mode
- Sign document (detached SHA-256/HMAC)
- Verify document signature
- Security enforcement

### Document Merge Operations (3 tools)
- Merge multiple documents
- Combine with section breaks
- Preserve formatting

### Excel Operations (9 tools)
- Read Excel data
- Write Excel files
- Convert Excel to DOCX tables
- Export tables to Excel
- Cell formatting
- Excel integration

### Batch Operations (3 tools)
- Batch process multiple documents
- Apply operations to folders
- Automated workflows

### Footnote Operations (2 tools)
- Add footnotes
- Manage footnote references

### Endnote Operations (4 tools)
- Add endnotes
- Manage endnote references
- Endnote styling

### Extended Table Operations (7 tools)
- Advanced table manipulation
- Complex cell operations
- Table calculations
- Data validation

### Section Operations (7 tools)
- Add section breaks
- Multi-column page layout (w:cols)
- Section properties (full)
- Page orientation per section
- Different first page settings
- Section-specific formatting

### Analysis Operations (5 tools)
- Document outline analysis
- Statistics and reporting
- Formatting analysis report
- Lint document (whitespace, repeats, optional grammar)
- Quality checks

### Collaboration Operations (4 tools)
- Start collaboration session
- Apply collaboration changes
- Get collaboration log
- In-memory change tracking

## Advanced Features Highlights

### 🔒 Security & Protection
- Detached digital signatures (SHA-256/HMAC sidecar JSON)
- Filesystem sandbox with path validation
- Readonly mode enforcement
- Document password protection
- Sensitive content redaction

### 📝 Track Changes & Revisions
- Real OOXML revision markup (w:ins, w:del)
- Author attribution
- Timestamp tracking
- Accept/reject workflows

### 📋 Form Controls
- Content controls (sdt elements)
- Checkbox, dropdown, date picker
- Form-based workflows

### 🔄 Mail Merge
- MERGEFIELD insertion
- IF conditional logic evaluation
- COMPARE field evaluation
- Data-driven document generation

### 📄 Multi-Column Layout
- w:cols OOXML support
- Column count, spacing, separator
- Section-specific layouts

### ⚙️ Batch & Automation
- Process multiple documents
- Excel data integration
- Automated workflows
- Enterprise-scale operations

## Use Cases

- ✅ Enterprise document workflows
- ✅ Legal documents with track changes
- ✅ Secure documents with signatures
- ✅ Mail merge campaigns
- ✅ Form-based data collection
- ✅ Batch document processing
- ✅ Excel-to-Word reporting
- ✅ Collaborative editing sessions
- ✅ Multi-column newsletters
- ✅ Protected templates with editable ranges
- ✅ Quality assurance with linting
- ✅ Watermarked confidential documents

## Dependencies

All core dependencies from Tier 1/2, plus optional:
- `language-tool-python` - For grammar checking in linting
- `openpyxl` or `pandas` - For Excel operations

## Running the Server

```bash
cd mcpdocxAdd2
python -m mcp_documents.server
```

## Configuration Example

```json
{
  "mcpServers": {
    "mcp-documents-advanced": {
      "command": "python",
      "args": ["-m", "mcp_documents.server"],
      "cwd": "/path/to/mcpdocxAdd2"
    }
  }
}
```

## Security Notes

- Sandbox mode restricts file access to allowed directories
- Digital signatures use HMAC for integrity verification
- Readonly mode prevents document modifications
- All security features are optional and configurable
