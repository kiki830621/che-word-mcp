# che-word-mcp

A Swift-native MCP (Model Context Protocol) server for Microsoft Word document (.docx) manipulation. This is the **first Swift OOXML library** that directly manipulates Office Open XML without any third-party Word dependencies.

[中文說明](README_zh-TW.md)

## Features

- **Pure Swift Implementation**: No Node.js, Python, or external runtime required
- **Direct OOXML Manipulation**: Works directly with XML, no Microsoft Word installation needed
- **Single Binary**: Just one executable file
- **18 MCP Tools**: Comprehensive document manipulation capabilities
- **Cross-platform**: Works on macOS (and potentially other platforms supporting Swift)

## Installation

### Prerequisites

- macOS 13.0+ (Ventura or later)
- Swift 5.9+

### Build from Source

```bash
git clone https://github.com/kiki830621/che-word-mcp.git
cd che-word-mcp
swift build -c release
```

The binary will be located at `.build/release/CheWordMCP`

### Add to Claude Code

```bash
claude mcp add che-word-mcp /path/to/che-word-mcp/.build/release/CheWordMCP
```

### Add to Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "che-word-mcp": {
      "command": "/path/to/che-word-mcp/.build/release/CheWordMCP"
    }
  }
}
```

## Available Tools

### Document Management (6 tools)

| Tool | Description |
|------|-------------|
| `create_document` | Create a new Word document |
| `open_document` | Open an existing .docx file |
| `save_document` | Save document to .docx file |
| `close_document` | Close an open document |
| `list_open_documents` | List all open documents |
| `get_document_info` | Get document statistics (paragraphs, words, characters) |

### Content Operations (6 tools)

| Tool | Description |
|------|-------------|
| `get_text` | Get plain text content |
| `get_paragraphs` | Get all paragraphs with formatting info |
| `insert_paragraph` | Insert a new paragraph |
| `update_paragraph` | Update existing paragraph content |
| `delete_paragraph` | Delete a paragraph |
| `replace_text` | Search and replace text |

### Formatting (3 tools)

| Tool | Description |
|------|-------------|
| `format_text` | Apply text formatting (bold, italic, color, font) |
| `set_paragraph_format` | Set paragraph formatting (alignment, spacing) |
| `apply_style` | Apply built-in styles (Heading1, Title, etc.) |

### Tables (1 tool)

| Tool | Description |
|------|-------------|
| `insert_table` | Insert a table with optional data |

### Export (2 tools)

| Tool | Description |
|------|-------------|
| `export_text` | Export document as plain text |
| `export_markdown` | Export document as Markdown |

## Usage Examples

### Create a Document with Headings and Text

```
Create a new Word document called "report" with:
- Title: "Quarterly Report"
- Heading: "Introduction"
- A paragraph explaining the report purpose
Save it to ~/Documents/report.docx
```

### Open and Modify Existing Document

```
Open the document at ~/Documents/proposal.docx
Replace all occurrences of "2024" with "2025"
Save the changes
```

### Create a Document with Table

```
Create a document with a 3x4 table containing:
- Header row: Name, Age, Department
- Data rows with employee information
Export it as Markdown too
```

## Technical Details

### OOXML Structure

The server generates valid Office Open XML documents with the following structure:

```
document.docx (ZIP)
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── word/
│   ├── document.xml      # Main content
│   ├── styles.xml        # Style definitions
│   ├── settings.xml      # Document settings
│   ├── fontTable.xml     # Font definitions
│   └── _rels/
│       └── document.xml.rels
└── docProps/
    ├── core.xml          # Metadata (author, title)
    └── app.xml           # Application info
```

### Dependencies

- [MCP Swift SDK](https://github.com/modelcontextprotocol/swift-sdk) (v0.10.0+) - Model Context Protocol implementation
- [ZIPFoundation](https://github.com/weichsel/ZIPFoundation) (v0.9.0+) - ZIP compression/decompression

## Comparison with Other Solutions

| Feature | Anthropic Word MCP | python-docx | docx npm | **che-word-mcp** |
|---------|-------------------|-------------|----------|------------------|
| Language | Node.js | Python | Node.js | **Swift** |
| Backend | AppleScript | OOXML | OOXML | **OOXML** |
| Requires Word | Yes | No | No | **No** |
| Runtime | Node.js | Python | Node.js | **None** |
| Single Binary | No | No | No | **Yes** |

## Roadmap

- [ ] Image support
- [ ] Header/Footer support
- [ ] Page breaks and sections
- [ ] Numbered/Bulleted lists
- [ ] Comments and track changes
- [ ] Read existing document styles

## License

MIT License

## Author

Che Cheng ([@kiki830621](https://github.com/kiki830621))

## Related Projects

- [che-apple-mail-mcp](https://github.com/kiki830621/che-apple-mail-mcp) - Apple Mail MCP server
- [che-ical-mcp](https://github.com/kiki830621/che-ical-mcp) - macOS Calendar MCP server
