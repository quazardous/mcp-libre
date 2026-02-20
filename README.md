# LibreOffice MCP

MCP server that lets AI assistants create, read, convert and edit LibreOffice documents in real-time.

[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![LibreOffice](https://img.shields.io/badge/LibreOffice-24.2+-green.svg)](https://www.libreoffice.org/)
[![MCP Protocol](https://img.shields.io/badge/MCP-2024--11--05-orange.svg)](https://spec.modelcontextprotocol.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Requirements

- **LibreOffice** 24.2+
- **Python** 3.12+
- **UV** package manager

See [docs/prerequisites.md](docs/prerequisites.md) for install commands per platform.

## Installation

```bash
git clone https://github.com/quazardous/mcp-libre/
cd mcp-libre
uv sync
```

The install scripts check/install all dependencies (Python 3.12+, LibreOffice, UV) and generate the Claude Desktop configuration automatically:

| Platform | Command | Details |
|----------|---------|---------|
| **Linux** | `./install.sh` | Supports apt, dnf, pacman, zypper, brew |
| **Windows** | `.\install.ps1` | See [docs/windows-setup.md](docs/windows-setup.md) |

Options: `--check-only` (status only), `--skip-optional` (skip Node.js/Java), `--plugin` (build + install extension).

## Modes

### Headless mode (default)

The MCP server launches LibreOffice in headless mode for each operation. No GUI needed.

```
MCP_LIBREOFFICE_GUI=0
```

### GUI mode (recommended)

The MCP server delegates to a LibreOffice extension running inside an open LibreOffice instance. Real-time, faster, and you see changes live.

```
MCP_LIBREOFFICE_GUI=1
MCP_PLUGIN_URL=http://localhost:8765
```

Requires the LibreOffice extension installed (see below).

## LibreOffice Extension

The extension embeds an HTTP API server inside LibreOffice (port 8765) with direct UNO API access.

### Install

| Platform | Command |
|----------|---------|
| **Linux** | `./install.sh --plugin` |
| **Windows** | `.\install.ps1 -Plugin` |

### Dev workflow

| Platform | Command |
|----------|---------|
| **Linux** | `./scripts/dev-deploy.sh` |
| **Windows** | `.\scripts\dev-deploy.ps1` |

Then restart LibreOffice to pick up changes.

The extension adds an **MCP Server** menu in LibreOffice with Start/Stop, Restart, Status, and About.

## Tools

### Document operations

| Tool | Description |
|------|-------------|
| `create_document` | Create Writer, Calc, Impress, Draw documents |
| `read_document_text` | Extract text content |
| `convert_document` | Convert between formats (PDF, DOCX, HTML...) |
| `get_document_info` | Get document metadata |
| `read_spreadsheet_data` | Read spreadsheet data as 2D arrays |
| `insert_text_at_position` | Insert text at start, end, or replace |
| `search_documents` | Find documents by content |
| `batch_convert_documents` | Batch format conversion |
| `merge_text_documents` | Merge multiple documents |
| `get_document_statistics` | Word count, sentences, etc. |

### Context-efficient navigation (via plugin)

| Tool | Description |
|------|-------------|
| `get_document_tree` | Heading tree without loading full text |
| `get_heading_children` | Drill down into a heading |
| `read_document_paragraphs` | Read paragraphs by index range |
| `get_document_paragraph_count` | Total paragraph count |
| `get_document_page_count` | Page count |
| `search_in_document` | Search with paragraph context |
| `replace_in_document` | Find & replace preserving formatting |
| `insert_text_at_paragraph` | Insert before/after a paragraph |

### Bookmarks, sections & annotations

| Tool | Description |
|------|-------------|
| `list_document_bookmarks` | List all bookmarks |
| `resolve_document_bookmark` | Resolve bookmark to paragraph index |
| `list_document_sections` | List named text sections |
| `read_document_section` | Read a section's content |
| `add_document_ai_summary` | Add AI annotation to a heading |
| `get_document_ai_summaries` | List all AI annotations |
| `remove_document_ai_summary` | Remove an AI annotation |

### Live viewing

| Tool | Description |
|------|-------------|
| `open_document_in_libreoffice` | Open in GUI |
| `refresh_document_in_libreoffice` | Force refresh in GUI |
| `create_live_editing_session` | Live editing session |
| `watch_document_changes` | Monitor file changes |

## Configuration

The install scripts generate the Claude Desktop configuration automatically. To configure manually, add the MCP server entry to your config:

- **Claude Desktop**: `~/.config/claude/claude_desktop_config.json` (Linux) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows)
- **Claude Code**: `.mcp.json` at the project root

**Linux:**

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "/home/you/.local/bin/uv",
      "args": ["run", "python", "src/main.py"],
      "cwd": "/home/you/mcp-libre",
      "env": {
        "PYTHONPATH": "/home/you/mcp-libre/src",
        "MCP_LIBREOFFICE_GUI": "1",
        "MCP_PLUGIN_URL": "http://localhost:8765"
      }
    }
  }
}
```

**Windows:**

```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "C:/Users/you/.local/bin/uv.exe",
      "args": ["run", "python", "src/main.py"],
      "cwd": "C:/Users/you/mcp-libre",
      "env": {
        "PYTHONPATH": "C:/Users/you/mcp-libre/src",
        "MCP_LIBREOFFICE_GUI": "1",
        "MCP_PLUGIN_URL": "http://localhost:8765"
      }
    }
  }
}
```

## Documentation

- [Prerequisites](docs/prerequisites.md)
- [Windows Setup](docs/windows-setup.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Live Viewing](docs/live-viewing.md)
- [Changelog](CHANGELOG.md)

## Copyright

This project is a major fork of [patrup/mcp-libre](https://github.com/patrup/mcp-libre), which provided the initial MCP server and headless LibreOffice integration. The current version has been extensively rewritten: embedded LibreOffice extension with UNO bridge, main-thread execution via AsyncCallback, context-efficient document navigation (heading tree, locators, bookmarks), comments/review workflow, track changes, styles, tables, images, document protection, and 59+ tools.

## License

[MIT](LICENSE)
