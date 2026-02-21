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

## LibreOffice Extension

The extension embeds an MCP server inside LibreOffice (port 8765) with direct UNO API access. All document operations run on LibreOffice's main thread for full fidelity.

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

## Tools (67+)

| Category | Examples |
|----------|---------|
| **Document** | `create_document`, `read_document_text`, `convert_document`, `open_document_in_libreoffice` |
| **Navigation** | `get_document_tree` (heading tree + page numbers), `search_in_document`, `get_page_objects` |
| **Editing** | `insert_text_at_paragraph`, `set_paragraph_text`, `replace_in_document` |
| **Comments & review** | `list_comments`, `add_comment`, `resolve_comment`, track changes |
| **Images & frames** | `insert_image`, `set_image_properties` (resize, crop, alt-text), `replace_image` |
| **Tables** | `list_tables`, `read_table`, `write_table_cell`, `create_table` |
| **Styles** | `list_styles`, `get_style_info` |
| **Calc** | `read_spreadsheet_cells`, `write_spreadsheet_cell`, `list_sheets` |
| **Impress** | `list_slides`, `read_slide`, `get_presentation_info` |
| **Batch** | `batch_convert_documents`, `merge_text_documents` |

## Configuration

The extension runs an HTTP server on port 8765 inside LibreOffice. Configure your MCP client to connect to it:

- **Claude Desktop**: `~/.config/claude/claude_desktop_config.json` (Linux) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows)
- **Claude Code**: `.mcp.json` at the project root

```json
{
  "mcpServers": {
    "libreoffice": {
      "type": "http",
      "url": "http://localhost:8765/mcp"
    }
  }
}
```

See [config/claude_code.json.template](config/claude_code.json.template) and [config/claude_desktop.json.template](config/claude_desktop.json.template) for ready-to-use templates.

## Documentation

- [Prerequisites](docs/prerequisites.md)
- [Windows Setup](docs/windows-setup.md)
- [Troubleshooting](docs/troubleshooting.md)
- [Live Viewing](docs/live-viewing.md)
- [Changelog](CHANGELOG.md)

## Copyright

This project is a major fork of [patrup/mcp-libre](https://github.com/patrup/mcp-libre), which provided the initial MCP server and headless LibreOffice integration. The current version has been extensively rewritten: embedded LibreOffice extension with UNO bridge, main-thread execution via AsyncCallback, context-efficient document navigation (heading tree, locators, bookmarks), comments/review workflow, track changes, styles, tables, images, document protection, and 67+ tools.

## License

[MIT](LICENSE)
