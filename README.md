# LibreOffice MCP

MCP server that lets AI assistants create, read, convert and edit LibreOffice documents in real-time.

[![LibreOffice](https://img.shields.io/badge/LibreOffice-24.2+-green.svg)](https://www.libreoffice.org/)
[![MCP Protocol](https://img.shields.io/badge/MCP-2024--11--05-orange.svg)](https://spec.modelcontextprotocol.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## How it works

A LibreOffice extension runs an MCP server directly inside LibreOffice (port 8765). Your AI assistant connects to it and can manipulate documents in real-time — you see changes live in the LibreOffice window.

## Quick start

1. **Install LibreOffice** 24.2+ if not already installed
2. **Install the extension** — download the `.oxt` file from [Releases](https://github.com/quazardous/mcp-libre/releases) and double-click it (or drag it into LibreOffice)
3. **Restart LibreOffice** — an "MCP Server" menu appears, the server starts automatically
4. **Configure your MCP client** — add this to your config:

**Claude Desktop** (`~/.config/claude/claude_desktop_config.json` or `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "libreoffice": {
      "type": "http",
      "url": "https://localhost:8765/mcp"
    }
  }
}
```

**Claude Code** (`.mcp.json` at the project root):

```json
{
  "mcpServers": {
    "libreoffice": {
      "type": "http",
      "url": "https://localhost:8765/mcp"
    }
  }
}
```

That's it. Open a document in LibreOffice, then ask Claude to edit it.

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

## Development

If you want to modify the extension or contribute, clone the repo and use the dev scripts:

```bash
git clone https://github.com/quazardous/mcp-libre/
cd mcp-libre
```

| Platform | Build & install | Dev hot-deploy |
|----------|----------------|----------------|
| **Linux** | `./scripts/install-plugin.sh` | `./scripts/dev-deploy.sh` |
| **Windows** | `.\scripts\install-plugin.ps1` | `.\scripts\dev-deploy.ps1` |

Dev-deploy syncs your changes to LibreOffice without rebuilding the `.oxt` — just restart LibreOffice to test.

See [DEVEL.md](DEVEL.md) for architecture, adding new tools, and known pitfalls.

## Documentation

- [Troubleshooting](docs/troubleshooting.md)
- [Live Viewing](docs/live-viewing.md)
- [Windows Setup](docs/windows-setup.md)
- [Changelog](CHANGELOG.md)

## Copyright

This project is a major fork of [patrup/mcp-libre](https://github.com/patrup/mcp-libre), which provided the initial MCP server and headless LibreOffice integration. The current version has been extensively rewritten with an embedded LibreOffice extension, 67+ tools, and real-time document editing.

## License

[MIT](LICENSE)
