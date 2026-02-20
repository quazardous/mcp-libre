# LibreOffice MCP Server

MCP server that lets AI assistants create, read, convert and edit LibreOffice documents.

[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![LibreOffice](https://img.shields.io/badge/LibreOffice-24.2+-green.svg)](https://www.libreoffice.org/)
[![MCP Protocol](https://img.shields.io/badge/MCP-2024--11--05-orange.svg)](https://spec.modelcontextprotocol.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Requirements

- **LibreOffice** 24.2+ (accessible via command line)
- **Python** 3.12+
- **UV** package manager

## Installation

### Windows

```powershell
git clone https://github.com/patrup/mcp-libre/
cd mcp-libre
.\setup-windows.ps1
```

The script installs all dependencies via winget, adds LibreOffice to PATH, runs `uv sync`, and configures Claude Desktop automatically.

Options: `-CheckOnly` (dry run), `-SkipOptional` (skip Node.js/Java).

### Linux / macOS

```bash
git clone https://github.com/patrup/mcp-libre/
cd mcp-libre
uv sync
./mcp-helper.sh check   # verify setup
```

## Quick Start

```bash
# Test the server
uv run python src/main.py --test

# Run as MCP server (stdio)
uv run python src/main.py
```

## Tools

| Tool | Description |
|------|-------------|
| `create_document` | Create Writer, Calc, Impress, Draw documents |
| `read_document_text` | Extract text content |
| `convert_document` | Convert between 50+ formats (PDF, DOCX, HTML...) |
| `get_document_info` | Get document metadata |
| `read_spreadsheet_data` | Read spreadsheet data as 2D arrays |
| `insert_text_at_position` | Edit document text at a given position |
| `search_documents` | Find documents by content |
| `batch_convert_documents` | Batch format conversion |
| `merge_text_documents` | Merge multiple documents |
| `get_document_statistics` | Word count, sentences, etc. |
| `open_document_in_libreoffice` | Open in GUI for live viewing |
| `create_live_editing_session` | Live editing with real-time preview |
| `watch_document_changes` | Monitor document changes |
| `refresh_document_in_libreoffice` | Force refresh in GUI |

## Integration

### Claude Desktop

**Windows** (done automatically by `setup-windows.ps1`):

Config at `%APPDATA%\Claude\claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "libreoffice": {
      "command": "uv",
      "args": ["run", "python", "src/main.py"],
      "cwd": "C:/Users/you/mcp-libre",
      "env": {
        "PYTHONPATH": "C:/Users/you/mcp-libre/src",
        "MCP_LIBREOFFICE_GUI": "0"
      }
    }
  }
}
```

**Linux / macOS**:
```bash
./generate-config.sh claude
```

### LibreOffice Extension (optional, 10x faster)

```bash
cd plugin/
./install.sh install   # build & install
./install.sh test      # verify
```

Provides an HTTP API on `localhost:8765` with direct UNO API access. See [plugin/README.md](plugin/README.md).

### Super Assistant Chrome Extension

```bash
./generate-config.sh mcp
npx @srbhptl39/mcp-superassistant-proxy@latest --config ~/Documents/mcp/mcp.config.json
```

## Documentation

See the [docs/](docs/) folder: [Windows Setup](docs/WINDOWS_SETUP.md) | [Prerequisites](docs/PREREQUISITES.md) | [Examples](docs/EXAMPLES.md) | [Troubleshooting](docs/TROUBLESHOOTING.md) | [Plugin Guide](docs/PLUGIN_MIGRATION_GUIDE.md) | [Live Viewing](docs/LIVE_VIEWING_GUIDE.md)

## License

[MIT](LICENSE)
