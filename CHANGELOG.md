# Changelog

## [1.2.0] - 2026-02-20

### Added
- All MCP tools now work in GUI mode via plugin HTTP API
- Tools `create_document`, `read_document_text`, `convert_document`, `read_spreadsheet_data`, `insert_text_at_position`, `open_document_in_libreoffice`, `refresh_document_in_libreoffice` delegate to plugin when `MCP_LIBREOFFICE_GUI=1`

### Fixed
- "LibreOffice is already running" error when using MCP tools in GUI mode

## [1.1.3] - 2026-02-20

### Added
- Dynamic menu icons (play/stop/hourglass) based on server state
- Live-updating Server Status dialog with async health probe
- Dev workflow with junction symlink (`dev-deploy.ps1`)
- Version displayed in logs and About dialog
- Native LibreOffice Options dialog for config
- Auto-restart on config change

### Fixed
- Menu graying out during server startup
- Icons not showing on first menu display
- Slow Server Status dialog (now async)

## [1.0.0] - 2025-06-28

### Added
- LibreOffice plugin/extension with embedded HTTP server
- Context-efficient document tools (heading tree, paragraphs, search, bookmarks)
- AI annotation system for document headings
- Live viewing and document management tools

## [0.1.0] - 2025-06-27

### Added
- Initial MCP server with document CRUD operations
- Writer, Calc, Impress, Draw support
- Document conversion, search, batch operations
- Configuration generator for Claude Desktop
