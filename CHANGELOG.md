# Changelog

## [2.0.1] - 2026-02-21

### Fixed
- About dialog and extension metadata updated (name, version, description)
- Build script: `.oxt` packaging on Windows (Compress-Archive .zip rename)

## [2.0.0] - 2026-02-21

### Changed
- **Full architecture rewrite** — cleaner codebase, easier to maintain and extend
- Each tool is now a self-contained module (was two giant files totaling 6000+ lines)
- Adding a new tool = adding one small file, no other changes needed

### Added
- **Backpressure** — concurrent tool calls now queue properly instead of freezing the UI
- `insert_paragraphs_batch` — insert multiple paragraphs in one operation
- `style` parameter on `insert_text_at_paragraph`

### Removed
- Old standalone server (`src/`, `libremcp.py`) — everything runs inside the LibreOffice extension now
- Legacy test files

## [1.4.0] - 2026-02-20

### Added
- `insert_image` — insert image with optional caption, anchored at a paragraph
- `delete_image` — remove image (and its frame)
- `replace_image` — swap an image's source file, keeping position and caption
- `goto_page` — scroll LibreOffice to a specific page
- `get_page_objects` now returns frames with their contained images

### Changed
- Renamed project to "LibreOffice MCP"

### Fixed
- Frame-image detection on certain documents

## [1.3.1] - 2026-02-20

### Added
- **Page proximity** — find images and tables near a paragraph
- `get_document_tree` now includes page numbers and total page count
- `close_document` tool
- `search_documents` searches open documents by default (faster, no side effects)

### Changed
- `open_document` reuses already-open documents instead of erroring

### Fixed
- Document path matching on Windows

## [1.3.0] - 2026-02-20

### Added
- **Main thread execution** — all operations run on LibreOffice's main thread, fixing UI freezes and crashes
- **59 tools** (was 24): comments, track changes, styles, tables, images, paragraph editing, document management, protection, recent documents
- Agent guide (`AGENT.md`) and developer guide (`DEVEL.md`)

### Changed
- Cross-platform install scripts (Linux + Windows)
- Single source of truth for version number

### Fixed
- Image property serialization
- Recent documents path on some LibreOffice versions

## [1.2.0] - 2026-02-20

### Added
- All tools work in GUI mode via the plugin HTTP API

### Fixed
- "LibreOffice is already running" error in GUI mode

## [1.1.3] - 2026-02-20

### Added
- Dynamic menu icons (play/stop) based on server state
- Live-updating Server Status dialog
- Dev hot-deploy workflow
- Settings dialog in LibreOffice Options
- Auto-restart on config change

### Fixed
- Menu graying out during server startup
- Slow Status dialog

## [1.0.0] - 2025-06-28

### Added
- LibreOffice extension with embedded MCP server
- Document navigation (heading tree, paragraphs, search, bookmarks)
- AI annotation system
- Live viewing tools

## [0.1.0] - 2025-06-27

### Added
- Initial MCP server
- Writer, Calc, Impress, Draw support
- Document conversion, search, batch operations
