# Changelog

<!--
  Note to contributors: keep entries human-scoped and simple.
  One line per change, plain language, no internal jargon.
  Group by Added / Changed / Fixed.
-->

## [2.3.0] - 2026-02-22

### Added
- **ChatGPT Desktop support** — MCP server now works with ChatGPT via Tailscale Funnel or cloudflared tunnel
- **Tailscale Funnel** as default tunnel provider (stable HTTPS URL, no random subdomains)
- **Image URL support** — `insert_image` and `replace_image` accept HTTP/HTTPS URLs (auto-download)
- **`clone_heading_block`** — duplicate an entire heading with all its sub-headings and body content
- **`document_health_check`** — diagnostic tool: empty headings, broken bookmarks, orphan images, heading level skips
- **`scan_tasks`** — scan comments for actionable prefixes (TODO-AI, FIX, QUESTION, VALIDATION, NOTE) to find tasks without reading the document body
- **Workflow dashboard** — `get_workflow_status` / `set_workflow_status` use a master comment (author: MCP-WORKFLOW) as a shared project dashboard for multi-agent collaboration
- **Local proximity navigation** — `navigate_heading` moves between headings locally (next, previous, parent, first_child, siblings) without rescanning the full document
- **`get_surroundings`** — discover images, tables, frames, comments, and paragraph context within a radius of any locator (one-shot spatial awareness for AI agents)
- **Multi-agent comments** — `list_comments` gains `author_filter` param to filter by AI agent (ChatGPT, Claude, etc.)
- **Tunnel settings** split into a dedicated dialog tab with per-provider fields (show/hide based on selection)
- **Copy MCP URL** button in tunnel and status dialogs (clipboard via LO API)
- `/sse` and `/messages` endpoints for ChatGPT SSE transport compatibility
- `Mcp-Session-Id` header on all MCP responses
- `resources/list` and `prompts/list` handlers (return empty, required by spec)

### Changed
- **Protocol version** bumped to `2025-11-25` (echoes client's version for compatibility)
- **`initialize` response** reduced from 20KB to ~400 bytes (was embedding full AGENT.md, now a short instruction string)
- Tool descriptions rewritten to guide AI toward bookmark-based navigation (ESSENTIAL FIRST CALL, never scan manually)
- `add_comment` description encourages agents to use their own name as author
- Tunnel provider list reordered: tailscale first, then cloudflared, bore, ngrok
- Log file now resets on each LibreOffice restart (`mode="w"`)

### Fixed
- **No more viewport jumping** — page number lookups, document tree, page objects scan all use `lockControllers` + cursor save/restore (user's view stays where they left it)
- Bookmark resolution no longer enumerates the document twice (reuses paragraph range list)
- Tailscale Funnel "listener already exists" error — auto-reset before start and on stop
- Options dialog crash when detecting tunnel vs server page (`getControl` returns `None`, not exception)
- Batch JSON-RPC requests (array) now handled correctly on `/mcp` and `/sse`

## [2.2.0] - 2026-02-22

### Added
- Multi-provider tunnel support (cloudflared, bore, ngrok, tailscale)
- Cloudflared named tunnel support (stable URL via Cloudflare account)
- Bore tunnel integration

## [2.1.2] - 2026-02-21

### Added
- HTTPS/HTTP toggle in the menu bar (MCP Server > HTTPS: On/Off)
- Dynamic menu text updates to reflect current SSL state

### Fixed
- `tools/list` crash (`'McpTool' object has no attribute 'get'`) — used `getattr()` instead of `.get()` on tool objects
- `description.xml` broken namespace (`dep:name` undeclared) — use proper `xmlns:l` / `xmlns:d` namespaces
- Build scripts now copy real `description.xml` from `plugin/` instead of generating a hardcoded one
- Build scripts patch version from `version.py` (single source of truth) into `description.xml` at build time
- Linux build script: added `Jobs.xcu` to the XCU copy loop
- Publisher URL in `description.xml` now points to the actual GitHub repo
- Removed stale Python bridge config from `claude_desktop_config.json` (now uses HTTP direct like `.mcp.json`)

## [2.1.1] - 2026-02-21

### Fixed
- Auto-start not triggering after fresh install (`Jobs.xcu` missing from build manifest)
- Module-level auto-start fallback for reliable server startup

## [2.1.0] - 2026-02-21

### Added
- **HTTPS by default** - the MCP server now uses TLS with an auto-generated self-signed certificate (zero config needed)
- `EnableSSL` toggle in LibreOffice Options (Tools > Options > MCP Server)

### Changed
- All URLs updated from `http://` to `https://`

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
