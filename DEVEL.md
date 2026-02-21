# DEVEL

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/install-plugin.ps1` | Build + install the .oxt extension via unopkg |
| `scripts/install-plugin.ps1 -BuildOnly` | Build the .oxt only (no install) |
| `scripts/install-plugin.sh` | Build + install (Linux/macOS) |
| `scripts/dev-deploy.ps1` | Dev mode: sync plugin/ → build/dev/ + junction in LO |
| `scripts/dev-deploy.sh` | Dev mode (Linux/macOS) |
| `scripts/kill-libreoffice.ps1` | Kill all soffice processes |
| `scripts/launch-lo-debug.ps1` | Launch LO with SAL_LOG for debugging |

## Hot deploy

```powershell
.\scripts\dev-deploy.ps1     # sync plugin/ → build/dev/ + junction into LO extensions
# restart LibreOffice
curl -k https://localhost:8765/health
```

## Logs

- **Plugin (LO side):** `~/mcp-extension.log`
- **LO internal debug** (slow startup!): `.\scripts\launch-lo-debug.ps1 -Full`

## Architecture

```
plugin/               → LO extension (.oxt), runs INSIDE LibreOffice
  pythonpath/
    registration.py   → UNO component, menu, HTTP server lifecycle
    ai_interface.py   → MCP HTTP server (port 8765), backpressure
    main_thread_executor.py → VCL main-thread dispatch
    mcp_server.py     → auto-discover tools, dispatch execute

    services/
      __init__.py     → ServiceRegistry (container for all services)
      base.py         → BaseService (desktop, ctx, resolve_doc, resolve_locator)
      writer/
        __init__.py   → WriterService facade
        tree.py       → TreeService (heading tree, caching)
        paragraphs.py → ParagraphService (read, edit, insert)
        search.py     → SearchService (search, replace)
        structural.py → StructuralService (sections, bookmarks, pages)
      calc.py         → CalcService (cells, sheets, ranges)
      impress.py      → ImpressService (slides, presentations)
      images.py       → ImageService (images + text frames)
      comments.py     → CommentService (comments, AI annotations, track changes)
      tables.py       → TableService (writer tables)
      styles.py       → StyleService (style listing, introspection)

    tools/            → One class per MCP tool (auto-discovered)
      __init__.py     → discover_tools() — auto-collects McpTool subclasses
      base.py         → McpTool ABC (name, description, parameters, execute)
      navigation.py   → GetDocumentTree, GetHeadingChildren, ReadParagraphs, ...
      editing.py      → InsertAtParagraph, InsertBatch, DeleteParagraph, ...
      search.py       → SearchInDocument, ReplaceInDocument
      document.py     → CreateDocument, OpenDocument, SaveDocument, ...
      structural.py   → ListSections, ReadSection, ListBookmarks, PageCount, ...
      annotations.py  → AddAiSummary, GetAiSummaries, RemoveAiSummary
      comments.py     → ListComments, AddComment, ResolveComment, DeleteComment
      tracking.py     → SetTrackChanges, GetTrackedChanges, AcceptAll, RejectAll
      images.py       → ListImages, GetImageInfo, InsertImage, DeleteImage, ...
      frames.py       → ListFrames, GetFrameInfo, SetFrameProperties
      tables.py       → ListTables, ReadTable, WriteTableCell, CreateTable
      styles.py       → ListStyles, GetStyleInfo
      calc.py         → ReadCells, WriteCell, ListSheets, GetSheetInfo
      impress.py      → ListSlides, ReadSlideText, GetPresentationInfo
      protection.py   → SetDocumentProtection
      metadata.py     → GetDocumentProperties, SetDocumentProperties
```

**Flow:** Claude → MCP streamable-http POST → ai_interface.py → backpressure → VCL main thread → tool.execute() → service → UNO API

## Known pitfalls

- **NEVER nuke `uno_packages/cache/`** — use `unopkg remove` + `unopkg add`
- **`allow_reuse_address = True`** on Windows = port shadowing. Forbidden.
- **PowerShell UTF-8 BOM**: `Set-Content -Encoding UTF8` adds BOM → LO Python crashes. Use `System.Text.UTF8Encoding($false)`.
- **Relative imports** in plugin: `from .module import X` → `from module import X` (LO adds `pythonpath/` to sys.path directly). Build scripts convert automatically.
- **`registration.py` loaded twice** by LO: once for UNO registration, once at runtime.

## Locator convention (`type:value`)

Writer: `bookmark:_mcp_x`, `paragraph:42`, `page:3`, `section:Intro`, `heading:2.1`
Calc: direct params (`range`, `cell`, `sheet_name`)
Impress: direct params (`slide_index`)

## Adding a new tool

1. Create a class in `tools/<domain>.py` extending `McpTool`
2. Set `name`, `description`, `parameters` (JSON Schema)
3. Implement `execute(**kwargs)` — delegate to the appropriate service
4. Done. Auto-discovery picks it up on next server restart.
