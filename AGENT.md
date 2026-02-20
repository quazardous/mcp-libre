# LibreOffice MCP — Agent Guide

This MCP server gives you live access to LibreOffice documents.
Follow these rules to keep your context window small and your calls fast.

## Core Principle: Drill Down, Never Dump

Documents can have thousands of paragraphs. Never try to load everything.
Navigate the heading tree like a filesystem: list, then open what you need.

## Workflow

### 1. Discover structure

```
get_document_tree(depth=1, content_strategy="first_lines")
```

Returns top-level headings with a one-line preview. This is your table of contents.
For a 800-paragraph document, this returns ~10 items.

### 2. Drill into a section

```
get_heading_children(locator="bookmark:_mcp_abc123", depth=1)
```

Each heading comes with a stable `bookmark` — use it for navigation.
Bookmarks survive document edits; paragraph indices don't.

### 3. Read only what you need

```
read_paragraphs(locator="paragraph:58", count=10)
```

Read small batches. Default is 10 paragraphs. Increase only if you have a reason.

### 4. Search instead of scanning

```
search_in_document(pattern="évaluation", max_results=5, context_paragraphs=1)
```

Find specific content without reading the whole document.
Returns matches with surrounding paragraphs for context.

## content_strategy

Controls how much body text is returned alongside headings:

| Value | What you get | When to use |
|---|---|---|
| `"none"` | Headings only | Navigating structure |
| `"first_lines"` | Headings + first line of body | Default, good balance |
| `"ai_summary_first"` | AI summaries if cached, else first lines | Repeated analysis |
| `"full"` | Complete body text | Small sections only |

## Locators

Locators are the preferred way to address content:

- `bookmark:_mcp_abc123` — stable across edits (best)
- `heading:2.1` — heading by outline number
- `paragraph:42` — by paragraph index (fragile after edits)
- `page:3` — by page number

Always prefer bookmarks when available.

## Anti-Patterns

| Don't | Do instead |
|---|---|
| `get_text_content_live` on big docs | `get_document_tree` + `read_paragraphs` |
| `depth=0` on unknown docs | `depth=1`, then drill into sections |
| `content_strategy="full"` at depth > 1 | `"first_lines"`, then read specific paragraphs |
| Read all paragraphs sequentially | `search_in_document` to find what you need |
| Store paragraph indices for later | Store bookmarks — they survive edits |
| Insert text then manually set style | Use `insert_at_paragraph` + `set_paragraph_style` |
| Copy file manually to duplicate | Use `save_document_copy` |
| Edit without tracking | Enable `set_track_changes(true)` before edits |
| Ignore human comments | Read `list_comments`, act, then `resolve_comment` with reason |
| Leave doc unprotected while working | `set_document_protection(true)` at start, `false` when done |

## Tool Reference

### Navigation (read-only)

| Tool | What it does |
|---|---|
| `list_open_documents` | List all open docs in LibreOffice |
| `open_document` | Open a file in LibreOffice |
| `get_document_tree` | Heading tree (with depth + content_strategy) |
| `get_heading_children` | Drill into a heading's children |
| `read_paragraphs` | Read N paragraphs from a locator |
| `get_paragraph_count` | Total paragraph count |
| `search_in_document` | Native search with context |
| `get_page_count` | Page count |
| `list_sections` | Named text sections |
| `read_section` | Content of a named section |
| `list_bookmarks` | All bookmarks |
| `resolve_bookmark` | Bookmark → current paragraph index |

### Editing

| Tool | What it does |
|---|---|
| `replace_in_document` | Global find & replace (preserves formatting) |
| `insert_at_paragraph` | Insert text before/after a paragraph |
| `set_paragraph_text` | Replace entire paragraph content (preserves style) |
| `set_paragraph_style` | Change paragraph style (e.g. "Heading 1", "Text Body") |
| `delete_paragraph` | Remove a paragraph |
| `duplicate_paragraph` | Copy a paragraph (with style) after itself. `count>1` for blocks |

### Document Protection

| Tool | What it does |
|---|---|
| `set_document_protection` | Lock/unlock the UI (no password, just a toggle) |

Lock the document before making edits so the human can't accidentally interfere. Uses Writer's ProtectForm setting — UNO/MCP calls pass through normally. **Always unlock when done.**

```
set_document_protection(enabled=true)   # UI read-only, MCP still works
# ... do your work ...
set_document_protection(enabled=false)  # human can edit again
```

### Comments & Review

| Tool | What it does |
|---|---|
| `list_comments` | List all human comments (excludes AI summaries) |
| `add_comment` | Add a comment at a paragraph |
| `resolve_comment` | Close a comment with a reason (adds reply + marks resolved) |
| `delete_comment` | Delete a comment and its replies |

**Workflow**: Human leaves TODOs as comments → agent reads with `list_comments` → acts on each → closes with `resolve_comment(name, resolution="Done: updated age to 14")`.

### Track Changes

| Tool | What it does |
|---|---|
| `set_track_changes` | Enable/disable change recording |
| `get_tracked_changes` | List all redlines (type, author, date) |
| `accept_all_changes` | Accept all tracked changes |
| `reject_all_changes` | Reject all tracked changes |

**Workflow**: `set_track_changes(true)` → make edits → human reviews diffs in LO → `accept_all_changes()` or `reject_all_changes()`.

### Styles

| Tool | What it does |
|---|---|
| `list_styles` | List styles in a family (Paragraph, Character, Page...) |
| `get_style_info` | Detailed style properties (font, size, margins) |

Use `list_styles(family="ParagraphStyles")` to discover existing styles before applying with `set_paragraph_style`. Keep the document clean — use existing styles, don't create a patchwork.

### Document management

| Tool | What it does |
|---|---|
| `save_document` | Save active document |
| `save_document_copy` | Save As / duplicate under a new name |
| `refresh_indexes` | Refresh Table of Contents and other indexes |
| `update_fields` | Refresh all fields (dates, page numbers, cross-refs) |
| `get_document_properties` | Read metadata (title, author, subject, keywords) |
| `set_document_properties` | Update metadata |

### AI annotations

| Tool | What it does |
|---|---|
| `add_ai_summary` | Attach AI summary to a heading |
| `get_ai_summaries` | List all AI annotations |
| `remove_ai_summary` | Remove an AI annotation |

### Writer Tables

| Tool | What it does |
|---|---|
| `list_tables` | List all text tables (name, rows, cols) |
| `read_table` | Read all cell contents as 2D array |
| `write_table_cell` | Write to a cell (e.g. 'B3') |
| `create_table` | Create a new table at a paragraph position |

### Images

| Tool | What it does |
|---|---|
| `list_images` | List all images (name, dimensions, title) |
| `get_image_info` | Detailed info (URL, anchor, position) |
| `set_image_properties` | Resize, reanchor, set title/alt-text |

Resize keeping aspect ratio: `set_image_properties(image_name="Image1", width_mm=80)`.
Set caption: `set_image_properties(image_name="Image1", title="Photo de famille")`.
Anchor types: 0=AT_PARAGRAPH, 1=AS_CHARACTER (inline), 2=AT_PAGE, 4=AT_CHARACTER.

### Recent Documents

| Tool | What it does |
|---|---|
| `get_recent_documents` | LO history (paths + titles, max 20) |

Use this to find documents without knowing their exact path.

### Calc (.ods, .xlsx)

| Tool | What it does |
|---|---|
| `list_sheets` | Enumerate sheets |
| `get_sheet_info` | Used range, dimensions |
| `read_cells` | Read a cell range (e.g. `A1:D10`) |
| `write_cell` | Write a single cell |

Prefix cell addresses with sheet name for multi-sheet docs: `Sheet1.A1:D10`.

### Impress (.odp, .pptx)

| Tool | What it does |
|---|---|
| `list_slides` | Slide names and titles |
| `read_slide_text` | Text + notes from one slide |
| `get_presentation_info` | Metadata, dimensions |

## Typical Session: Document Revision

```
# 1. What's open?
list_open_documents

# 2. Duplicate the document for the new year
save_document_copy(target_path="C:/Users/.../IEF 2026.odt")
open_document(file_path="C:/Users/.../IEF 2026.odt")

# 3. Get the overview
get_document_tree(depth=1)

# 4. Lock the doc so human doesn't interfere
set_document_protection(enabled=true)

# 5. Check for human review comments
list_comments  # → human left "TODO: update ages" at paragraph 45

# 6. Enable track changes so human can review diffs
set_track_changes(enabled=true)

# 7. Global date replacement
replace_in_document(search="2025-2026", replace="2026-2027")

# 8. Refresh the Table of Contents
refresh_indexes()

# 9. Find specific content to update
search_in_document(pattern="13 ans", max_results=5)

# 10. Targeted paragraph edit (don't change it everywhere)
set_paragraph_text(locator="paragraph:45", text="Zachary, qui est âgé de 14 ans...")

# 11. Close the comment with what was done
resolve_comment(comment_name="__Annotation__42", resolution="Done: age updated 13→14")

# 12. Add a new section by duplicating an existing one as template
duplicate_paragraph(locator="bookmark:_mcp_abc123", count=4)

# 13. Update document metadata for the new year
set_document_properties(title="IEF 2026-2027", subject="Instruction en famille")

# 14. Save and unlock
save_document()
set_document_protection(enabled=false)  # human can edit again
```

## Performance Notes

- All UNO calls run on the LibreOffice main thread (thread-safe via AsyncCallback).
- Calls that take over 30 seconds will timeout (HTTP 504).
- Opening very large documents is the slowest operation — subsequent reads are fast.
- `search_in_document` uses native LibreOffice search, not paragraph-by-paragraph scanning.
- `refresh_indexes` and `update_fields` are fast even on large documents.
