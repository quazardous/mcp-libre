"""Structural tools — sections, bookmarks, indexes, fields."""

from .base import McpTool


class ListDocumentSections(McpTool):
    name = "list_sections"
    description = "List all named text sections in a document."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.writer.list_sections(file_path)


class ReadDocumentSection(McpTool):
    name = "read_section"
    description = "Read the content of a named text section."
    parameters = {
        "type": "object",
        "properties": {
            "section_name": {
                "type": "string",
                "description": "Name of the section",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["section_name"],
    }

    def execute(self, section_name, file_path=None, **_):
        return self.services.writer.read_section(section_name, file_path)


class ListDocumentBookmarks(McpTool):
    name = "list_bookmarks"
    description = "List all bookmarks in a document."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.writer.list_bookmarks(file_path)


class ResolveDocumentBookmark(McpTool):
    name = "resolve_bookmark"
    description = (
        "Resolve a heading bookmark to its current paragraph index. "
        "Bookmarks are stable identifiers that survive document edits. "
        "Use this to find the current position of a previously bookmarked heading."
    )
    parameters = {
        "type": "object",
        "properties": {
            "bookmark_name": {
                "type": "string",
                "description": "Bookmark name (e.g. _mcp_a1b2c3d4)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["bookmark_name"],
    }

    def execute(self, bookmark_name, file_path=None, **_):
        return self.services.writer.resolve_bookmark(bookmark_name, file_path)


class RefreshDocumentIndexes(McpTool):
    name = "refresh_indexes"
    description = (
        "Refresh all document indexes (Table of Contents, alphabetical, etc.). "
        "Call this after modifying headings or text referenced by indexes."
    )
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.writer.refresh_indexes(file_path)


class UpdateDocumentFields(McpTool):
    name = "update_fields"
    description = "Refresh all text fields (dates, page numbers, cross-references)."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.writer.update_fields(file_path)
