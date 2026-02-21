"""Search tools — find and replace in documents."""

from .base import McpTool


class SearchInDocument(McpTool):
    name = "search_in_document"
    description = (
        "Search for text in a document with paragraph context. "
        "Uses LibreOffice native search. Returns matches with surrounding "
        "paragraphs for context, without loading the entire document."
    )
    parameters = {
        "type": "object",
        "properties": {
            "pattern": {
                "type": "string",
                "description": "Search string or regex",
            },
            "regex": {
                "type": "boolean",
                "description": "Use regular expression (default: False)",
            },
            "case_sensitive": {
                "type": "boolean",
                "description": "Case-sensitive search (default: False)",
            },
            "max_results": {
                "type": "integer",
                "description": "Max results to return (default: 20)",
            },
            "context_paragraphs": {
                "type": "integer",
                "description": "Paragraphs of context around each match (default: 1)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["pattern"],
    }

    def execute(self, pattern, regex=False, case_sensitive=False,
                max_results=20, context_paragraphs=1,
                file_path=None, **_):
        return self.services.writer.search_document(
            pattern, regex, case_sensitive, max_results,
            context_paragraphs, file_path)


class ReplaceInDocument(McpTool):
    name = "replace_in_document"
    description = "Find and replace text preserving all formatting."
    parameters = {
        "type": "object",
        "properties": {
            "search": {
                "type": "string",
                "description": "Text to find",
            },
            "replace": {
                "type": "string",
                "description": "Replacement text",
            },
            "regex": {
                "type": "boolean",
                "description": "Use regular expression (default: False)",
            },
            "case_sensitive": {
                "type": "boolean",
                "description": "Case-sensitive matching (default: False)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["search", "replace"],
    }

    def execute(self, search, replace, regex=False, case_sensitive=False,
                file_path=None, **_):
        return self.services.writer.replace_in_document(
            search, replace, regex, case_sensitive, file_path)
