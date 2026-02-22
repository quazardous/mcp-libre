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


class SearchBoolean(McpTool):
    name = "search_boolean"
    description = (
        "Boolean full-text search with stemming (Snowball). "
        "Language auto-detected from document locale. "
        "Supports AND, OR, NOT, NEAR/N operators. "
        "Stemming handles singular/plural and word forms "
        "(enfants=enfant, protection=protections). "
        "Results include page numbers, paragraph context, "
        "and nearest heading. "
        "Use around_page + page_radius to restrict results "
        "near a specific page."
    )
    parameters = {
        "type": "object",
        "properties": {
            "query": {
                "type": "string",
                "description": (
                    "Boolean query. Examples: "
                    "'enfant AND protection', "
                    "'juge OR tribunal', "
                    "'enfant NOT maltraitance', "
                    "'enfant NEAR/3 protection', "
                    "'enfants protection' (implicit AND)"
                ),
            },
            "max_results": {
                "type": "integer",
                "description": "Max results to return (default: 20)",
            },
            "context_paragraphs": {
                "type": "integer",
                "description": (
                    "Paragraphs of context around each match "
                    "(default: 1)"
                ),
            },
            "around_page": {
                "type": "integer",
                "description": (
                    "Restrict results to pages near this page "
                    "(optional)"
                ),
            },
            "page_radius": {
                "type": "integer",
                "description": (
                    "Page radius for around_page filter "
                    "(default: 1, meaning +/-1 page)"
                ),
            },
            "include_pages": {
                "type": "boolean",
                "description": (
                    "Add page numbers to results. "
                    "Costs ~30s on first call (cached after). "
                    "Automatic when around_page is set. "
                    "(default: false)"
                ),
            },
            "file_path": {
                "type": "string",
                "description": (
                    "Absolute path to the document (optional)"
                ),
            },
        },
        "required": ["query"],
    }

    def execute(self, query, max_results=20, context_paragraphs=1,
                around_page=None, page_radius=1,
                include_pages=False, file_path=None, **_):
        return self.services.writer.search_boolean(
            query, max_results, context_paragraphs,
            around_page, page_radius, include_pages, file_path)


class GetIndexStats(McpTool):
    name = "get_index_stats"
    description = (
        "Get full-text index statistics: paragraph count, unique stems, "
        "detected language, top 20 most frequent stems. "
        "Builds the index if not cached."
    )
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": (
                    "Absolute path to the document (optional)"
                ),
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.writer.get_index_stats(file_path)
