"""Navigation tools — tree, paragraphs, pages."""

from .base import McpTool


class GetDocumentTree(McpTool):
    name = "get_document_tree"
    description = (
        "Get the heading tree of a document without loading full text. "
        "Use depth to control how many levels are returned "
        "(1=top-level only, 2=two levels, 0=full tree). "
        "Use content_strategy to control body text visibility: "
        "none, first_lines, ai_summary_first, full."
    )
    parameters = {
        "type": "object",
        "properties": {
            "content_strategy": {
                "type": "string",
                "enum": ["none", "first_lines", "ai_summary_first", "full"],
                "description": "What to show for body text (default: first_lines)",
            },
            "depth": {
                "type": "integer",
                "description": "Heading levels to return (default: 1)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, content_strategy="first_lines", depth=1,
                file_path=None, **_):
        return self.services.writer.get_document_tree(
            content_strategy, depth, file_path)


class GetHeadingChildren(McpTool):
    name = "get_heading_children"
    description = (
        "Drill down into a heading to see its children. "
        "Use locator for unified addressing (e.g. 'bookmark:_mcp_x', "
        "'heading:2.1', 'paragraph:5'). Legacy params heading_bookmark "
        "and heading_para_index are still supported."
    )
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Unified locator string (preferred)",
            },
            "heading_para_index": {
                "type": "integer",
                "description": "Paragraph index of the parent heading (legacy)",
            },
            "heading_bookmark": {
                "type": "string",
                "description": "Bookmark name (legacy)",
            },
            "content_strategy": {
                "type": "string",
                "enum": ["none", "first_lines", "ai_summary_first", "full"],
                "description": "none, first_lines, ai_summary_first, full",
            },
            "depth": {
                "type": "integer",
                "description": "Sub-levels to include (default: 1)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, heading_para_index=None, heading_bookmark=None,
                locator=None, content_strategy="first_lines", depth=1,
                file_path=None, **_):
        return self.services.writer.get_heading_children(
            heading_para_index, heading_bookmark, locator,
            content_strategy, depth, file_path)


class ReadDocumentParagraphs(McpTool):
    name = "read_paragraphs"
    description = (
        "Read specific paragraphs by locator or index range. "
        "Use get_document_tree first to find which paragraphs to read. "
        "Locator examples: 'paragraph:0', 'page:2', 'bookmark:_mcp_x', "
        "'section:Introduction'."
    )
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Unified locator string (e.g. 'paragraph:0', 'page:2')",
            },
            "start_index": {
                "type": "integer",
                "description": "Zero-based index of first paragraph (legacy)",
            },
            "count": {
                "type": "integer",
                "description": "Number of paragraphs to read (default: 10)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, start_index=None, count=10, locator=None,
                file_path=None, **_):
        return self.services.writer.read_paragraphs(
            start_index, count, locator, file_path)


class GetDocumentParagraphCount(McpTool):
    name = "get_paragraph_count"
    description = "Get total paragraph count of a document."
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
        return self.services.writer.get_paragraph_count(file_path)


class GetDocumentPageCount(McpTool):
    name = "get_page_count"
    description = "Get the page count of a document."
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
        return self.services.writer.get_page_count(file_path)


class GotoPage(McpTool):
    name = "goto_page"
    description = (
        "Scroll the LibreOffice view to a specific page. "
        "Use this to visually navigate to a page so the user can see it."
    )
    parameters = {
        "type": "object",
        "properties": {
            "page": {
                "type": "integer",
                "description": "Page number (1-based)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["page"],
    }

    def execute(self, page, file_path=None, **_):
        return self.services.writer.goto_page(page, file_path)


class GetPageObjects(McpTool):
    name = "get_page_objects"
    description = (
        "Get images and tables on a page. "
        "Pass page number directly, OR a locator/paragraph_index to "
        "resolve the page automatically. Use this to find objects near "
        "a paragraph or comment."
    )
    parameters = {
        "type": "object",
        "properties": {
            "page": {
                "type": "integer",
                "description": "Page number (1-based)",
            },
            "locator": {
                "type": "string",
                "description": "Locator to resolve page from (e.g. 'paragraph:89')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index to resolve page from",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, page=None, locator=None, paragraph_index=None,
                file_path=None, **_):
        return self.services.writer.get_page_objects(
            page, locator, paragraph_index, file_path)
