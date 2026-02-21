"""Editing tools — insert, delete, modify paragraphs."""

from .base import McpTool


class InsertTextAtParagraph(McpTool):
    name = "insert_at_paragraph"
    description = (
        "Insert text before or after a specific paragraph. "
        "Preserves all existing formatting. Use locator for unified "
        "addressing (e.g. 'paragraph:5', 'bookmark:_mcp_x')."
    )
    parameters = {
        "type": "object",
        "properties": {
            "text": {
                "type": "string",
                "description": "Text to insert",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator string (preferred)",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Target paragraph index (legacy)",
            },
            "position": {
                "type": "string",
                "enum": ["before", "after"],
                "description": "'before' or 'after' (default: after)",
            },
            "style": {
                "type": "string",
                "description": "Paragraph style for the new paragraph "
                               "(e.g. 'Text Body', 'Heading 1'). "
                               "If omitted, inherits from adjacent paragraph.",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["text"],
    }

    def execute(self, paragraph_index=None, text="", position="after",
                locator=None, style=None, file_path=None, **_):
        return self.services.writer.insert_at_paragraph(
            paragraph_index, text, position, locator, style, file_path)


class InsertParagraphsBatch(McpTool):
    name = "insert_paragraphs_batch"
    description = (
        "Insert multiple paragraphs in one call. "
        "Each item in paragraphs is {\"text\": \"...\", \"style\": \"...\"}. "
        "Style is optional. All paragraphs are inserted in a single "
        "UNO transaction."
    )
    parameters = {
        "type": "object",
        "properties": {
            "paragraphs": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "text": {"type": "string"},
                        "style": {"type": "string"},
                    },
                    "required": ["text"],
                },
                "description": "List of {text, style?} objects to insert",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator string (preferred)",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Target paragraph index (legacy)",
            },
            "position": {
                "type": "string",
                "enum": ["before", "after"],
                "description": "'before' or 'after' (default: after)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["paragraphs"],
    }

    def execute(self, paragraphs=None, paragraph_index=None,
                position="after", locator=None, file_path=None, **_):
        return self.services.writer.insert_paragraphs_batch(
            paragraphs or [], paragraph_index, position, locator, file_path)


class DeleteDocumentParagraph(McpTool):
    name = "delete_paragraph"
    description = "Delete a paragraph from the document."
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5', 'bookmark:_mcp_x')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, paragraph_index=None, locator=None,
                file_path=None, **_):
        return self.services.writer.delete_paragraph(
            paragraph_index, locator, file_path)


class SetDocumentParagraphText(McpTool):
    name = "set_paragraph_text"
    description = "Replace the entire text of a paragraph (preserves style)."
    parameters = {
        "type": "object",
        "properties": {
            "text": {
                "type": "string",
                "description": "New text content for the paragraph",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5', 'bookmark:_mcp_x')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["text"],
    }

    def execute(self, paragraph_index=None, text="", locator=None,
                file_path=None, **_):
        return self.services.writer.set_paragraph_text(
            paragraph_index, text, locator, file_path)


class SetDocumentParagraphStyle(McpTool):
    name = "set_paragraph_style"
    description = (
        "Set the paragraph style (e.g. 'Heading 1', 'Text Body', 'List Bullet')."
    )
    parameters = {
        "type": "object",
        "properties": {
            "style_name": {
                "type": "string",
                "description": "Name of the paragraph style to apply",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5', 'bookmark:_mcp_x')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["style_name"],
    }

    def execute(self, style_name, paragraph_index=None, locator=None,
                file_path=None, **_):
        return self.services.writer.set_paragraph_style(
            style_name, paragraph_index, locator, file_path)


class DuplicateDocumentParagraph(McpTool):
    name = "duplicate_paragraph"
    description = (
        "Duplicate a paragraph (with its style) after itself. "
        "Use count > 1 to duplicate a block (e.g. heading + body paragraphs)."
    )
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5', 'bookmark:_mcp_x')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index (legacy)",
            },
            "count": {
                "type": "integer",
                "description": "Number of consecutive paragraphs to duplicate (default: 1)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, paragraph_index=None, locator=None, count=1,
                file_path=None, **_):
        return self.services.writer.duplicate_paragraph(
            paragraph_index, locator, count, file_path)
