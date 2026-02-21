"""Comment tools — list, add, resolve, delete comments."""

from .base import McpTool


class ListDocumentComments(McpTool):
    name = "list_comments"
    description = (
        "List all comments/annotations in the document. "
        "Returns human comments (excludes MCP-AI summaries). Each comment "
        "includes author, content, resolved status, paragraph index, and "
        "whether it's a reply."
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
        return self.services.comments.list_comments(file_path)


class AddDocumentComment(McpTool):
    name = "add_comment"
    description = "Add a comment at a paragraph."
    parameters = {
        "type": "object",
        "properties": {
            "content": {
                "type": "string",
                "description": "Comment text",
            },
            "author": {
                "type": "string",
                "description": "Author name (default: AI Agent)",
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
        "required": ["content"],
    }

    def execute(self, content, author="AI Agent", paragraph_index=None,
                locator=None, file_path=None, **_):
        return self.services.comments.add_comment(
            content, author, paragraph_index, locator, file_path)


class ResolveDocumentComment(McpTool):
    name = "resolve_comment"
    description = (
        "Resolve a comment with an optional reason. "
        "Adds a reply with the resolution text, then marks as resolved. "
        "Use list_document_comments to find comment names."
    )
    parameters = {
        "type": "object",
        "properties": {
            "comment_name": {
                "type": "string",
                "description": "Name/ID of the comment to resolve",
            },
            "resolution": {
                "type": "string",
                "description": "Reason for resolution (e.g. 'Done: updated age to 14')",
            },
            "author": {
                "type": "string",
                "description": "Author of the resolution (default: AI Agent)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["comment_name"],
    }

    def execute(self, comment_name, resolution="", author="AI Agent",
                file_path=None, **_):
        return self.services.comments.resolve_comment(
            comment_name, resolution, author, file_path)


class DeleteDocumentComment(McpTool):
    name = "delete_comment"
    description = "Delete a comment and all its replies."
    parameters = {
        "type": "object",
        "properties": {
            "comment_name": {
                "type": "string",
                "description": "Name/ID of the comment to delete",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["comment_name"],
    }

    def execute(self, comment_name, file_path=None, **_):
        return self.services.comments.delete_comment(
            comment_name, file_path)
