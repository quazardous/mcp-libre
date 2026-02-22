"""Comment tools — list, add, resolve, delete comments."""

from .base import McpTool


class ListDocumentComments(McpTool):
    name = "list_comments"
    description = (
        "List comments in the document. Use author_filter to see only "
        "a specific agent's comments (e.g. 'ChatGPT', 'Claude'). "
        "Each comment includes author, content, resolved status, and "
        "paragraph index. Multi-agent: each AI should use its own author name."
    )
    parameters = {
        "type": "object",
        "properties": {
            "author_filter": {
                "type": "string",
                "description": "Filter by author name (e.g. 'ChatGPT', 'Claude'). "
                               "Omit to list all comments.",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, author_filter=None, file_path=None, **_):
        result = self.services.comments.list_comments(file_path)
        if author_filter and result.get("success") and result.get("comments"):
            af = author_filter.lower()
            result["comments"] = [
                c for c in result["comments"]
                if af in c.get("author", "").lower()
            ]
            result["filtered_by"] = author_filter
        return result


class AddDocumentComment(McpTool):
    name = "add_comment"
    description = (
        "Add a comment at a paragraph. Use your AI name as author "
        "(e.g. 'ChatGPT', 'Claude') for multi-agent collaboration. "
        "Other agents can filter comments by author."
    )
    parameters = {
        "type": "object",
        "properties": {
            "content": {
                "type": "string",
                "description": "Comment text",
            },
            "author": {
                "type": "string",
                "description": "Your AI agent name (e.g. 'ChatGPT', 'Claude'). "
                               "Use a consistent name for filtering.",
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


class ScanTasks(McpTool):
    name = "scan_tasks"
    description = (
        "Scan comments for actionable task prefixes: "
        "TODO-AI, FIX, QUESTION, VALIDATION, NOTE. "
        "Returns unresolved tasks with locators — use this to find "
        "what needs attention without reading the document body. "
        "Multi-agent: each AI leaves prefixed comments, others pick them up."
    )
    parameters = {
        "type": "object",
        "properties": {
            "unresolved_only": {
                "type": "boolean",
                "description": "Only return unresolved tasks (default: true)",
            },
            "prefix_filter": {
                "type": "string",
                "enum": ["TODO-AI", "FIX", "QUESTION", "VALIDATION", "NOTE"],
                "description": "Only return tasks with this prefix",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, unresolved_only=True, prefix_filter=None,
                file_path=None, **_):
        return self.services.comments.scan_tasks(
            unresolved_only, prefix_filter, file_path)


class GetWorkflowStatus(McpTool):
    name = "get_workflow_status"
    description = (
        "Read the master workflow dashboard comment (author: MCP-WORKFLOW). "
        "Returns key-value pairs like Phase, Images status, Annexes, etc. "
        "Use set_workflow_status to create or update."
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
        return self.services.comments.get_workflow_status(file_path)


class SetWorkflowStatus(McpTool):
    name = "set_workflow_status"
    description = (
        "Create or update the master workflow dashboard comment. "
        "Content should be key: value lines, e.g.:\\n"
        "  Phase: Rédaction\\n"
        "  Images: 3/10 insérées\\n"
        "  Annexes: En attente\\n"
        "The dashboard is a single comment at the start of the document "
        "authored by MCP-WORKFLOW. All agents can read/update it."
    )
    parameters = {
        "type": "object",
        "properties": {
            "content": {
                "type": "string",
                "description": "Dashboard content as key: value lines",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["content"],
    }

    def execute(self, content, file_path=None, **_):
        return self.services.comments.set_workflow_status(content, file_path)


class DeleteDocumentComment(McpTool):
    name = "delete_comment"
    description = (
        "Delete comments by name or author. "
        "Use author='MCP-BATCH' to clean up batch stop comments. "
        "Use author='MCP-WORKFLOW' to clean up workflow comments."
    )
    parameters = {
        "type": "object",
        "properties": {
            "comment_name": {
                "type": "string",
                "description": "Name/ID of the comment to delete",
            },
            "author": {
                "type": "string",
                "description": (
                    "Delete ALL comments by this author "
                    "(e.g. 'MCP-BATCH', 'MCP-WORKFLOW')"
                ),
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, comment_name=None, author=None,
                file_path=None, **_):
        return self.services.comments.delete_comment(
            comment_name, author, file_path)
