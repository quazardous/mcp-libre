"""AI annotation tools — summaries attached to headings."""

from .base import McpTool


class AddDocumentAiSummary(McpTool):
    name = "add_ai_summary"
    description = (
        "Add an AI annotation/summary to a heading. "
        "The summary is stored as a Writer annotation with Author='MCP-AI'. "
        "It will be shown when using content_strategy='ai_summary_first'."
    )
    parameters = {
        "type": "object",
        "properties": {
            "summary": {
                "type": "string",
                "description": "Summary text to attach",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5', 'heading:2.1')",
            },
            "para_index": {
                "type": "integer",
                "description": "Paragraph index of the heading (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["summary"],
    }

    def execute(self, para_index=None, summary="", locator=None,
                file_path=None, **_):
        return self.services.writer.add_ai_summary(
            para_index, summary, locator, file_path)


class GetDocumentAiSummaries(McpTool):
    name = "get_ai_summaries"
    description = "List all AI annotations in a document."
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
        return self.services.writer.get_ai_summaries(file_path)


class RemoveDocumentAiSummary(McpTool):
    name = "remove_ai_summary"
    description = "Remove an AI annotation from a heading."
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Unified locator (e.g. 'paragraph:5')",
            },
            "para_index": {
                "type": "integer",
                "description": "Paragraph index of the heading (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, para_index=None, locator=None,
                file_path=None, **_):
        return self.services.writer.remove_ai_summary(
            para_index, locator, file_path)
