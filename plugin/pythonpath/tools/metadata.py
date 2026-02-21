"""Document metadata tools — read/write properties."""

from .base import McpTool


class GetDocumentMetadata(McpTool):
    name = "get_document_properties"
    description = "Read document metadata (title, author, subject, keywords, dates)."
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
        return self.services.writer.get_document_properties(file_path)


class SetDocumentMetadata(McpTool):
    name = "set_document_properties"
    description = "Update document metadata (title, author, subject, etc.)."
    parameters = {
        "type": "object",
        "properties": {
            "title": {"type": "string", "description": "Document title"},
            "author": {"type": "string", "description": "Document author"},
            "subject": {"type": "string", "description": "Document subject"},
            "description": {"type": "string", "description": "Document description"},
            "keywords": {
                "type": "array",
                "items": {"type": "string"},
                "description": "List of keywords",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, title=None, author=None, subject=None,
                description=None, keywords=None,
                file_path=None, **_):
        return self.services.writer.set_document_properties(
            title, author, subject, description, keywords, file_path)
