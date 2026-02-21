"""Document protection tools."""

from .base import McpTool


class SetDocumentProtection(McpTool):
    name = "set_document_protection"
    description = (
        "Lock or unlock the document for human editing. "
        "When locked (enabled=True), the document UI becomes read-only. "
        "All MCP/UNO calls still work normally through the protection. "
        "No password, just a boolean toggle."
    )
    parameters = {
        "type": "object",
        "properties": {
            "enabled": {
                "type": "boolean",
                "description": "True to lock (human can't edit), False to unlock",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["enabled"],
    }

    def execute(self, enabled, file_path=None, **_):
        return self.services.writer.set_document_protection(
            enabled, file_path)
