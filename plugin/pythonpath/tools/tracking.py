"""Track changes tools — enable/disable, list, accept, reject."""

from .base import McpTool


class SetDocumentTrackChanges(McpTool):
    name = "set_track_changes"
    description = (
        "Enable or disable change tracking (record changes). "
        "Enable before making edits so the human can review diffs. "
        "Disable after changes are accepted."
    )
    parameters = {
        "type": "object",
        "properties": {
            "enabled": {
                "type": "boolean",
                "description": "True to enable tracking, False to disable",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["enabled"],
    }

    def execute(self, enabled, file_path=None, **_):
        return self.services.comments.set_track_changes(enabled, file_path)


class GetDocumentTrackedChanges(McpTool):
    name = "get_tracked_changes"
    description = (
        "List all tracked changes (redlines) in the document. "
        "Returns change type, author, date, and comment for each redline."
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
        return self.services.comments.get_tracked_changes(file_path)


class AcceptAllDocumentChanges(McpTool):
    name = "accept_all_changes"
    description = "Accept all tracked changes in the document."
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
        return self.services.comments.accept_all_changes(file_path)


class RejectAllDocumentChanges(McpTool):
    name = "reject_all_changes"
    description = "Reject all tracked changes in the document."
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
        return self.services.comments.reject_all_changes(file_path)
