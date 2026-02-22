"""Proximity tools — local heading navigation and surroundings discovery."""

from .base import McpTool


class NavigateHeading(McpTool):
    name = "navigate_heading"
    description = (
        "Navigate locally between headings from any position. "
        "Directions: next, previous, parent, first_child, "
        "next_sibling, previous_sibling. "
        "Uses the cached heading tree — O(1) after first call. "
        "When on body text, 'previous' returns the owning heading, "
        "'next' returns the heading after the current section. "
        "Use bookmark locators for best performance."
    )
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Starting position: 'bookmark:_mcp_xxx', "
                               "'paragraph:42', etc.",
            },
            "direction": {
                "type": "string",
                "enum": ["next", "previous", "parent", "first_child",
                         "next_sibling", "previous_sibling"],
                "description": "Navigation direction relative to "
                               "current heading context.",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["locator", "direction"],
    }

    def execute(self, locator, direction, file_path=None, **_):
        return self.services.writer.navigate_heading(
            locator, direction, file_path)


class GetSurroundings(McpTool):
    name = "get_surroundings"
    description = (
        "Discover what's near a position: images, tables, frames, "
        "comments, headings, and paragraph text within a radius. "
        "Returns parent heading chain for instant context. "
        "Use 'include' to skip expensive categories. "
        "Ideal after receiving a TODO comment — instantly see "
        "nearby images, tables, and section context."
    )
    parameters = {
        "type": "object",
        "properties": {
            "locator": {
                "type": "string",
                "description": "Center position: 'bookmark:_mcp_xxx', "
                               "'paragraph:42', etc.",
            },
            "radius": {
                "type": "integer",
                "description": "Paragraphs to scan in each direction "
                               "(default: 10, max: 50)",
            },
            "include": {
                "type": "array",
                "items": {
                    "type": "string",
                    "enum": ["paragraphs", "images", "tables",
                             "frames", "comments", "headings"],
                },
                "description": "Object types to include (default: all). "
                               "Use ['headings','comments'] for fast "
                               "context-only queries.",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["locator"],
    }

    def execute(self, locator, radius=10, include=None,
                file_path=None, **_):
        return self.services.writer.get_surroundings(
            locator, radius, include, file_path)