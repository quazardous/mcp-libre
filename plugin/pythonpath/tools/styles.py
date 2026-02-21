"""Style tools — list and inspect document styles."""

from .base import McpTool


class ListDocumentStyles(McpTool):
    name = "list_styles"
    description = (
        "List available styles in a family. "
        "Use this to discover which styles exist before applying them. "
        "Filter by is_in_use to see what the document actually uses."
    )
    parameters = {
        "type": "object",
        "properties": {
            "family": {
                "type": "string",
                "description": "ParagraphStyles, CharacterStyles, PageStyles, "
                               "FrameStyles, NumberingStyles (default: ParagraphStyles)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, family="ParagraphStyles", file_path=None, **_):
        return self.services.styles.list_styles(family, file_path)


class GetDocumentStyleInfo(McpTool):
    name = "get_style_info"
    description = (
        "Get detailed properties of a style (font, size, margins, etc.)."
    )
    parameters = {
        "type": "object",
        "properties": {
            "style_name": {
                "type": "string",
                "description": "Name of the style",
            },
            "family": {
                "type": "string",
                "description": "Style family (default: ParagraphStyles)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["style_name"],
    }

    def execute(self, style_name, family="ParagraphStyles",
                file_path=None, **_):
        return self.services.styles.get_style_info(
            style_name, family, file_path)
