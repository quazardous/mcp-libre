"""Impress tools — slides, presentations."""

from .base import McpTool


class ListPresentationSlides(McpTool):
    name = "list_slides"
    description = (
        "List all slides in an Impress presentation. "
        "Returns slide count, names, layout info, and first-shape title."
    )
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the presentation (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.impress.list_slides(file_path)


class ReadPresentationSlide(McpTool):
    name = "read_slide_text"
    description = (
        "Get all text from a presentation slide and its notes page."
    )
    parameters = {
        "type": "object",
        "properties": {
            "slide_index": {
                "type": "integer",
                "description": "Zero-based slide index",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the presentation (optional)",
            },
        },
        "required": ["slide_index"],
    }

    def execute(self, slide_index, file_path=None, **_):
        return self.services.impress.read_slide_text(
            slide_index, file_path)


class GetPresentationInfo(McpTool):
    name = "get_presentation_info"
    description = "Get presentation metadata: slide count, dimensions, master pages."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the presentation (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.impress.get_presentation_info(file_path)
