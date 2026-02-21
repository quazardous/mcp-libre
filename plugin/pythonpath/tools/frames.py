"""Text frame tools — list, inspect, modify frames."""

from .base import McpTool


class ListDocumentFrames(McpTool):
    name = "list_text_frames"
    description = (
        "List all text frames in the document. "
        "Returns name, dimensions, anchor type, orientation, paragraph_index, "
        "and contained images for each frame."
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
        return self.services.images.list_text_frames(file_path)


class GetDocumentFrameInfo(McpTool):
    name = "get_text_frame_info"
    description = (
        "Get detailed info about a specific text frame. "
        "Returns size, position, anchor type, orientation, wrap mode, "
        "paragraph_index, contained text (caption), and contained images."
    )
    parameters = {
        "type": "object",
        "properties": {
            "frame_name": {
                "type": "string",
                "description": "Name of the text frame (use list_document_frames)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["frame_name"],
    }

    def execute(self, frame_name, file_path=None, **_):
        return self.services.images.get_text_frame_info(
            frame_name, file_path)


class SetDocumentFrameProperties(McpTool):
    name = "set_text_frame_properties"
    description = "Modify text frame properties (size, position, wrap, anchor)."
    parameters = {
        "type": "object",
        "properties": {
            "frame_name": {
                "type": "string",
                "description": "Name of the frame (use list_document_frames to find)",
            },
            "width_mm": {"type": "integer", "description": "New width in millimeters"},
            "height_mm": {"type": "integer", "description": "New height in millimeters"},
            "anchor_type": {
                "type": "integer",
                "description": "0=AT_PARAGRAPH, 1=AS_CHARACTER, 2=AT_PAGE, 4=AT_CHARACTER",
            },
            "hori_orient": {
                "type": "integer",
                "description": "0=NONE, 1=RIGHT, 2=CENTER, 3=LEFT",
            },
            "vert_orient": {
                "type": "integer",
                "description": "0=NONE, 1=TOP, 2=CENTER, 3=BOTTOM",
            },
            "hori_pos_mm": {
                "type": "integer",
                "description": "Horizontal position in mm (when hori_orient=NONE)",
            },
            "vert_pos_mm": {
                "type": "integer",
                "description": "Vertical position in mm (when vert_orient=NONE)",
            },
            "wrap": {
                "type": "integer",
                "description": "0=NONE, 1=COLUMN, 2=PARALLEL, 3=DYNAMIC, 4=THROUGH",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Move anchor to this paragraph index",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["frame_name"],
    }

    def execute(self, frame_name, width_mm=None, height_mm=None,
                anchor_type=None, hori_orient=None, vert_orient=None,
                hori_pos_mm=None, vert_pos_mm=None, wrap=None,
                paragraph_index=None, file_path=None, **_):
        return self.services.images.set_text_frame_properties(
            frame_name, width_mm, height_mm, anchor_type,
            hori_orient, vert_orient, hori_pos_mm, vert_pos_mm,
            wrap, paragraph_index, file_path)
