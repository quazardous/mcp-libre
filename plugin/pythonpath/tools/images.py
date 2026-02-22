"""Image tools — list, inspect, modify, insert, delete, replace images."""

from .base import McpTool


class ListDocumentImages(McpTool):
    name = "list_images"
    description = (
        "List all images/graphic objects in the document. "
        "Returns name, dimensions, title, and description for each image."
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
        return self.services.images.list_images(file_path)


class GetDocumentImageInfo(McpTool):
    name = "get_image_info"
    description = (
        "Get detailed info about a specific image. "
        "Returns URL, dimensions, anchor type, orientation, and paragraph index."
    )
    parameters = {
        "type": "object",
        "properties": {
            "image_name": {
                "type": "string",
                "description": "Name of the image/graphic object",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["image_name"],
    }

    def execute(self, image_name, file_path=None, **_):
        return self.services.images.get_image_info(image_name, file_path)


class SetDocumentImageProperties(McpTool):
    name = "set_image_properties"
    description = (
        "Resize, reposition, crop, or update caption/alt-text for an image."
    )
    parameters = {
        "type": "object",
        "properties": {
            "image_name": {
                "type": "string",
                "description": "Name of the image (use list_document_images to find)",
            },
            "width_mm": {"type": "integer", "description": "New width in millimeters"},
            "height_mm": {"type": "integer", "description": "New height in millimeters"},
            "title": {"type": "string", "description": "Image title / caption text"},
            "description": {"type": "string", "description": "Alt-text for accessibility"},
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
            "hori_orient_relation": {
                "type": "integer",
                "description": "0=PARAGRAPH, 1=FRAME, 2=PAGE...",
            },
            "vert_orient_relation": {
                "type": "integer",
                "description": "0=PARAGRAPH, 1=FRAME, 2=PAGE...",
            },
            "crop_top_mm": {"type": "integer", "description": "Crop from top in mm"},
            "crop_bottom_mm": {"type": "integer", "description": "Crop from bottom in mm"},
            "crop_left_mm": {"type": "integer", "description": "Crop from left in mm"},
            "crop_right_mm": {"type": "integer", "description": "Crop from right in mm"},
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["image_name"],
    }

    def execute(self, image_name, width_mm=None, height_mm=None,
                title=None, description=None, anchor_type=None,
                hori_orient=None, vert_orient=None,
                hori_orient_relation=None, vert_orient_relation=None,
                crop_top_mm=None, crop_bottom_mm=None,
                crop_left_mm=None, crop_right_mm=None,
                file_path=None, **_):
        return self.services.images.set_image_properties(
            image_name, width_mm, height_mm, title, description,
            anchor_type, hori_orient, vert_orient,
            hori_orient_relation, vert_orient_relation,
            crop_top_mm, crop_bottom_mm, crop_left_mm, crop_right_mm,
            file_path)


class InsertDocumentImage(McpTool):
    name = "insert_image"
    description = (
        "Insert an image from a file path or URL into the document. "
        "Supports http/https URLs (image is downloaded automatically). "
        "By default the image is wrapped in a text frame (caption frame). "
        "Set with_frame=False to insert a standalone image."
    )
    parameters = {
        "type": "object",
        "properties": {
            "image_path": {
                "type": "string",
                "description": "Absolute path to the image file on disk, or an HTTP/HTTPS URL",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator for insertion point (e.g. 'paragraph:5')",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index to insert after (legacy)",
            },
            "caption": {
                "type": "string",
                "description": "Caption text below the image (optional)",
            },
            "with_frame": {
                "type": "boolean",
                "description": "Wrap in a text frame (default: True)",
            },
            "width_mm": {
                "type": "integer",
                "description": "Width in mm (default: 80)",
            },
            "height_mm": {
                "type": "integer",
                "description": "Height in mm (default: 80)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["image_path"],
    }

    def execute(self, image_path, paragraph_index=None, locator=None,
                caption=None, with_frame=True, width_mm=None,
                height_mm=None, file_path=None, **_):
        return self.services.images.insert_image(
            image_path, paragraph_index, locator, caption,
            with_frame, width_mm, height_mm, file_path)


class DeleteDocumentImage(McpTool):
    name = "delete_image"
    description = (
        "Delete an image from the document. "
        "If the image is inside a text frame and remove_frame=True (default), "
        "the entire frame (image + caption) is removed."
    )
    parameters = {
        "type": "object",
        "properties": {
            "image_name": {
                "type": "string",
                "description": "Name of the image (use list_document_images to find)",
            },
            "remove_frame": {
                "type": "boolean",
                "description": "Also remove parent frame (default: True)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["image_name"],
    }

    def execute(self, image_name, remove_frame=True,
                file_path=None, **_):
        return self.services.images.delete_image(
            image_name, remove_frame, file_path)


class ReplaceDocumentImage(McpTool):
    name = "replace_image"
    description = (
        "Replace an image's source file, keeping its frame and position. "
        "The image stays in its current frame with the same anchor, "
        "orientation, and caption. Only the graphic source changes. "
        "Supports http/https URLs (image is downloaded automatically)."
    )
    parameters = {
        "type": "object",
        "properties": {
            "image_name": {
                "type": "string",
                "description": "Name of the image to replace",
            },
            "new_image_path": {
                "type": "string",
                "description": "Absolute path to the new image file on disk, or an HTTP/HTTPS URL",
            },
            "width_mm": {
                "type": "integer",
                "description": "New width in mm (optional, keeps current if omitted)",
            },
            "height_mm": {
                "type": "integer",
                "description": "New height in mm (optional, keeps current if omitted)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["image_name", "new_image_path"],
    }

    def execute(self, image_name, new_image_path, width_mm=None,
                height_mm=None, file_path=None, **_):
        return self.services.images.replace_image(
            image_name, new_image_path, width_mm, height_mm, file_path)
