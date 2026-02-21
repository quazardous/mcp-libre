"""
ImageService — images, graphic objects, and text frames.
"""

import logging
import os
from typing import Any, Dict

logger = logging.getLogger(__name__)


class ImageService:
    """Image and text-frame operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    # ==================================================================
    # Images
    # ==================================================================

    def list_images(self, file_path: str = None) -> Dict[str, Any]:
        """List all images/graphic objects in the document."""
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, 'getGraphicObjects'):
                return {"success": False,
                        "error": "Document does not support graphic objects"}

            graphics = doc.getGraphicObjects()
            images = []
            for name in graphics.getElementNames():
                graphic = graphics.getByName(name)
                entry = {"name": name}
                try:
                    size = graphic.getPropertyValue("Size")
                    entry["width_mm"] = size.Width // 100
                    entry["height_mm"] = size.Height // 100
                except Exception:
                    try:
                        entry["width_mm"] = graphic.Width // 100
                        entry["height_mm"] = graphic.Height // 100
                    except Exception:
                        pass
                try:
                    entry["description"] = graphic.getPropertyValue(
                        "Description")
                    entry["title"] = graphic.getPropertyValue("Title")
                except Exception:
                    pass
                try:
                    anchor = graphic.getAnchor()
                    entry["paragraph_index"] = self._base.anchor_para_index(
                        doc, anchor)
                    page = self._base.resolve_page(doc, name, anchor)
                    if page is not None:
                        entry["page"] = page
                except Exception:
                    pass
                images.append(entry)

            return {"success": True, "images": images,
                    "count": len(images)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_image_info(self, image_name: str,
                       file_path: str = None) -> Dict[str, Any]:
        """Get detailed info about a specific image."""
        try:
            doc = self._base.resolve_document(file_path)
            graphics = doc.getGraphicObjects()

            if not graphics.hasByName(image_name):
                return {"success": False,
                        "error": f"Image '{image_name}' not found",
                        "available": list(graphics.getElementNames())}

            graphic = graphics.getByName(image_name)
            info = {"name": image_name, "success": True}

            for prop in ("GraphicURL", "Description", "Title"):
                try:
                    info[prop] = graphic.getPropertyValue(prop)
                except Exception:
                    pass

            anchor_names = ["AT_PARAGRAPH", "AS_CHARACTER",
                            "AT_PAGE", "AT_FRAME", "AT_CHARACTER"]
            for prop in ("AnchorType", "HoriOrient", "VertOrient"):
                try:
                    val = graphic.getPropertyValue(prop)
                    try:
                        str_val = val.value
                    except AttributeError:
                        str_val = str(val)
                    info[prop] = str_val
                    if prop == "AnchorType" and str_val in anchor_names:
                        info["anchor_type_id"] = anchor_names.index(str_val)
                except Exception:
                    pass

            try:
                size = graphic.getPropertyValue("Size")
                info["width_mm"] = size.Width // 100
                info["height_mm"] = size.Height // 100
            except Exception:
                try:
                    info["width_mm"] = graphic.Width // 100
                    info["height_mm"] = graphic.Height // 100
                except Exception:
                    pass

            for prop, key in (("HoriOrientPosition", "hori_pos_mm"),
                              ("VertOrientPosition", "vert_pos_mm")):
                try:
                    info[key] = graphic.getPropertyValue(prop) // 100
                except Exception:
                    pass

            for prop in ("HoriOrientRelation", "VertOrientRelation"):
                try:
                    info[prop] = int(graphic.getPropertyValue(prop))
                except Exception:
                    pass

            for prop, key in (("TopMargin", "top_margin_mm"),
                              ("BottomMargin", "bottom_margin_mm"),
                              ("LeftMargin", "left_margin_mm"),
                              ("RightMargin", "right_margin_mm")):
                try:
                    info[key] = graphic.getPropertyValue(prop) // 100
                except Exception:
                    pass

            for prop in ("Surround", "TextWrapType"):
                try:
                    val = graphic.getPropertyValue(prop)
                    wrap_names = {0: "NONE", 1: "COLUMN", 2: "PARALLEL",
                                  3: "DYNAMIC", 4: "THROUGH"}
                    ival = int(val)
                    info["wrap"] = wrap_names.get(ival, str(ival))
                    info["wrap_id"] = ival
                    break
                except Exception:
                    pass

            try:
                crop = graphic.getPropertyValue("GraphicCrop")
                info["crop"] = {
                    "top_mm": crop.Top // 100,
                    "bottom_mm": crop.Bottom // 100,
                    "left_mm": crop.Left // 100,
                    "right_mm": crop.Right // 100,
                }
            except Exception:
                pass

            try:
                anchor = graphic.getAnchor()
                pidx = self._base.anchor_para_index(doc, anchor)
                if pidx is not None:
                    info["paragraph_index"] = pidx
                page = self._base.resolve_page(doc, image_name, anchor)
                if page is not None:
                    info["page"] = page
            except Exception:
                pass

            return info
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_image_properties(self, image_name: str,
                             width_mm: int = None,
                             height_mm: int = None,
                             title: str = None,
                             description: str = None,
                             anchor_type: int = None,
                             hori_orient: int = None,
                             vert_orient: int = None,
                             hori_orient_relation: int = None,
                             vert_orient_relation: int = None,
                             crop_top_mm: int = None,
                             crop_bottom_mm: int = None,
                             crop_left_mm: int = None,
                             crop_right_mm: int = None,
                             file_path: str = None) -> Dict[str, Any]:
        """Resize, reposition, crop, or update caption/alt-text."""
        try:
            doc = self._base.resolve_document(file_path)
            graphics = doc.getGraphicObjects()

            if not graphics.hasByName(image_name):
                return {"success": False,
                        "error": f"Image '{image_name}' not found",
                        "available": list(graphics.getElementNames())}

            graphic = graphics.getByName(image_name)
            changed = []

            if width_mm is not None or height_mm is not None:
                try:
                    size = graphic.getPropertyValue("Size")
                except Exception:
                    from com.sun.star.awt import Size as AwtSize
                    size = AwtSize(graphic.Width, graphic.Height)

                cur_w = size.Width
                cur_h = size.Height

                if width_mm is not None and height_mm is not None:
                    size.Width = width_mm * 100
                    size.Height = height_mm * 100
                elif width_mm is not None:
                    ratio = (width_mm * 100) / cur_w if cur_w else 1
                    size.Width = width_mm * 100
                    size.Height = int(cur_h * ratio)
                else:
                    ratio = (height_mm * 100) / cur_h if cur_h else 1
                    size.Height = height_mm * 100
                    size.Width = int(cur_w * ratio)

                graphic.setPropertyValue("Size", size)
                changed.append(
                    f"size={size.Width // 100}x{size.Height // 100}mm")

            if title is not None:
                graphic.setPropertyValue("Title", title)
                changed.append(f"title={title}")

            if description is not None:
                graphic.setPropertyValue("Description", description)
                changed.append("description set")

            if anchor_type is not None:
                graphic.setPropertyValue("AnchorType", anchor_type)
                labels = {0: "AT_PARAGRAPH", 1: "AS_CHARACTER",
                          2: "AT_PAGE", 3: "AT_FRAME", 4: "AT_CHARACTER"}
                changed.append(
                    f"anchor={labels.get(anchor_type, anchor_type)}")

            if hori_orient is not None:
                graphic.setPropertyValue("HoriOrient", hori_orient)
                changed.append(f"hori_orient={hori_orient}")

            if vert_orient is not None:
                graphic.setPropertyValue("VertOrient", vert_orient)
                changed.append(f"vert_orient={vert_orient}")

            if hori_orient_relation is not None:
                graphic.setPropertyValue(
                    "HoriOrientRelation", hori_orient_relation)
                changed.append(
                    f"hori_orient_relation={hori_orient_relation}")

            if vert_orient_relation is not None:
                graphic.setPropertyValue(
                    "VertOrientRelation", vert_orient_relation)
                changed.append(
                    f"vert_orient_relation={vert_orient_relation}")

            if any(v is not None for v in (crop_top_mm, crop_bottom_mm,
                                           crop_left_mm, crop_right_mm)):
                try:
                    crop = graphic.getPropertyValue("GraphicCrop")
                except Exception:
                    from com.sun.star.text import GraphicCrop
                    crop = GraphicCrop()
                if crop_top_mm is not None:
                    crop.Top = crop_top_mm * 100
                if crop_bottom_mm is not None:
                    crop.Bottom = crop_bottom_mm * 100
                if crop_left_mm is not None:
                    crop.Left = crop_left_mm * 100
                if crop_right_mm is not None:
                    crop.Right = crop_right_mm * 100
                graphic.setPropertyValue("GraphicCrop", crop)
                changed.append(
                    f"crop=T{crop.Top // 100}/B{crop.Bottom // 100}"
                    f"/L{crop.Left // 100}/R{crop.Right // 100}mm")

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "image": image_name,
                    "changes": changed}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_image(self, image_path: str,
                     paragraph_index: int = None,
                     locator: str = None,
                     caption: str = None,
                     with_frame: bool = True,
                     width_mm: int = None,
                     height_mm: int = None,
                     file_path: str = None) -> Dict[str, Any]:
        """Insert an image from a file path at a paragraph position."""
        try:
            if not os.path.isfile(image_path):
                return {"success": False,
                        "error": f"Image file not found: {image_path}"}

            doc = self._base.resolve_document(file_path)

            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            target, _ = self._registry.writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            import uno
            image_url = uno.systemPathToFileUrl(image_path)

            w_100mm = (width_mm or 80) * 100
            h_100mm = (height_mm or 80) * 100

            from com.sun.star.awt import Size as AwtSize
            doc_text = doc.getText()

            if with_frame:
                frame = doc.createInstance(
                    "com.sun.star.text.TextFrame")
                frame_size = AwtSize(w_100mm, h_100mm)
                frame.setPropertyValue("Size", frame_size)
                frame.setPropertyValue("AnchorType", 4)
                frame.setPropertyValue("HoriOrient", 0)
                frame.setPropertyValue("VertOrient", 0)

                cursor = doc_text.createTextCursorByRange(target.getEnd())
                doc_text.insertTextContent(cursor, frame, False)

                graphic = doc.createInstance(
                    "com.sun.star.text.TextGraphicObject")
                graphic.setPropertyValue("GraphicURL", image_url)
                graphic_size = AwtSize(w_100mm, h_100mm)
                graphic.setPropertyValue("Size", graphic_size)
                graphic.setPropertyValue("AnchorType", 0)
                graphic.setPropertyValue("HoriOrient", 2)
                graphic.setPropertyValue("VertOrient", 1)

                frame_text = frame.getText()
                frame_cursor = frame_text.createTextCursor()
                frame_text.insertTextContent(
                    frame_cursor, graphic, False)

                if caption:
                    frame_cursor = frame_text.createTextCursorByRange(
                        frame_text.getEnd())
                    frame_text.insertControlCharacter(
                        frame_cursor, 0, False)
                    frame_cursor = frame_text.createTextCursorByRange(
                        frame_text.getEnd())
                    frame_text.insertString(frame_cursor, caption, False)

                if doc.hasLocation():
                    self._base.store_doc(doc)

                return {"success": True,
                        "frame_name": frame.getName(),
                        "image_name": graphic.getName(),
                        "with_frame": True,
                        "caption": caption}
            else:
                graphic = doc.createInstance(
                    "com.sun.star.text.TextGraphicObject")
                graphic.setPropertyValue("GraphicURL", image_url)
                graphic_size = AwtSize(w_100mm, h_100mm)
                graphic.setPropertyValue("Size", graphic_size)
                graphic.setPropertyValue("AnchorType", 4)

                cursor = doc_text.createTextCursorByRange(target.getEnd())
                doc_text.insertTextContent(cursor, graphic, False)

                if doc.hasLocation():
                    self._base.store_doc(doc)

                return {"success": True,
                        "image_name": graphic.getName(),
                        "with_frame": False}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_image(self, image_name: str,
                     remove_frame: bool = True,
                     file_path: str = None) -> Dict[str, Any]:
        """Delete an image. Optionally remove its parent frame."""
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, 'getGraphicObjects'):
                return {"success": False,
                        "error": "Document does not support graphic objects"}

            graphics = doc.getGraphicObjects()
            if not graphics.hasByName(image_name):
                return {"success": False,
                        "error": f"Image '{image_name}' not found",
                        "available": list(graphics.getElementNames())}

            graphic = graphics.getByName(image_name)
            doc_text = doc.getText()

            anchor_text = graphic.getAnchor().getText()
            frame_name = None
            parent_frame = None
            if hasattr(doc, 'getTextFrames'):
                frames_access = doc.getTextFrames()
                for fname in frames_access.getElementNames():
                    fr = frames_access.getByName(fname)
                    if fr.getText() == anchor_text:
                        frame_name = fname
                        parent_frame = fr
                        break

            if frame_name is not None and remove_frame:
                doc_text.removeTextContent(parent_frame)
            elif frame_name is not None:
                anchor_text.removeTextContent(graphic)
            else:
                doc_text.removeTextContent(graphic)

            if doc.hasLocation():
                self._base.store_doc(doc)

            result = {"success": True, "deleted_image": image_name}
            if frame_name and remove_frame:
                result["deleted_frame"] = frame_name
            elif frame_name:
                result["kept_frame"] = frame_name
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def replace_image(self, image_name: str, new_image_path: str,
                      width_mm: int = None, height_mm: int = None,
                      file_path: str = None) -> Dict[str, Any]:
        """Replace an image's graphic source, keeping frame/position."""
        try:
            if not os.path.isfile(new_image_path):
                return {"success": False,
                        "error": f"Image file not found: {new_image_path}"}

            doc = self._base.resolve_document(file_path)
            graphics = doc.getGraphicObjects()
            if not graphics.hasByName(image_name):
                return {"success": False,
                        "error": f"Image '{image_name}' not found",
                        "available": list(graphics.getElementNames())}

            graphic = graphics.getByName(image_name)

            import uno
            image_url = uno.systemPathToFileUrl(new_image_path)
            graphic.setPropertyValue("GraphicURL", image_url)

            if width_mm is not None or height_mm is not None:
                size = graphic.getPropertyValue("Size")
                cur_w = size.Width
                cur_h = size.Height
                if width_mm is not None and height_mm is not None:
                    size.Width = width_mm * 100
                    size.Height = height_mm * 100
                elif width_mm is not None:
                    ratio = (width_mm * 100) / cur_w if cur_w else 1
                    size.Width = width_mm * 100
                    size.Height = int(cur_h * ratio)
                else:
                    ratio = (height_mm * 100) / cur_h if cur_h else 1
                    size.Height = height_mm * 100
                    size.Width = int(cur_w * ratio)
                graphic.setPropertyValue("Size", size)

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "image_name": image_name,
                    "new_source": new_image_path}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Text Frames
    # ==================================================================

    def list_text_frames(self, file_path: str = None) -> Dict[str, Any]:
        """List all text frames in the document."""
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, 'getTextFrames'):
                return {"success": False,
                        "error": "Document does not support text frames"}

            frames_access = doc.getTextFrames()

            frame_images = {}
            if hasattr(doc, 'getGraphicObjects'):
                graphics = doc.getGraphicObjects()
                for gname in graphics.getElementNames():
                    graphic = graphics.getByName(gname)
                    try:
                        anchor_text = graphic.getAnchor().getText()
                        for fname in frames_access.getElementNames():
                            fr = frames_access.getByName(fname)
                            if fr.getText() == anchor_text:
                                frame_images.setdefault(
                                    fname, []).append(gname)
                                break
                    except Exception:
                        pass

            result = []
            for fname in frames_access.getElementNames():
                frame = frames_access.getByName(fname)
                entry = {"name": fname}
                try:
                    size = frame.getPropertyValue("Size")
                    entry["width_mm"] = size.Width // 100
                    entry["height_mm"] = size.Height // 100
                except Exception:
                    pass
                try:
                    val = frame.getPropertyValue("AnchorType")
                    try:
                        entry["anchor_type"] = val.value
                    except AttributeError:
                        entry["anchor_type"] = str(val)
                except Exception:
                    pass
                for prop, key in (("HoriOrient", "hori_orient"),
                                  ("VertOrient", "vert_orient")):
                    try:
                        entry[key] = int(frame.getPropertyValue(prop))
                    except Exception:
                        pass
                try:
                    anchor = frame.getAnchor()
                    pidx = self._base.anchor_para_index(doc, anchor)
                    if pidx is not None:
                        entry["paragraph_index"] = pidx
                    page = self._base.resolve_page(doc, fname, anchor)
                    if page is not None:
                        entry["page"] = page
                except Exception:
                    pass
                if fname in frame_images:
                    entry["images"] = frame_images[fname]
                result.append(entry)

            return {"success": True, "frames": result,
                    "count": len(result)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_text_frame_info(self, frame_name: str,
                            file_path: str = None) -> Dict[str, Any]:
        """Get detailed info about a specific text frame."""
        try:
            doc = self._base.resolve_document(file_path)
            frames_access = doc.getTextFrames()

            if not frames_access.hasByName(frame_name):
                return {"success": False,
                        "error": f"Frame '{frame_name}' not found",
                        "available": list(frames_access.getElementNames())}

            frame = frames_access.getByName(frame_name)
            info = {"name": frame_name, "success": True}

            try:
                size = frame.getPropertyValue("Size")
                info["width_mm"] = size.Width // 100
                info["height_mm"] = size.Height // 100
            except Exception:
                pass

            anchor_names = ["AT_PARAGRAPH", "AS_CHARACTER",
                            "AT_PAGE", "AT_FRAME", "AT_CHARACTER"]
            try:
                val = frame.getPropertyValue("AnchorType")
                try:
                    str_val = val.value
                except AttributeError:
                    str_val = str(val)
                info["anchor_type"] = str_val
                if str_val in anchor_names:
                    info["anchor_type_id"] = anchor_names.index(str_val)
            except Exception:
                pass

            for prop in ("HoriOrient", "VertOrient"):
                try:
                    info[prop] = int(frame.getPropertyValue(prop))
                except Exception:
                    pass

            for prop, key in (("HoriOrientPosition", "hori_pos_mm"),
                              ("VertOrientPosition", "vert_pos_mm")):
                try:
                    info[key] = frame.getPropertyValue(prop) // 100
                except Exception:
                    pass

            for prop in ("Surround", "TextWrapType"):
                try:
                    val = frame.getPropertyValue(prop)
                    wrap_names = {0: "NONE", 1: "COLUMN", 2: "PARALLEL",
                                  3: "DYNAMIC", 4: "THROUGH"}
                    ival = int(val)
                    info["wrap"] = wrap_names.get(ival, str(ival))
                    info["wrap_id"] = ival
                    break
                except Exception:
                    pass

            try:
                anchor = frame.getAnchor()
                pidx = self._base.anchor_para_index(doc, anchor)
                if pidx is not None:
                    info["paragraph_index"] = pidx
                page = self._base.resolve_page(doc, frame_name, anchor)
                if page is not None:
                    info["page"] = page
            except Exception:
                pass

            try:
                frame_text = frame.getText().getString()
                if frame_text:
                    info["text"] = frame_text
            except Exception:
                pass

            if hasattr(doc, 'getGraphicObjects'):
                imgs = []
                graphics = doc.getGraphicObjects()
                ft = frame.getText()
                for gname in graphics.getElementNames():
                    graphic = graphics.getByName(gname)
                    try:
                        if graphic.getAnchor().getText() == ft:
                            imgs.append(gname)
                    except Exception:
                        pass
                if imgs:
                    info["images"] = imgs

            return info
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_text_frame_properties(self, frame_name: str,
                                  width_mm: int = None,
                                  height_mm: int = None,
                                  anchor_type: int = None,
                                  hori_orient: int = None,
                                  vert_orient: int = None,
                                  hori_pos_mm: int = None,
                                  vert_pos_mm: int = None,
                                  wrap: int = None,
                                  paragraph_index: int = None,
                                  file_path: str = None) -> Dict[str, Any]:
        """Modify text frame properties."""
        try:
            doc = self._base.resolve_document(file_path)
            frames_access = doc.getTextFrames()

            if not frames_access.hasByName(frame_name):
                return {"success": False,
                        "error": f"Frame '{frame_name}' not found",
                        "available": list(frames_access.getElementNames())}

            frame = frames_access.getByName(frame_name)
            changed = []

            if width_mm is not None or height_mm is not None:
                size = frame.getPropertyValue("Size")
                if width_mm is not None:
                    size.Width = width_mm * 100
                if height_mm is not None:
                    size.Height = height_mm * 100
                frame.setPropertyValue("Size", size)
                changed.append(
                    f"size={size.Width // 100}x{size.Height // 100}mm")

            if anchor_type is not None:
                frame.setPropertyValue("AnchorType", anchor_type)
                labels = {0: "AT_PARAGRAPH", 1: "AS_CHARACTER",
                          2: "AT_PAGE", 3: "AT_FRAME", 4: "AT_CHARACTER"}
                changed.append(
                    f"anchor={labels.get(anchor_type, anchor_type)}")

            if hori_orient is not None:
                frame.setPropertyValue("HoriOrient", hori_orient)
                changed.append(f"hori_orient={hori_orient}")

            if vert_orient is not None:
                frame.setPropertyValue("VertOrient", vert_orient)
                changed.append(f"vert_orient={vert_orient}")

            if hori_pos_mm is not None:
                frame.setPropertyValue(
                    "HoriOrientPosition", hori_pos_mm * 100)
                changed.append(f"hori_pos={hori_pos_mm}mm")

            if vert_pos_mm is not None:
                frame.setPropertyValue(
                    "VertOrientPosition", vert_pos_mm * 100)
                changed.append(f"vert_pos={vert_pos_mm}mm")

            if wrap is not None:
                try:
                    frame.setPropertyValue("Surround", wrap)
                except Exception:
                    frame.setPropertyValue("TextWrapType", wrap)
                wrap_names = {0: "NONE", 1: "COLUMN", 2: "PARALLEL",
                              3: "DYNAMIC", 4: "THROUGH"}
                changed.append(f"wrap={wrap_names.get(wrap, wrap)}")

            if paragraph_index is not None:
                text = doc.getText()
                cursor = text.createTextCursor()
                cursor.gotoStart(False)
                for _ in range(paragraph_index):
                    if not cursor.gotoNextParagraph(False):
                        break
                frame.attach(cursor)
                changed.append(f"paragraph_index={paragraph_index}")

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "frame": frame_name,
                    "changes": changed}
        except Exception as e:
            return {"success": False, "error": str(e)}
