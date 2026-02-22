"""
StructuralService — sections, pages, indexes, locator resolution.
"""

import logging
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)


class StructuralService:
    """Structural navigation: sections, pages, indexes, locators."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base

    # ==================================================================
    # Locator resolution
    # ==================================================================

    def resolve_writer_locator(self, doc, loc_type: str,
                               loc_value: str) -> Dict[str, Any]:
        """Resolve Writer-specific locators: bookmark, page, section, heading."""
        if loc_type == "bookmark":
            result = self.resolve_bookmark(loc_value, doc=doc)
            if not result.get("success"):
                raise ValueError(
                    result.get("error", f"Bookmark '{loc_value}' not found"))
            return {"para_index": result["para_index"]}

        if loc_type == "page":
            page_num = int(loc_value)
            try:
                controller = doc.getCurrentController()
                vc = controller.getViewCursor()
                saved = doc.getText().createTextCursorByRange(
                    vc.getStart())
                doc.lockControllers()
                try:
                    vc.jumpToPage(page_num)
                    vc.jumpToStartOfPage()
                    anchor = vc.getStart()
                finally:
                    vc.gotoRange(saved, False)
                    doc.unlockControllers()
                para_ranges = self._writer.get_paragraph_ranges(doc)
                text_obj = doc.getText()
                para_idx = self._writer.find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                return {"para_index": para_idx}
            except Exception as e:
                raise ValueError(
                    f"Cannot resolve page:{loc_value} — {e}")

        if loc_type == "section":
            if not hasattr(doc, "getTextSections"):
                raise ValueError("Document does not support sections")
            sections = doc.getTextSections()
            if not sections.hasByName(loc_value):
                raise ValueError(f"Section '{loc_value}' not found")
            section = sections.getByName(loc_value)
            anchor = section.getAnchor()
            para_ranges = self._writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()
            para_idx = self._writer.find_paragraph_for_range(
                anchor, para_ranges, text_obj)
            return {"para_index": para_idx, "section_name": loc_value}

        if loc_type == "heading":
            parts = [int(p) for p in loc_value.split(".")]
            tree = self._writer.tree.build_heading_tree(doc)
            node = tree
            for part in parts:
                children = node.get("children", [])
                if part < 1 or part > len(children):
                    raise ValueError(
                        f"Heading index {part} out of range "
                        f"(1..{len(children)}) in 'heading:{loc_value}'")
                node = children[part - 1]
            return {"para_index": node["para_index"]}

        raise ValueError(f"Unknown Writer locator type: '{loc_type}'")

    # ==================================================================
    # Bookmark resolution
    # ==================================================================

    def resolve_bookmark(self, bookmark_name: str,
                         file_path: str = None,
                         doc=None) -> Dict[str, Any]:
        """Resolve a bookmark name to its current paragraph index."""
        try:
            if doc is None:
                doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getBookmarks"):
                return {"success": False,
                        "error": "Document doesn't support bookmarks"}

            bookmarks = doc.getBookmarks()
            if not bookmarks.hasByName(bookmark_name):
                return {"success": False,
                        "error": f"Bookmark '{bookmark_name}' not found"}

            bm = bookmarks.getByName(bookmark_name)
            anchor = bm.getAnchor()
            para_ranges = self._writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()
            para_idx = self._writer.find_paragraph_for_range(
                anchor, para_ranges, text_obj)

            # Reuse para_ranges to get heading info (no second enum)
            heading_info = {}
            if para_idx < len(para_ranges):
                element = para_ranges[para_idx]
                if element.supportsService(
                        "com.sun.star.text.Paragraph"):
                    try:
                        heading_info["text"] = element.getString()
                        heading_info["outline_level"] = \
                            element.getPropertyValue("OutlineLevel")
                    except Exception:
                        pass

            return {"success": True, "bookmark": bookmark_name,
                    "para_index": para_idx, **heading_info}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Sections
    # ==================================================================

    def list_sections(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getTextSections"):
                return {"success": True, "sections": [], "count": 0}
            supplier = doc.getTextSections()
            names = supplier.getElementNames()
            sections = []
            for name in names:
                section = supplier.getByName(name)
                sections.append({
                    "name": name,
                    "is_visible": getattr(section, "IsVisible", True),
                    "is_protected": getattr(section, "IsProtected", False),
                })
            return {"success": True, "sections": sections,
                    "count": len(sections)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_section(self, section_name: str,
                     file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getTextSections"):
                return {"success": False,
                        "error": "Document does not support sections"}
            sections = doc.getTextSections()
            if not sections.hasByName(section_name):
                return {"success": False,
                        "error": f"Section '{section_name}' not found",
                        "available": list(sections.getElementNames())}
            section = sections.getByName(section_name)
            content = section.getAnchor().getString()
            return {"success": True, "section_name": section_name,
                    "content": content, "length": len(content)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Bookmarks
    # ==================================================================

    def list_bookmarks(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getBookmarks"):
                return {"success": True, "bookmarks": [], "count": 0}
            bookmarks = doc.getBookmarks()
            names = bookmarks.getElementNames()
            result = []
            for name in names:
                bm = bookmarks.getByName(name)
                anchor_text = bm.getAnchor().getString()
                result.append({
                    "name": name,
                    "preview": anchor_text[:100] if anchor_text else "",
                })
            return {"success": True, "bookmarks": result,
                    "count": len(result)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Pages
    # ==================================================================

    def get_page_count(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            controller = doc.getCurrentController()
            if controller:
                vc = controller.getViewCursor()
                saved = doc.getText().createTextCursorByRange(
                    vc.getStart())
                doc.lockControllers()
                try:
                    vc.jumpToLastPage()
                    page_count = vc.getPage()
                finally:
                    vc.gotoRange(saved, False)
                    doc.unlockControllers()
                return {"success": True, "page_count": page_count}
            return {"success": False,
                    "error": "Could not determine page count"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def goto_page(self, page: int,
                  file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            controller = doc.getCurrentController()
            vc = controller.getViewCursor()
            vc.jumpToPage(page)
            return {"success": True, "page": vc.getPage()}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_page_objects(self, page: int = None, locator: str = None,
                         paragraph_index: int = None,
                         file_path: str = None) -> Dict[str, Any]:
        """Get images, tables, and frames on a page.

        Uses lockControllers + cursor save/restore to prevent
        visible viewport jumping while scanning objects.
        """
        try:
            doc = self._base.resolve_document(file_path)
            controller = doc.getCurrentController()
            vc = controller.getViewCursor()

            if page is None:
                if locator:
                    resolved = self._base.resolve_locator(doc, locator)
                    para_idx = resolved.get("para_index", 0)
                elif paragraph_index is not None:
                    para_idx = paragraph_index
                else:
                    try:
                        page = vc.getPage()
                    except Exception:
                        page = 1
                    para_idx = None

                if page is None and para_idx is not None:
                    page = self._base.get_page_for_paragraph(
                        doc, para_idx)

            # Lock display + save cursor to prevent viewport jumping
            saved = doc.getText().createTextCursorByRange(vc.getStart())
            doc.lockControllers()
            try:
                images = self._scan_page_objects(doc, vc, page)
            finally:
                vc.gotoRange(saved, False)
                doc.unlockControllers()

            return {"success": True, "page": page, **images}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _scan_page_objects(self, doc, vc, page):
        """Scan page objects (called with controllers locked)."""
        images = []
        if hasattr(doc, "getGraphicObjects"):
            graphics = doc.getGraphicObjects()
            for name in graphics.getElementNames():
                try:
                    g = graphics.getByName(name)
                    anchor = g.getAnchor()
                    vc.gotoRange(anchor, False)
                    if vc.getPage() == page:
                        size = g.getPropertyValue("Size")
                        images.append({
                            "name": name,
                            "width_mm": size.Width // 100,
                            "height_mm": size.Height // 100,
                            "title": g.getPropertyValue("Title"),
                            "paragraph_index":
                                self._base.anchor_para_index(
                                    doc, anchor),
                        })
                except Exception:
                    pass

        tables = []
        if hasattr(doc, "getTextTables"):
            text_tables = doc.getTextTables()
            for name in text_tables.getElementNames():
                try:
                    t = text_tables.getByName(name)
                    anchor = t.getAnchor()
                    vc.gotoRange(anchor, False)
                    if vc.getPage() == page:
                        tables.append({
                            "name": name,
                            "rows": t.getRows().getCount(),
                            "cols": t.getColumns().getCount(),
                        })
                except Exception:
                    pass

        frames = []
        if hasattr(doc, "getTextFrames"):
            text_frames = doc.getTextFrames()
            frame_images = {}
            for img in images:
                iname = img["name"]
                try:
                    g = doc.getGraphicObjects().getByName(iname)
                    anchor_text = g.getAnchor().getText()
                    for fname in text_frames.getElementNames():
                        fr = text_frames.getByName(fname)
                        if fr.getText() == anchor_text:
                            frame_images.setdefault(
                                fname, []).append(iname)
                            break
                except Exception:
                    pass
            for fname in text_frames.getElementNames():
                try:
                    fr = text_frames.getByName(fname)
                    anchor = fr.getAnchor()
                    vc.gotoRange(anchor, False)
                    if vc.getPage() == page:
                        size = fr.getPropertyValue("Size")
                        entry = {
                            "name": fname,
                            "width_mm": size.Width // 100,
                            "height_mm": size.Height // 100,
                            "paragraph_index":
                                self._base.anchor_para_index(
                                    doc, anchor),
                        }
                        if fname in frame_images:
                            entry["images"] = frame_images[fname]
                        frames.append(entry)
                except Exception:
                    pass

        return {"images": images, "tables": tables, "frames": frames}

    # ==================================================================
    # Indexes & fields
    # ==================================================================

    def refresh_indexes(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getDocumentIndexes"):
                return {"success": False,
                        "error": "Document does not support indexes"}
            indexes = doc.getDocumentIndexes()
            count = indexes.getCount()
            refreshed = []
            for i in range(count):
                idx = indexes.getByIndex(i)
                idx.update()
                name = (idx.getName()
                        if hasattr(idx, "getName") else f"index_{i}")
                refreshed.append(name)
            if count > 0 and doc.hasLocation():
                self._base.store_doc(doc)
            return {"success": True, "refreshed": refreshed,
                    "count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def update_fields(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, "getTextFields"):
                return {"success": False,
                        "error": "Document does not support text fields"}
            fields = doc.getTextFields()
            fields.refresh()
            enum = fields.createEnumeration()
            count = 0
            while enum.hasMoreElements():
                enum.nextElement()
                count += 1
            return {"success": True, "fields_refreshed": count}
        except Exception as e:
            return {"success": False, "error": str(e)}
