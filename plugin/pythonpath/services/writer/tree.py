"""
TreeService — heading tree, bookmarks, content strategies, AI annotations.

Includes per-document caching for heading tree, bookmark map,
and AI summaries map.  Caches are invalidated by WriterService
after any document edit.
"""

import logging
import uuid
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class TreeService:
    """Heading tree navigation with per-document caching."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base

        # Per-document caches: {doc_key: value}
        self._tree_cache: Dict[str, Dict] = {}
        self._bookmark_cache: Dict[str, Dict[int, str]] = {}
        self._ai_summary_cache: Dict[str, Dict[int, str]] = {}

    def invalidate_cache(self, doc=None):
        """Clear caches (all or for a specific document)."""
        if doc is None:
            self._tree_cache.clear()
            self._bookmark_cache.clear()
            self._ai_summary_cache.clear()
        else:
            key = self._base.doc_key(doc)
            self._tree_cache.pop(key, None)
            self._bookmark_cache.pop(key, None)
            self._ai_summary_cache.pop(key, None)

    # ==================================================================
    # Heading bookmarks (stable IDs)
    # ==================================================================

    def get_mcp_bookmark_map(self, doc) -> Dict[int, str]:
        """Get {para_index: bookmark_name} for all _mcp_ bookmarks."""
        key = self._base.doc_key(doc)
        if key in self._bookmark_cache:
            return self._bookmark_cache[key]

        result = {}
        try:
            if not hasattr(doc, "getBookmarks"):
                return result
            bookmarks = doc.getBookmarks()
            names = bookmarks.getElementNames()
            if not names:
                return result
            para_ranges = self._writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()
            for name in names:
                if not name.startswith("_mcp_"):
                    continue
                bm = bookmarks.getByName(name)
                anchor = bm.getAnchor()
                para_idx = self._writer.find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                result[para_idx] = name
        except Exception as e:
            logger.error("Failed to get MCP bookmark map: %s", e)

        self._bookmark_cache[key] = result
        return result

    def ensure_heading_bookmarks(self, doc) -> Dict[int, str]:
        """Ensure every heading has an _mcp_ bookmark. Returns map."""
        existing_map = self.get_mcp_bookmark_map(doc)
        text = doc.getText()
        enum = text.createEnumeration()
        para_index = 0
        bookmark_map = {}
        needs_bookmark = []

        while enum.hasMoreElements():
            element = enum.nextElement()
            if element.supportsService("com.sun.star.text.Paragraph"):
                outline_level = 0
                try:
                    outline_level = element.getPropertyValue("OutlineLevel")
                except Exception:
                    pass
                if outline_level > 0:
                    if para_index in existing_map:
                        bookmark_map[para_index] = existing_map[para_index]
                    else:
                        needs_bookmark.append(
                            (para_index, element.getStart()))
            para_index += 1
            self._base.yield_to_gui()

        for para_idx, start_range in needs_bookmark:
            bm_name = f"_mcp_{uuid.uuid4().hex[:8]}"
            bookmark = doc.createInstance("com.sun.star.text.Bookmark")
            bookmark.Name = bm_name
            cursor = text.createTextCursorByRange(start_range)
            text.insertTextContent(cursor, bookmark, False)
            bookmark_map[para_idx] = bm_name

        if needs_bookmark and doc.hasLocation():
            self._base.store_doc(doc)

        # Update cache
        key = self._base.doc_key(doc)
        self._bookmark_cache[key] = bookmark_map
        return bookmark_map

    def find_nearest_heading_bookmark(self, para_index: int,
                                      bookmark_map: Dict[int, str]
                                      ) -> Optional[Dict[str, Any]]:
        """Find nearest heading bookmark at or before para_index."""
        best_idx = -1
        for idx in bookmark_map:
            if idx <= para_index and idx > best_idx:
                best_idx = idx
        if best_idx >= 0:
            return {"bookmark": bookmark_map[best_idx],
                    "heading_para_index": best_idx}
        return None

    # ==================================================================
    # Tree building
    # ==================================================================

    def build_heading_tree(self, doc) -> Dict[str, Any]:
        """Build heading tree from paragraph enumeration. Single pass."""
        key = self._base.doc_key(doc)
        if key in self._tree_cache:
            return self._tree_cache[key]

        text = doc.getText()
        enum = text.createEnumeration()
        root = {"level": 0, "text": "root", "para_index": -1,
                "children": [], "body_paragraphs": 0}
        stack = [root]
        para_index = 0

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService(
                "com.sun.star.text.Paragraph")
            is_table = element.supportsService(
                "com.sun.star.text.TextTable")

            if is_para:
                outline_level = 0
                try:
                    outline_level = element.getPropertyValue("OutlineLevel")
                except Exception:
                    pass
                if outline_level > 0:
                    while (len(stack) > 1
                           and stack[-1]["level"] >= outline_level):
                        stack.pop()
                    node = {"level": outline_level,
                            "text": element.getString(),
                            "para_index": para_index,
                            "children": [], "body_paragraphs": 0}
                    stack[-1]["children"].append(node)
                    stack.append(node)
                else:
                    stack[-1]["body_paragraphs"] += 1
            elif is_table:
                stack[-1]["body_paragraphs"] += 1

            para_index += 1
            self._base.yield_to_gui()

        self._tree_cache[key] = root
        return root

    def _count_all_children(self, node: Dict) -> int:
        count = len(node.get("children", []))
        for child in node.get("children", []):
            if "children" in child:
                count += self._count_all_children(child)
        return count + node.get("body_paragraphs", 0)

    def _find_node_by_para_index(self, node: Dict,
                                  para_index: int) -> Optional[Dict]:
        if node.get("para_index") == para_index:
            return node
        for child in node.get("children", []):
            found = self._find_node_by_para_index(child, para_index)
            if found is not None:
                return found
        return None

    # ==================================================================
    # Content strategies
    # ==================================================================

    def _get_body_preview(self, doc, heading_para_index: int,
                          max_chars: int = 100) -> str:
        text = doc.getText()
        enum = text.createEnumeration()
        idx = 0
        preview_parts = []
        found_heading = (heading_para_index == -1)

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService(
                "com.sun.star.text.Paragraph")
            if idx == heading_para_index:
                found_heading = True
                idx += 1
                continue
            if found_heading and is_para:
                outline_level = 0
                try:
                    outline_level = element.getPropertyValue("OutlineLevel")
                except Exception:
                    pass
                if outline_level > 0:
                    break
                para_text = element.getString().strip()
                if para_text:
                    preview_parts.append(para_text)
                    if sum(len(p) for p in preview_parts) >= max_chars:
                        break
            idx += 1

        full_preview = " ".join(preview_parts)
        if len(full_preview) > max_chars:
            full_preview = full_preview[:max_chars] + "..."
        return full_preview

    def _get_full_body_text(self, doc, heading_para_index: int) -> str:
        text = doc.getText()
        enum = text.createEnumeration()
        idx = 0
        parts = []
        found_heading = (heading_para_index == -1)

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService(
                "com.sun.star.text.Paragraph")
            if idx == heading_para_index:
                found_heading = True
                idx += 1
                continue
            if found_heading and is_para:
                outline_level = 0
                try:
                    outline_level = element.getPropertyValue("OutlineLevel")
                except Exception:
                    pass
                if outline_level > 0:
                    break
                parts.append(element.getString())
            idx += 1

        return "\n".join(parts)

    def get_ai_summaries_map(self, doc) -> Dict[int, str]:
        """Build {para_index: summary} map from MCP-AI annotations."""
        key = self._base.doc_key(doc)
        if key in self._ai_summary_cache:
            return self._ai_summary_cache[key]

        summaries = {}
        try:
            fields_supplier = doc.getTextFields()
            enum = fields_supplier.createEnumeration()
            para_ranges = self._writer.get_paragraph_ranges(doc)

            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    continue
                if author != "MCP-AI":
                    continue
                content = field.getPropertyValue("Content")
                anchor = field.getAnchor()
                para_idx = self._writer.find_paragraph_for_range(
                    anchor, para_ranges, doc.getText())
                summaries[para_idx] = content
        except Exception as e:
            logger.error("Failed to get AI summaries: %s", e)

        self._ai_summary_cache[key] = summaries
        return summaries

    def _apply_content_strategy(self, node: Dict, doc,
                                ai_summaries: Dict[int, str],
                                strategy: str,
                                max_chars: int = 100):
        para_idx = node.get("para_index", -1)
        if strategy == "none":
            pass
        elif strategy == "ai_summary_first":
            if para_idx in ai_summaries:
                node["ai_summary"] = ai_summaries[para_idx]
            else:
                node["body_preview"] = self._get_body_preview(
                    doc, para_idx, max_chars)
        elif strategy == "first_lines":
            node["body_preview"] = self._get_body_preview(
                doc, para_idx, max_chars)
            if para_idx in ai_summaries:
                node["ai_summary"] = ai_summaries[para_idx]
        elif strategy == "full":
            node["body_text"] = self._get_full_body_text(doc, para_idx)

    def _serialize_tree_node(self, child: Dict, doc,
                              ai_summaries: Dict[int, str],
                              content_strategy: str, depth: int,
                              current_depth: int = 1,
                              bookmark_map: Dict[int, str] = None
                              ) -> Dict[str, Any]:
        node = {
            "type": "heading",
            "level": child["level"],
            "text": child["text"],
            "para_index": child["para_index"],
            "bookmark": (bookmark_map or {}).get(child["para_index"]),
            "children_count": self._count_all_children(child),
            "body_paragraphs": child["body_paragraphs"],
        }
        self._apply_content_strategy(
            node, doc, ai_summaries, content_strategy)
        if depth == 0 or current_depth < depth:
            if child.get("children"):
                node["children"] = [
                    self._serialize_tree_node(
                        sub, doc, ai_summaries, content_strategy,
                        depth, current_depth + 1, bookmark_map)
                    for sub in child["children"]
                ]
        return node

    # ==================================================================
    # Public tree API
    # ==================================================================

    def get_document_tree(self, content_strategy: str = "first_lines",
                          depth: int = 1,
                          file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_writer(doc):
                return {"success": False, "error": "Not a Writer document"}

            tree = self.build_heading_tree(doc)
            bookmark_map = self.ensure_heading_bookmarks(doc)
            ai_summaries = (
                self.get_ai_summaries_map(doc)
                if content_strategy in ("ai_summary_first", "first_lines")
                else {})

            children = [
                self._serialize_tree_node(
                    child, doc, ai_summaries, content_strategy,
                    depth, bookmark_map=bookmark_map)
                for child in tree["children"]
            ]

            text = doc.getText()
            enum = text.createEnumeration()
            total = 0
            while enum.hasMoreElements():
                enum.nextElement()
                total += 1

            try:
                self._base.annotate_pages(children, doc)
            except Exception:
                pass

            page_count = 0
            try:
                cursor = text.createTextCursor()
                cursor.gotoEnd(False)
                page_count = self._base.get_page_for_range(doc, cursor)
            except Exception:
                pass

            return {
                "success": True,
                "content_strategy": content_strategy,
                "depth": depth,
                "children": children,
                "body_before_first_heading": tree["body_paragraphs"],
                "total_paragraphs": total,
                "page_count": page_count,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_heading_children(self, heading_para_index: int = None,
                             heading_bookmark: str = None,
                             locator: str = None,
                             content_strategy: str = "first_lines",
                             depth: int = 1,
                             file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_writer(doc):
                return {"success": False, "error": "Not a Writer document"}

            if locator is not None and heading_para_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                heading_para_index = resolved.get("para_index")
            elif heading_bookmark is not None and heading_para_index is None:
                if not hasattr(doc, "getBookmarks"):
                    return {"success": False,
                            "error": "Document doesn't support bookmarks"}
                bm_sup = doc.getBookmarks()
                if not bm_sup.hasByName(heading_bookmark):
                    return {"success": False,
                            "error": f"Bookmark '{heading_bookmark}' not found"}
                bm = bm_sup.getByName(heading_bookmark)
                anchor = bm.getAnchor()
                para_ranges = self._writer.get_paragraph_ranges(doc)
                heading_para_index = self._writer.find_paragraph_for_range(
                    anchor, para_ranges, doc.getText())

            if heading_para_index is None:
                return {"success": False,
                        "error": "Provide locator, heading_para_index, "
                                 "or heading_bookmark"}

            tree = self.build_heading_tree(doc)
            bookmark_map = self.ensure_heading_bookmarks(doc)
            target = self._find_node_by_para_index(
                tree, heading_para_index)
            if target is None:
                return {"success": False,
                        "error": f"Heading at paragraph "
                                 f"{heading_para_index} not found"}

            ai_summaries = (
                self.get_ai_summaries_map(doc)
                if content_strategy in ("ai_summary_first", "first_lines")
                else {})

            children = []
            text = doc.getText()
            enum = text.createEnumeration()
            idx = 0
            found_heading = False
            parent_level = target["level"]

            while enum.hasMoreElements():
                element = enum.nextElement()
                is_para = element.supportsService(
                    "com.sun.star.text.Paragraph")
                if idx == heading_para_index:
                    found_heading = True
                    idx += 1
                    continue
                if found_heading and is_para:
                    outline_level = 0
                    try:
                        outline_level = element.getPropertyValue(
                            "OutlineLevel")
                    except Exception:
                        pass
                    if outline_level > 0 and outline_level <= parent_level:
                        break
                    if outline_level > 0:
                        break
                    para_text = element.getString()
                    preview = (para_text[:100] + "..."
                               if len(para_text) > 100 else para_text)
                    if content_strategy == "full":
                        children.append({"type": "body",
                                         "para_index": idx,
                                         "text": para_text})
                    elif content_strategy != "none":
                        children.append({"type": "body",
                                         "para_index": idx,
                                         "preview": preview})
                    else:
                        children.append({"type": "body",
                                         "para_index": idx})
                idx += 1
                self._base.yield_to_gui()

            for child in target["children"]:
                node = self._serialize_tree_node(
                    child, doc, ai_summaries, content_strategy,
                    depth, bookmark_map=bookmark_map)
                children.append(node)

            return {
                "success": True,
                "parent": {
                    "level": target["level"],
                    "text": target["text"],
                    "para_index": target["para_index"],
                    "bookmark": bookmark_map.get(target["para_index"]),
                },
                "content_strategy": content_strategy,
                "depth": depth,
                "children": children,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # AI annotations
    # ==================================================================

    def add_ai_summary(self, para_index: int = None, summary: str = "",
                       locator: str = None,
                       file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and para_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                para_index = resolved.get("para_index")
            if para_index is None:
                return {"success": False,
                        "error": "Provide locator or para_index"}

            doc_text = doc.getText()
            self._remove_ai_annotation_at(doc, para_index)

            target, _ = self._writer.find_paragraph_element(
                doc, para_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {para_index} not found"}

            annotation = doc.createInstance(
                "com.sun.star.text.textfield.Annotation")
            annotation.setPropertyValue("Author", "MCP-AI")
            annotation.setPropertyValue("Content", summary)
            cursor = doc_text.createTextCursorByRange(target.getStart())
            doc_text.insertTextContent(cursor, annotation, False)

            # Invalidate AI summary cache
            self._ai_summary_cache.pop(
                self._base.doc_key(doc), None)

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "message": f"Added AI summary at paragraph {para_index}",
                    "para_index": para_index,
                    "summary_length": len(summary)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_ai_summaries(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            summaries_map = self.get_ai_summaries_map(doc)
            summaries = [{"para_index": idx, "summary": text}
                         for idx, text in sorted(summaries_map.items())]
            return {"success": True, "summaries": summaries,
                    "count": len(summaries)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def remove_ai_summary(self, para_index: int = None,
                          locator: str = None,
                          file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and para_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                para_index = resolved.get("para_index")
            if para_index is None:
                return {"success": False,
                        "error": "Provide locator or para_index"}
            removed = self._remove_ai_annotation_at(doc, para_index)
            # Invalidate AI summary cache
            self._ai_summary_cache.pop(
                self._base.doc_key(doc), None)
            if removed and doc.hasLocation():
                self._base.store_doc(doc)
            return {"success": True, "removed": removed,
                    "para_index": para_index}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _remove_ai_annotation_at(self, doc, para_index: int) -> bool:
        try:
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            para_ranges = self._writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()
            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    continue
                if author != "MCP-AI":
                    continue
                anchor = field.getAnchor()
                idx = self._writer.find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                if idx == para_index:
                    text_obj.removeTextContent(field)
                    return True
        except Exception as e:
            logger.error("Failed to remove AI annotation: %s", e)
        return False
