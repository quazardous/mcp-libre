"""
ProximityService — local heading navigation and surroundings discovery.

Operates on the cached heading tree and bookmark map from TreeService.
Avoids full-document enumeration for heading navigation (O(log h) cached).
"""

import bisect
import logging
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class ProximityService:
    """Local proximity navigation on Writer documents."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base

        # Per-document flattened heading list cache
        self._flat_cache: Dict[str, List[Dict]] = {}

    def invalidate_cache(self, doc=None):
        """Clear flattened-tree cache."""
        if doc is None:
            self._flat_cache.clear()
        else:
            self._flat_cache.pop(self._base.doc_key(doc), None)

    # ==================================================================
    # Flattened tree (ordered heading list with parent pointers)
    # ==================================================================

    def _flatten_tree(self, root: Dict, doc) -> List[Dict]:
        """DFS-flatten heading tree into document-order list.

        Each entry: {"node": <tree_node>, "parent": <parent_entry|None>}
        Cached per document, invalidated alongside tree cache.
        """
        key = self._base.doc_key(doc)
        if key in self._flat_cache:
            return self._flat_cache[key]

        flat = []
        self._flatten_recurse(root["children"], None, flat)
        self._flat_cache[key] = flat
        return flat

    def _flatten_recurse(self, children: List[Dict],
                         parent_entry: Optional[Dict],
                         flat: List[Dict]):
        for child in children:
            entry = {"node": child, "parent": parent_entry}
            flat.append(entry)
            if child.get("children"):
                self._flatten_recurse(child["children"], entry, flat)

    # ==================================================================
    # Heading context (binary search)
    # ==================================================================

    def _find_heading_context(self, flat: List[Dict],
                               para_index: int
                               ) -> Tuple[Optional[int], Dict]:
        """Find the heading at or most recently before para_index.

        Returns (flat_index, info_dict).
        If no heading exists before para_index, returns (None, {...}).
        """
        if not flat:
            return None, {"was_heading": False}

        para_indexes = [e["node"]["para_index"] for e in flat]
        pos = bisect.bisect_right(para_indexes, para_index) - 1

        if pos < 0:
            return None, {"was_heading": False}

        entry = flat[pos]
        node = entry["node"]
        was_heading = (node["para_index"] == para_index)

        return pos, {
            "was_heading": was_heading,
            "heading_node": node,
            "heading_text": node.get("text", ""),
            "heading_level": node.get("level", 0),
        }

    # ==================================================================
    # Heading chain (ancestors)
    # ==================================================================

    def _get_heading_chain(self, flat: List[Dict],
                            ctx_idx: int,
                            bookmark_map: Dict[int, str]
                            ) -> List[Dict]:
        """Build ancestor chain from current position to root.

        Returns list from outermost (level 1) to innermost.
        """
        chain = []
        entry = flat[ctx_idx] if ctx_idx is not None else None
        while entry is not None:
            node = entry["node"]
            chain.append({
                "level": node["level"],
                "text": node["text"],
                "para_index": node["para_index"],
                "bookmark": bookmark_map.get(node["para_index"]),
            })
            entry = entry["parent"]
        chain.reverse()
        return chain

    # ==================================================================
    # Build heading result
    # ==================================================================

    def _build_heading_result(self, node: Dict,
                               bookmark_map: Dict[int, str]) -> Dict:
        return {
            "level": node["level"],
            "text": node["text"],
            "para_index": node["para_index"],
            "bookmark": bookmark_map.get(node["para_index"]),
            "body_paragraphs": node.get("body_paragraphs", 0),
            "children_count": len(node.get("children", [])),
        }

    # ==================================================================
    # navigate_heading
    # ==================================================================

    def navigate_heading(self, locator: str, direction: str,
                         file_path: str = None) -> Dict[str, Any]:
        """Navigate from locator to a related heading.

        Directions: next, previous, parent, first_child,
                    next_sibling, previous_sibling.
        """
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_writer(doc):
                return {"success": False, "error": "Not a Writer document"}

            resolved = self._base.resolve_locator(doc, locator)
            para_index = resolved.get("para_index")
            if para_index is None:
                return {"success": False,
                        "error": f"Cannot resolve locator: {locator}"}

            tree = self._writer.tree.build_heading_tree(doc)
            bookmark_map = self._writer.tree.ensure_heading_bookmarks(doc)
            flat = self._flatten_tree(tree, doc)

            if not flat:
                return {"success": False,
                        "error": "Document has no headings"}

            ctx_idx, ctx_info = self._find_heading_context(
                flat, para_index)

            from_info = {
                "para_index": para_index,
                "was_heading": ctx_info["was_heading"],
            }
            if ctx_idx is not None:
                ctx_node = flat[ctx_idx]["node"]
                from_info["context_heading"] = ctx_node["text"]
                from_info["context_level"] = ctx_node["level"]
                from_info["context_bookmark"] = bookmark_map.get(
                    ctx_node["para_index"])

            target_entry = None
            error_msg = None

            if direction == "next":
                next_idx = (ctx_idx + 1) if ctx_idx is not None else 0
                if next_idx < len(flat):
                    target_entry = flat[next_idx]
                else:
                    error_msg = "No next heading — already at last heading"

            elif direction == "previous":
                if ctx_idx is None:
                    error_msg = ("No previous heading — position is "
                                 "before any heading")
                elif not ctx_info["was_heading"] and ctx_idx >= 0:
                    target_entry = flat[ctx_idx]
                elif ctx_idx > 0:
                    target_entry = flat[ctx_idx - 1]
                else:
                    error_msg = ("No previous heading — already at "
                                 "first heading")

            elif direction == "parent":
                if ctx_idx is None:
                    error_msg = ("No parent heading — position is "
                                 "before any heading")
                else:
                    parent_entry = flat[ctx_idx]["parent"]
                    if parent_entry is not None:
                        target_entry = parent_entry
                    else:
                        error_msg = ("No parent heading — already at "
                                     "top level")

            elif direction == "first_child":
                if ctx_idx is None:
                    error_msg = ("No child headings — position is "
                                 "before any heading")
                else:
                    children = flat[ctx_idx]["node"].get("children", [])
                    if children:
                        child_pi = children[0]["para_index"]
                        for entry in flat:
                            if entry["node"]["para_index"] == child_pi:
                                target_entry = entry
                                break
                    if target_entry is None and error_msg is None:
                        error_msg = "No child headings under this heading"

            elif direction == "next_sibling":
                if ctx_idx is None:
                    error_msg = ("No sibling headings — position is "
                                 "before any heading")
                else:
                    target_entry = self._find_sibling(
                        flat, ctx_idx, offset=1)
                    if target_entry is None:
                        error_msg = "No next sibling heading"

            elif direction == "previous_sibling":
                if ctx_idx is None:
                    error_msg = ("No sibling headings — position is "
                                 "before any heading")
                else:
                    target_entry = self._find_sibling(
                        flat, ctx_idx, offset=-1)
                    if target_entry is None:
                        error_msg = "No previous sibling heading"

            else:
                return {"success": False,
                        "error": f"Unknown direction: {direction}. "
                                 "Use: next, previous, parent, "
                                 "first_child, next_sibling, "
                                 "previous_sibling"}

            if error_msg:
                return {"success": False, "error": error_msg,
                        "from": from_info}

            return {
                "success": True,
                "heading": self._build_heading_result(
                    target_entry["node"], bookmark_map),
                "from": from_info,
                "direction": direction,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _find_sibling(self, flat: List[Dict], ctx_idx: int,
                      offset: int) -> Optional[Dict]:
        """Find next (+1) or previous (-1) sibling."""
        entry = flat[ctx_idx]
        parent = entry["parent"]
        if parent is None:
            top_level = [e for e in flat if e["parent"] is None]
            for i, e in enumerate(top_level):
                if e is entry:
                    sib_idx = i + offset
                    if 0 <= sib_idx < len(top_level):
                        return top_level[sib_idx]
                    return None
            return None

        siblings = parent["node"].get("children", [])
        current_pi = entry["node"]["para_index"]
        for i, sib in enumerate(siblings):
            if sib["para_index"] == current_pi:
                sib_idx = i + offset
                if 0 <= sib_idx < len(siblings):
                    target_pi = siblings[sib_idx]["para_index"]
                    for e in flat:
                        if e["node"]["para_index"] == target_pi:
                            return e
                return None
        return None

    # ==================================================================
    # get_surroundings
    # ==================================================================

    def get_surroundings(self, locator: str, radius: int = 10,
                         include: List[str] = None,
                         file_path: str = None) -> Dict[str, Any]:
        """Discover objects within radius paragraphs of locator."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_writer(doc):
                return {"success": False, "error": "Not a Writer document"}

            resolved = self._base.resolve_locator(doc, locator)
            center_idx = resolved.get("para_index")
            if center_idx is None:
                return {"success": False,
                        "error": f"Cannot resolve locator: {locator}"}

            radius = max(1, min(50, radius))
            if include is None:
                include = ["paragraphs", "images", "tables",
                           "frames", "comments", "headings"]

            para_ranges = self._writer.get_paragraph_ranges(doc)
            total = len(para_ranges)

            if center_idx >= total:
                return {"success": False,
                        "error": f"Paragraph {center_idx} out of range "
                                 f"(document has {total} paragraphs)"}

            start_idx = max(0, center_idx - radius)
            end_idx = min(total - 1, center_idx + radius)
            text_obj = doc.getText()

            result = {
                "success": True,
                "center_para_index": center_idx,
                "range": {"start": start_idx, "end": end_idx},
            }

            # Heading chain
            bookmark_map = {}
            if "headings" in include:
                tree = self._writer.tree.build_heading_tree(doc)
                bookmark_map = (
                    self._writer.tree.ensure_heading_bookmarks(doc))
                flat = self._flatten_tree(tree, doc)
                ctx_idx, _ = self._find_heading_context(
                    flat, center_idx)
                result["heading_chain"] = self._get_heading_chain(
                    flat, ctx_idx, bookmark_map)

            # Paragraphs
            if "paragraphs" in include:
                if not bookmark_map:
                    bookmark_map = (
                        self._writer.tree.get_mcp_bookmark_map(doc))
                paragraphs = []
                for i in range(start_idx, end_idx + 1):
                    el = para_ranges[i]
                    if not el.supportsService(
                            "com.sun.star.text.Paragraph"):
                        continue
                    level = 0
                    try:
                        level = el.getPropertyValue("OutlineLevel")
                    except Exception:
                        pass
                    text = el.getString()
                    entry = {
                        "para_index": i,
                        "text": text[:200] if len(text) > 200
                        else text,
                        "outline_level": level,
                    }
                    bm = bookmark_map.get(i)
                    if bm:
                        entry["bookmark"] = bm
                    paragraphs.append(entry)
                result["paragraphs"] = paragraphs

            # Images
            if "images" in include and hasattr(doc, "getGraphicObjects"):
                images = []
                graphics = doc.getGraphicObjects()
                for name in graphics.getElementNames():
                    try:
                        g = graphics.getByName(name)
                        anchor = g.getAnchor()
                        pi = self._writer.find_paragraph_for_range(
                            anchor, para_ranges, text_obj)
                        if start_idx <= pi <= end_idx:
                            size = g.getPropertyValue("Size")
                            images.append({
                                "name": name,
                                "paragraph_index": pi,
                                "title": g.getPropertyValue("Title"),
                                "width_mm": size.Width // 100,
                                "height_mm": size.Height // 100,
                            })
                    except Exception:
                        pass
                result["images"] = images

            # Tables
            if "tables" in include and hasattr(doc, "getTextTables"):
                tables = []
                text_tables = doc.getTextTables()
                for name in text_tables.getElementNames():
                    try:
                        t = text_tables.getByName(name)
                        anchor = t.getAnchor()
                        pi = self._writer.find_paragraph_for_range(
                            anchor, para_ranges, text_obj)
                        if start_idx <= pi <= end_idx:
                            tables.append({
                                "name": name,
                                "paragraph_index": pi,
                                "rows": t.getRows().getCount(),
                                "cols": t.getColumns().getCount(),
                            })
                    except Exception:
                        pass
                result["tables"] = tables

            # Frames
            if "frames" in include and hasattr(doc, "getTextFrames"):
                frames = []
                text_frames = doc.getTextFrames()
                for fname in text_frames.getElementNames():
                    try:
                        fr = text_frames.getByName(fname)
                        anchor = fr.getAnchor()
                        pi = self._writer.find_paragraph_for_range(
                            anchor, para_ranges, text_obj)
                        if start_idx <= pi <= end_idx:
                            size = fr.getPropertyValue("Size")
                            frames.append({
                                "name": fname,
                                "paragraph_index": pi,
                                "width_mm": size.Width // 100,
                                "height_mm": size.Height // 100,
                            })
                    except Exception:
                        pass
                result["frames"] = frames

            # Comments
            if "comments" in include:
                comments = []
                try:
                    fields = doc.getTextFields()
                    enum = fields.createEnumeration()
                    while enum.hasMoreElements():
                        field = enum.nextElement()
                        if not field.supportsService(
                                "com.sun.star.text.textfield"
                                ".Annotation"):
                            continue
                        try:
                            parent_name = field.getPropertyValue(
                                "ParentName")
                            if parent_name:
                                continue
                        except Exception:
                            pass
                        try:
                            author = field.getPropertyValue("Author")
                        except Exception:
                            author = ""
                        if author == "MCP-AI":
                            continue
                        anchor = field.getAnchor()
                        pi = self._writer.find_paragraph_for_range(
                            anchor, para_ranges, text_obj)
                        if start_idx <= pi <= end_idx:
                            content = field.getPropertyValue("Content")
                            name = ""
                            try:
                                name = field.getPropertyValue("Name")
                            except Exception:
                                pass
                            resolved_flag = False
                            try:
                                resolved_flag = field.getPropertyValue(
                                    "Resolved")
                            except Exception:
                                pass
                            comments.append({
                                "author": author,
                                "content": content[:200],
                                "comment_name": name,
                                "paragraph_index": pi,
                                "resolved": resolved_flag,
                            })
                except Exception:
                    pass
                result["comments"] = comments

            return result
        except Exception as e:
            return {"success": False, "error": str(e)}