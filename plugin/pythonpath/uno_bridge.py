"""
LibreOffice MCP Extension - UNO Bridge Module

This module provides a bridge between MCP operations and LibreOffice UNO API,
enabling direct manipulation of LibreOffice documents.
"""

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.text import XTextDocument
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.sheet import XSpreadsheetDocument
try:
    from com.sun.star.presentation import XPresentationDocument
except ImportError:
    XPresentationDocument = None
from com.sun.star.document import XDocumentEventListener
from com.sun.star.awt import XActionListener
from typing import Any, Optional, Dict, List, Tuple
import logging
import traceback
import uuid

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class UNOBridge:
    """Bridge between MCP operations and LibreOffice UNO API"""
    
    def __init__(self):
        """Initialize the UNO bridge"""
        try:
            self.ctx = uno.getComponentContext()
            self.smgr = self.ctx.ServiceManager
            self.desktop = self.smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            logger.info("UNO Bridge initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize UNO Bridge: {e}")
            raise
    
    def create_document(self, doc_type: str = "writer") -> Any:
        """
        Create new document using UNO API
        
        Args:
            doc_type: Type of document ('writer', 'calc', 'impress', 'draw')
            
        Returns:
            Document object
        """
        try:
            url_map = {
                "writer": "private:factory/swriter",
                "calc": "private:factory/scalc", 
                "impress": "private:factory/simpress",
                "draw": "private:factory/sdraw"
            }
            
            url = url_map.get(doc_type, "private:factory/swriter")
            doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
            logger.info(f"Created new {doc_type} document")
            return doc
            
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            raise
    
    def get_active_document(self) -> Optional[Any]:
        """Get currently active document"""
        try:
            doc = self.desktop.getCurrentComponent()
            if doc:
                logger.info("Retrieved active document")
            return doc
        except Exception as e:
            logger.error(f"Failed to get active document: {e}")
            return None
    
    # ----------------------------------------------------------------
    # Document resolution
    # ----------------------------------------------------------------

    def _find_open_document(self, file_url: str) -> Optional[Any]:
        """Find an already-open document by its URL."""
        try:
            components = self.desktop.getComponents()
            if components is None:
                return None
            enum = components.createEnumeration()
            while enum.hasMoreElements():
                doc = enum.nextElement()
                if hasattr(doc, 'getURL') and doc.getURL() == file_url:
                    return doc
            return None
        except Exception:
            return None

    def open_document(self, file_path: str) -> Dict[str, Any]:
        """Open a document by file path, or return it if already open."""
        try:
            file_url = uno.systemPathToFileUrl(file_path)

            existing = self._find_open_document(file_url)
            if existing is not None:
                return {"success": True, "doc": existing,
                        "url": file_url, "already_open": True}

            props = (
                PropertyValue("Hidden", 0, False, 0),
                PropertyValue("ReadOnly", 0, False, 0),
            )
            doc = self.desktop.loadComponentFromURL(file_url, "_blank", 0, props)
            if doc is None:
                return {"success": False,
                        "error": f"Failed to load document: {file_path}"}

            logger.info(f"Opened document: {file_path}")
            return {"success": True, "doc": doc,
                    "url": file_url, "already_open": False}
        except Exception as e:
            logger.error(f"Failed to open document: {e}")
            return {"success": False, "error": str(e)}

    def _resolve_document(self, file_path: str = None) -> Any:
        """Resolve a document: open by path or use active document."""
        if file_path:
            result = self.open_document(file_path)
            if not result["success"]:
                raise RuntimeError(result["error"])
            return result["doc"]
        doc = self.get_active_document()
        if doc is None:
            raise RuntimeError("No active document and no file path provided")
        return doc

    # ----------------------------------------------------------------
    # Heading bookmark management (stable IDs)
    # ----------------------------------------------------------------

    def _get_mcp_bookmark_map(self, doc) -> Dict[int, str]:
        """Get {para_index: bookmark_name} for all _mcp_ bookmarks."""
        result = {}
        try:
            if not hasattr(doc, 'getBookmarks'):
                return result
            bookmarks = doc.getBookmarks()
            names = bookmarks.getElementNames()
            if not names:
                return result
            para_ranges = self._get_paragraph_ranges(doc)
            text_obj = doc.getText()
            for name in names:
                if not name.startswith("_mcp_"):
                    continue
                bm = bookmarks.getByName(name)
                anchor = bm.getAnchor()
                para_idx = self._find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                result[para_idx] = name
        except Exception as e:
            logger.error(f"Failed to get MCP bookmark map: {e}")
        return result

    def _ensure_heading_bookmarks(self, doc) -> Dict[int, str]:
        """Ensure every heading has an _mcp_ bookmark. Returns {para_index: name}."""
        existing_map = self._get_mcp_bookmark_map(doc)
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

        for para_idx, start_range in needs_bookmark:
            bm_name = f"_mcp_{uuid.uuid4().hex[:8]}"
            bookmark = doc.createInstance("com.sun.star.text.Bookmark")
            bookmark.Name = bm_name
            cursor = text.createTextCursorByRange(start_range)
            text.insertTextContent(cursor, bookmark, False)
            bookmark_map[para_idx] = bm_name

        if needs_bookmark and doc.hasLocation():
            doc.store()

        return bookmark_map

    def _find_nearest_heading_bookmark(self, para_index: int,
                                       bookmark_map: Dict[int, str]
                                       ) -> Optional[Dict[str, Any]]:
        """Find the nearest heading bookmark at or before para_index."""
        best_idx = -1
        for idx in bookmark_map:
            if idx <= para_index and idx > best_idx:
                best_idx = idx
        if best_idx >= 0:
            return {"bookmark": bookmark_map[best_idx],
                    "heading_para_index": best_idx}
        return None

    def resolve_bookmark(self, bookmark_name: str,
                         file_path: str = None) -> Dict[str, Any]:
        """Resolve a bookmark name to its current paragraph index."""
        try:
            doc = self._resolve_document(file_path)
            if not hasattr(doc, 'getBookmarks'):
                return {"success": False,
                        "error": "Document doesn't support bookmarks"}

            bookmarks = doc.getBookmarks()
            if not bookmarks.hasByName(bookmark_name):
                return {"success": False,
                        "error": f"Bookmark '{bookmark_name}' not found"}

            bm = bookmarks.getByName(bookmark_name)
            anchor = bm.getAnchor()
            para_ranges = self._get_paragraph_ranges(doc)
            text_obj = doc.getText()
            para_idx = self._find_paragraph_for_range(
                anchor, para_ranges, text_obj)

            # Get heading info
            heading_info = {}
            enum = text_obj.createEnumeration()
            idx = 0
            while enum.hasMoreElements():
                element = enum.nextElement()
                if idx == para_idx:
                    if element.supportsService("com.sun.star.text.Paragraph"):
                        try:
                            heading_info["text"] = element.getString()
                            heading_info["outline_level"] = \
                                element.getPropertyValue("OutlineLevel")
                        except Exception:
                            pass
                    break
                idx += 1

            return {
                "success": True,
                "bookmark": bookmark_name,
                "para_index": para_idx,
                **heading_info
            }
        except Exception as e:
            logger.error(f"Failed to resolve bookmark: {e}")
            return {"success": False, "error": str(e)}

    # ----------------------------------------------------------------
    # Tree navigation
    # ----------------------------------------------------------------

    def _build_heading_tree(self, doc) -> Dict[str, Any]:
        """Build heading tree from paragraph enumeration. Single pass."""
        text = doc.getText()
        enum = text.createEnumeration()

        root = {"level": 0, "text": "root", "para_index": -1,
                "children": [], "body_paragraphs": 0}
        stack = [root]
        para_index = 0

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService("com.sun.star.text.Paragraph")
            is_table = element.supportsService("com.sun.star.text.TextTable")

            if is_para:
                outline_level = 0
                try:
                    outline_level = element.getPropertyValue("OutlineLevel")
                except Exception:
                    pass

                if outline_level > 0:
                    while len(stack) > 1 and stack[-1]["level"] >= outline_level:
                        stack.pop()
                    node = {
                        "level": outline_level,
                        "text": element.getString(),
                        "para_index": para_index,
                        "children": [],
                        "body_paragraphs": 0
                    }
                    stack[-1]["children"].append(node)
                    stack.append(node)
                else:
                    stack[-1]["body_paragraphs"] += 1
            elif is_table:
                stack[-1]["body_paragraphs"] += 1

            para_index += 1

        return root

    def _count_all_children(self, node: Dict) -> int:
        """Recursively count all descendants of a node."""
        count = len(node.get("children", []))
        for child in node.get("children", []):
            if "children" in child:
                count += self._count_all_children(child)
        return count + node.get("body_paragraphs", 0)

    def _find_node_by_para_index(self, node: Dict, para_index: int) -> Optional[Dict]:
        """Find a heading node in the tree by its paragraph index."""
        if node.get("para_index") == para_index:
            return node
        for child in node.get("children", []):
            found = self._find_node_by_para_index(child, para_index)
            if found is not None:
                return found
        return None

    def _get_body_preview(self, doc, heading_para_index: int,
                          max_chars: int = 100) -> str:
        """Get a preview of body text following a heading."""
        text = doc.getText()
        enum = text.createEnumeration()
        idx = 0
        preview_parts = []
        found_heading = (heading_para_index == -1)  # -1 means root (start)

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService("com.sun.star.text.Paragraph")

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
                    break  # next heading reached
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

    def _get_ai_summaries_map(self, doc) -> Dict[int, str]:
        """Build {para_index: summary} map from MCP-AI annotations."""
        summaries = {}
        try:
            fields_supplier = doc.getTextFields()
            enum = fields_supplier.createEnumeration()

            # Build paragraph ranges for position lookup
            para_ranges = self._get_paragraph_ranges(doc)

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

                # Find which paragraph this annotation is in
                para_idx = self._find_paragraph_for_range(
                    anchor, para_ranges, doc.getText())
                summaries[para_idx] = content
        except Exception as e:
            logger.error(f"Failed to get AI summaries: {e}")
        return summaries

    def _get_paragraph_ranges(self, doc) -> List[Any]:
        """Get list of paragraph elements for range comparison."""
        text = doc.getText()
        enum = text.createEnumeration()
        ranges = []
        while enum.hasMoreElements():
            element = enum.nextElement()
            ranges.append(element)
        return ranges

    def _find_paragraph_for_range(self, match_range, para_ranges: List,
                                  text_obj=None) -> int:
        """Find which paragraph index a text range belongs to."""
        try:
            if text_obj is None:
                text_obj = match_range.getText()
            match_start = match_range.getStart()

            for i, para in enumerate(para_ranges):
                try:
                    para_start = para.getStart()
                    para_end = para.getEnd()
                    cmp_start = text_obj.compareRegionStarts(
                        match_start, para_start)
                    cmp_end = text_obj.compareRegionStarts(
                        match_start, para_end)
                    if cmp_start <= 0 and cmp_end >= 0:
                        return i
                except Exception:
                    continue
        except Exception:
            pass
        return 0

    def _apply_content_strategy(self, node: Dict, doc,
                                ai_summaries: Dict[int, str],
                                strategy: str,
                                max_chars: int = 100):
        """Add content to a tree node based on strategy."""
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

    def _get_full_body_text(self, doc, heading_para_index: int) -> str:
        """Get full body text following a heading until next heading."""
        text = doc.getText()
        enum = text.createEnumeration()
        idx = 0
        parts = []
        found_heading = (heading_para_index == -1)

        while enum.hasMoreElements():
            element = enum.nextElement()
            is_para = element.supportsService("com.sun.star.text.Paragraph")

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

    def _serialize_tree_node(self, child: Dict, doc,
                             ai_summaries: Dict[int, str],
                             content_strategy: str,
                             depth: int,
                             current_depth: int = 1,
                             bookmark_map: Dict[int, str] = None) -> Dict[str, Any]:
        """Serialize a heading tree node to the given depth."""
        node = {
            "type": "heading",
            "level": child["level"],
            "text": child["text"],
            "para_index": child["para_index"],
            "bookmark": (bookmark_map or {}).get(child["para_index"]),
            "children_count": self._count_all_children(child),
            "body_paragraphs": child["body_paragraphs"]
        }
        self._apply_content_strategy(
            node, doc, ai_summaries, content_strategy)

        # Recurse into sub-headings if depth allows
        # depth=0 means unlimited, depth=1 means just this level, etc.
        if depth == 0 or current_depth < depth:
            if child.get("children"):
                node["children"] = [
                    self._serialize_tree_node(
                        sub, doc, ai_summaries, content_strategy,
                        depth, current_depth + 1, bookmark_map)
                    for sub in child["children"]
                ]
        return node

    def get_document_tree(self, content_strategy: str = "first_lines",
                          depth: int = 1,
                          file_path: str = None) -> Dict[str, Any]:
        """Get the document heading tree.

        Args:
            content_strategy: none, first_lines, ai_summary_first, full
            depth: How many levels to return (1=direct children only,
                   2=children+grandchildren, 0=unlimited)
            file_path: Optional file path
        """
        try:
            doc = self._resolve_document(file_path)
            if not isinstance(doc, XTextDocument):
                return {"success": False,
                        "error": "Not a Writer document"}

            tree = self._build_heading_tree(doc)
            bookmark_map = self._ensure_heading_bookmarks(doc)
            ai_summaries = (self._get_ai_summaries_map(doc)
                            if content_strategy in ("ai_summary_first",
                                                    "first_lines")
                            else {})

            children = [
                self._serialize_tree_node(
                    child, doc, ai_summaries, content_strategy, depth,
                    bookmark_map=bookmark_map)
                for child in tree["children"]
            ]

            # Count total paragraphs
            text = doc.getText()
            enum = text.createEnumeration()
            total = 0
            while enum.hasMoreElements():
                enum.nextElement()
                total += 1

            return {
                "success": True,
                "content_strategy": content_strategy,
                "depth": depth,
                "children": children,
                "body_before_first_heading": tree["body_paragraphs"],
                "total_paragraphs": total
            }
        except Exception as e:
            logger.error(f"Failed to get document tree: {e}")
            return {"success": False, "error": str(e)}

    def get_heading_children(self, heading_para_index: int = None,
                             heading_bookmark: str = None,
                             content_strategy: str = "first_lines",
                             depth: int = 1,
                             file_path: str = None) -> Dict[str, Any]:
        """Get children of a heading node.

        Args:
            heading_para_index: Paragraph index of the parent heading
            heading_bookmark: Bookmark name (alternative to para_index)
            content_strategy: none, first_lines, ai_summary_first, full
            depth: How many sub-levels to include (1=direct, 0=all)
            file_path: Optional file path
        """
        try:
            doc = self._resolve_document(file_path)
            if not isinstance(doc, XTextDocument):
                return {"success": False,
                        "error": "Not a Writer document"}

            # Resolve bookmark to para_index if provided
            if heading_bookmark is not None and heading_para_index is None:
                if not hasattr(doc, 'getBookmarks'):
                    return {"success": False,
                            "error": "Document doesn't support bookmarks"}
                bm_sup = doc.getBookmarks()
                if not bm_sup.hasByName(heading_bookmark):
                    return {"success": False,
                            "error": f"Bookmark '{heading_bookmark}' not found"}
                bm = bm_sup.getByName(heading_bookmark)
                anchor = bm.getAnchor()
                para_ranges = self._get_paragraph_ranges(doc)
                heading_para_index = self._find_paragraph_for_range(
                    anchor, para_ranges, doc.getText())

            if heading_para_index is None:
                return {"success": False,
                        "error": "Provide heading_para_index or heading_bookmark"}

            tree = self._build_heading_tree(doc)
            bookmark_map = self._ensure_heading_bookmarks(doc)
            target = self._find_node_by_para_index(tree, heading_para_index)
            if target is None:
                return {"success": False,
                        "error": f"Heading at paragraph {heading_para_index} not found"}

            ai_summaries = (self._get_ai_summaries_map(doc)
                            if content_strategy in ("ai_summary_first",
                                                    "first_lines")
                            else {})

            # Build children list: body paragraphs + sub-headings
            children = []

            # Get body paragraphs between heading and first sub-heading
            text = doc.getText()
            enum = text.createEnumeration()
            idx = 0
            found_heading = False
            parent_level = target["level"]

            while enum.hasMoreElements():
                element = enum.nextElement()
                is_para = element.supportsService("com.sun.star.text.Paragraph")

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

                    if outline_level > 0 and outline_level <= parent_level:
                        break  # next sibling or parent heading

                    if outline_level > 0:
                        # This is a sub-heading; stop reading body,
                        # it will be added from tree children
                        break

                    # Body paragraph
                    para_text = element.getString()
                    preview = para_text[:100] + "..." if len(para_text) > 100 else para_text
                    if content_strategy == "full":
                        children.append({
                            "type": "body",
                            "para_index": idx,
                            "text": para_text
                        })
                    elif content_strategy != "none":
                        children.append({
                            "type": "body",
                            "para_index": idx,
                            "preview": preview
                        })
                    else:
                        children.append({
                            "type": "body",
                            "para_index": idx
                        })

                idx += 1

            # Add sub-headings from tree (with depth control + bookmarks)
            for child in target["children"]:
                node = self._serialize_tree_node(
                    child, doc, ai_summaries, content_strategy, depth,
                    bookmark_map=bookmark_map)
                children.append(node)

            return {
                "success": True,
                "parent": {
                    "level": target["level"],
                    "text": target["text"],
                    "para_index": target["para_index"],
                    "bookmark": bookmark_map.get(target["para_index"])
                },
                "content_strategy": content_strategy,
                "depth": depth,
                "children": children
            }
        except Exception as e:
            logger.error(f"Failed to get heading children: {e}")
            return {"success": False, "error": str(e)}

    # ----------------------------------------------------------------
    # Paragraph reading
    # ----------------------------------------------------------------

    def get_paragraph_count(self, file_path: str = None) -> Dict[str, Any]:
        """Count paragraphs in the document."""
        try:
            doc = self._resolve_document(file_path)
            text = doc.getText()
            enum = text.createEnumeration()
            count = 0
            while enum.hasMoreElements():
                para = enum.nextElement()
                if para.supportsService("com.sun.star.text.Paragraph"):
                    count += 1
            return {"success": True, "paragraph_count": count}
        except Exception as e:
            logger.error(f"Failed to get paragraph count: {e}")
            return {"success": False, "error": str(e)}

    def read_paragraphs(self, start_index: int, count: int = 10,
                        file_path: str = None) -> Dict[str, Any]:
        """Read a range of paragraphs by index."""
        try:
            doc = self._resolve_document(file_path)
            text = doc.getText()
            enum = text.createEnumeration()

            # Get existing bookmark map for heading paragraphs
            bookmark_map = self._get_mcp_bookmark_map(doc)

            paragraphs = []
            current_index = 0
            end_index = start_index + count

            while enum.hasMoreElements() and current_index < end_index:
                element = enum.nextElement()

                if current_index >= start_index:
                    if element.supportsService("com.sun.star.text.Paragraph"):
                        style_name = ""
                        outline_level = 0
                        try:
                            style_name = element.getPropertyValue(
                                "ParaStyleName")
                            outline_level = element.getPropertyValue(
                                "OutlineLevel")
                        except Exception:
                            pass
                        para_entry = {
                            "index": current_index,
                            "text": element.getString(),
                            "style_name": style_name,
                            "outline_level": outline_level,
                            "is_table": False
                        }
                        if current_index in bookmark_map:
                            para_entry["bookmark"] = bookmark_map[current_index]
                        paragraphs.append(para_entry)
                    elif element.supportsService(
                            "com.sun.star.text.TextTable"):
                        info = self._extract_table_info(element)
                        paragraphs.append({
                            "index": current_index,
                            "text": f"[Table: {info.get('name', '?')}, "
                                    f"{info.get('rows', '?')}x"
                                    f"{info.get('cols', '?')}]",
                            "style_name": "",
                            "outline_level": 0,
                            "is_table": True,
                            "table_info": info
                        })

                current_index += 1

            return {
                "success": True,
                "paragraphs": paragraphs,
                "start_index": start_index,
                "count_returned": len(paragraphs),
                "has_more": enum.hasMoreElements()
            }
        except Exception as e:
            logger.error(f"Failed to read paragraphs: {e}")
            return {"success": False, "error": str(e)}

    def _extract_table_info(self, table) -> Dict[str, Any]:
        """Extract basic info from a TextTable element."""
        try:
            name = table.getName() if hasattr(table, 'getName') else "unnamed"
            rows = table.getRows().getCount()
            cols = table.getColumns().getCount()
            return {"name": name, "rows": rows, "cols": cols}
        except Exception:
            return {"name": "unknown", "rows": 0, "cols": 0}

    # ----------------------------------------------------------------
    # Search & replace
    # ----------------------------------------------------------------

    def search_document(self, pattern: str, regex: bool = False,
                        case_sensitive: bool = False,
                        max_results: int = 20,
                        context_paragraphs: int = 1,
                        file_path: str = None) -> Dict[str, Any]:
        """Search document using XSearchable with paragraph context."""
        try:
            doc = self._resolve_document(file_path)

            search_desc = doc.createSearchDescriptor()
            search_desc.SearchString = pattern
            search_desc.SearchRegularExpression = regex
            search_desc.SearchCaseSensitive = case_sensitive

            found = doc.findAll(search_desc)
            if found is None or found.getCount() == 0:
                return {"success": True, "matches": [],
                        "total_found": 0}

            total_found = found.getCount()
            para_ranges = self._get_paragraph_ranges(doc)
            text_obj = doc.getText()

            # Get bookmark map for nearest-heading lookup
            bookmark_map = self._get_mcp_bookmark_map(doc)

            # Build paragraph texts for context
            para_texts = []
            text_enum = text_obj.createEnumeration()
            while text_enum.hasMoreElements():
                el = text_enum.nextElement()
                if el.supportsService("com.sun.star.text.Paragraph"):
                    para_texts.append(el.getString())
                else:
                    para_texts.append("[Table]")

            results = []
            for i in range(min(total_found, max_results)):
                match_range = found.getByIndex(i)
                match_text = match_range.getString()
                match_para_idx = self._find_paragraph_for_range(
                    match_range, para_ranges, text_obj)

                ctx_start = max(0, match_para_idx - context_paragraphs)
                ctx_end = min(len(para_texts),
                              match_para_idx + context_paragraphs + 1)
                context = [
                    {"index": j, "text": para_texts[j]}
                    for j in range(ctx_start, ctx_end)
                ]
                match_entry = {
                    "match_index": i,
                    "match_text": match_text,
                    "paragraph_index": match_para_idx,
                    "context": context
                }
                # Add nearest heading bookmark
                nearest = self._find_nearest_heading_bookmark(
                    match_para_idx, bookmark_map)
                if nearest:
                    match_entry["nearest_heading"] = nearest
                results.append(match_entry)

            return {
                "success": True,
                "matches": results,
                "total_found": total_found,
                "returned": len(results)
            }
        except Exception as e:
            logger.error(f"Failed to search document: {e}")
            return {"success": False, "error": str(e)}

    def replace_in_document(self, search: str, replace: str,
                            regex: bool = False,
                            case_sensitive: bool = False,
                            file_path: str = None) -> Dict[str, Any]:
        """Find and replace using XReplaceable."""
        try:
            doc = self._resolve_document(file_path)

            replace_desc = doc.createReplaceDescriptor()
            replace_desc.SearchString = search
            replace_desc.ReplaceString = replace
            replace_desc.SearchRegularExpression = regex
            replace_desc.SearchCaseSensitive = case_sensitive

            count = doc.replaceAll(replace_desc)

            if count > 0 and doc.hasLocation():
                doc.store()

            return {
                "success": True,
                "replacements_made": count,
                "search": search,
                "replace": replace
            }
        except Exception as e:
            logger.error(f"Failed to replace in document: {e}")
            return {"success": False, "error": str(e)}

    # ----------------------------------------------------------------
    # Editing
    # ----------------------------------------------------------------

    def insert_at_paragraph(self, paragraph_index: int, text: str,
                            position: str = "after",
                            file_path: str = None) -> Dict[str, Any]:
        """Insert text before or after a specific paragraph."""
        try:
            doc = self._resolve_document(file_path)
            doc_text = doc.getText()
            enum = doc_text.createEnumeration()

            current_index = 0
            target_para = None
            while enum.hasMoreElements():
                element = enum.nextElement()
                if current_index == paragraph_index:
                    target_para = element
                    break
                current_index += 1

            if target_para is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found "
                                 f"(max: {current_index})"}

            cursor = doc_text.createTextCursorByRange(target_para)

            if position == "before":
                cursor.gotoStartOfParagraph(False)
                doc_text.insertString(cursor, text, False)
                doc_text.insertControlCharacter(
                    cursor, PARAGRAPH_BREAK, False)
            elif position == "after":
                cursor.gotoEndOfParagraph(False)
                doc_text.insertControlCharacter(
                    cursor, PARAGRAPH_BREAK, False)
                doc_text.insertString(cursor, text, False)
            else:
                return {"success": False,
                        "error": f"Invalid position: {position}"}

            if doc.hasLocation():
                doc.store()

            return {
                "success": True,
                "message": f"Inserted text {position} paragraph "
                           f"{paragraph_index}",
                "text_length": len(text)
            }
        except Exception as e:
            logger.error(f"Failed to insert at paragraph: {e}")
            return {"success": False, "error": str(e)}

    # ----------------------------------------------------------------
    # AI annotations
    # ----------------------------------------------------------------

    def add_ai_summary(self, para_index: int, summary: str,
                       file_path: str = None) -> Dict[str, Any]:
        """Add an MCP-AI annotation at a heading paragraph."""
        try:
            doc = self._resolve_document(file_path)
            doc_text = doc.getText()

            # Remove existing MCP-AI annotation at this paragraph first
            self._remove_ai_annotation_at(doc, para_index)

            # Find the target paragraph
            enum = doc_text.createEnumeration()
            idx = 0
            target = None
            while enum.hasMoreElements():
                element = enum.nextElement()
                if idx == para_index:
                    target = element
                    break
                idx += 1

            if target is None:
                return {"success": False,
                        "error": f"Paragraph {para_index} not found"}

            # Create annotation
            annotation = doc.createInstance(
                "com.sun.star.text.textfield.Annotation")
            annotation.setPropertyValue("Author", "MCP-AI")
            annotation.setPropertyValue("Content", summary)

            cursor = doc_text.createTextCursorByRange(target.getStart())
            doc_text.insertTextContent(cursor, annotation, False)

            if doc.hasLocation():
                doc.store()

            return {
                "success": True,
                "message": f"Added AI summary at paragraph {para_index}",
                "para_index": para_index,
                "summary_length": len(summary)
            }
        except Exception as e:
            logger.error(f"Failed to add AI summary: {e}")
            return {"success": False, "error": str(e)}

    def get_ai_summaries(self, file_path: str = None) -> Dict[str, Any]:
        """List all MCP-AI annotations in the document."""
        try:
            doc = self._resolve_document(file_path)
            summaries_map = self._get_ai_summaries_map(doc)

            summaries = [
                {"para_index": idx, "summary": text}
                for idx, text in sorted(summaries_map.items())
            ]
            return {
                "success": True,
                "summaries": summaries,
                "count": len(summaries)
            }
        except Exception as e:
            logger.error(f"Failed to get AI summaries: {e}")
            return {"success": False, "error": str(e)}

    def remove_ai_summary(self, para_index: int,
                          file_path: str = None) -> Dict[str, Any]:
        """Remove an MCP-AI annotation from a paragraph."""
        try:
            doc = self._resolve_document(file_path)
            removed = self._remove_ai_annotation_at(doc, para_index)

            if removed and doc.hasLocation():
                doc.store()

            return {
                "success": True,
                "removed": removed,
                "para_index": para_index
            }
        except Exception as e:
            logger.error(f"Failed to remove AI summary: {e}")
            return {"success": False, "error": str(e)}

    def _remove_ai_annotation_at(self, doc, para_index: int) -> bool:
        """Remove MCP-AI annotation at a specific paragraph. Returns True if removed."""
        try:
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            para_ranges = self._get_paragraph_ranges(doc)
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
                idx = self._find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                if idx == para_index:
                    text_obj.removeTextContent(field)
                    return True
        except Exception as e:
            logger.error(f"Failed to remove AI annotation: {e}")
        return False

    # ----------------------------------------------------------------
    # Structural navigation
    # ----------------------------------------------------------------

    def list_sections(self, file_path: str = None) -> Dict[str, Any]:
        """List all named text sections."""
        try:
            doc = self._resolve_document(file_path)
            if not hasattr(doc, 'getTextSections'):
                return {"success": True, "sections": [], "count": 0}

            supplier = doc.getTextSections()
            names = supplier.getElementNames()
            sections = []
            for name in names:
                section = supplier.getByName(name)
                sections.append({
                    "name": name,
                    "is_visible": getattr(section, 'IsVisible', True),
                    "is_protected": getattr(section, 'IsProtected', False)
                })
            return {"success": True, "sections": sections,
                    "count": len(sections)}
        except Exception as e:
            logger.error(f"Failed to list sections: {e}")
            return {"success": False, "error": str(e)}

    def read_section(self, section_name: str,
                     file_path: str = None) -> Dict[str, Any]:
        """Read the content of a named text section."""
        try:
            doc = self._resolve_document(file_path)
            if not hasattr(doc, 'getTextSections'):
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
            logger.error(f"Failed to read section: {e}")
            return {"success": False, "error": str(e)}

    def list_bookmarks(self, file_path: str = None) -> Dict[str, Any]:
        """List all bookmarks in the document."""
        try:
            doc = self._resolve_document(file_path)
            if not hasattr(doc, 'getBookmarks'):
                return {"success": True, "bookmarks": [], "count": 0}

            bookmarks = doc.getBookmarks()
            names = bookmarks.getElementNames()
            result = []
            for name in names:
                bm = bookmarks.getByName(name)
                anchor_text = bm.getAnchor().getString()
                result.append({
                    "name": name,
                    "preview": anchor_text[:100] if anchor_text else ""
                })
            return {"success": True, "bookmarks": result,
                    "count": len(result)}
        except Exception as e:
            logger.error(f"Failed to list bookmarks: {e}")
            return {"success": False, "error": str(e)}

    def get_page_count(self, file_path: str = None) -> Dict[str, Any]:
        """Get the page count via XPageCursor."""
        try:
            doc = self._resolve_document(file_path)
            try:
                controller = doc.getCurrentController()
                if controller:
                    vc = controller.getViewCursor()
                    vc.jumpToLastPage()
                    page_count = vc.getPage()
                    vc.jumpToFirstPage()
                    return {"success": True, "page_count": page_count}
            except Exception:
                pass
            return {"success": False,
                    "error": "Could not determine page count "
                             "(requires visible document)"}
        except Exception as e:
            logger.error(f"Failed to get page count: {e}")
            return {"success": False, "error": str(e)}

    # ----------------------------------------------------------------
    # Original methods
    # ----------------------------------------------------------------

    def get_document_info(self, doc: Any = None) -> Dict[str, Any]:
        """Get information about a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"error": "No document available"}
            
            info = {
                "title": getattr(doc, 'Title', 'Unknown') if hasattr(doc, 'Title') else "Unknown",
                "url": doc.getURL() if hasattr(doc, 'getURL') else "",
                "modified": doc.isModified() if hasattr(doc, 'isModified') else False,
                "type": self._get_document_type(doc),
                "has_selection": self._has_selection(doc)
            }
            
            # Add document-specific information
            if isinstance(doc, XTextDocument):
                text = doc.getText()
                info["word_count"] = len(text.getString().split())
                info["character_count"] = len(text.getString())
            elif isinstance(doc, XSpreadsheetDocument):
                sheets = doc.getSheets()
                info["sheet_count"] = sheets.getCount()
                info["sheet_names"] = [sheets.getByIndex(i).getName() 
                                     for i in range(sheets.getCount())]
            
            return info
            
        except Exception as e:
            logger.error(f"Failed to get document info: {e}")
            return {"error": str(e)}
    
    def insert_text(self, text: str, position: Optional[int] = None, doc: Any = None) -> Dict[str, Any]:
        """
        Insert text into a document
        
        Args:
            text: Text to insert
            position: Position to insert at (None for current cursor position)
            doc: Document to insert into (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No active document"}
            
            # Handle Writer documents
            if isinstance(doc, XTextDocument):
                text_obj = doc.getText()
                
                if position is None:
                    # Insert at current cursor position
                    cursor = doc.getCurrentController().getViewCursor()
                else:
                    # Insert at specific position
                    cursor = text_obj.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.goRight(position, False)
                
                text_obj.insertString(cursor, text, False)
                logger.info(f"Inserted {len(text)} characters into Writer document")
                return {"success": True, "message": f"Inserted {len(text)} characters"}
            
            # Handle other document types
            else:
                return {"success": False, "error": f"Text insertion not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to insert text: {e}")
            return {"success": False, "error": str(e)}
    
    def format_text(self, formatting: Dict[str, Any], doc: Any = None) -> Dict[str, Any]:
        """
        Apply formatting to selected text
        
        Args:
            formatting: Dictionary of formatting options
            doc: Document to format (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc or not isinstance(doc, XTextDocument):
                return {"success": False, "error": "No Writer document available"}
            
            # Get current selection
            selection = doc.getCurrentController().getSelection()
            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}
            
            # Apply formatting to selection
            text_range = selection.getByIndex(0)
            
            # Apply various formatting options
            if "bold" in formatting:
                text_range.CharWeight = 150.0 if formatting["bold"] else 100.0
            
            if "italic" in formatting:
                text_range.CharPosture = 2 if formatting["italic"] else 0
            
            if "underline" in formatting:
                text_range.CharUnderline = 1 if formatting["underline"] else 0
            
            if "font_size" in formatting:
                text_range.CharHeight = formatting["font_size"]
            
            if "font_name" in formatting:
                text_range.CharFontName = formatting["font_name"]
            
            logger.info("Applied formatting to selected text")
            return {"success": True, "message": "Formatting applied successfully"}
            
        except Exception as e:
            logger.error(f"Failed to format text: {e}")
            return {"success": False, "error": str(e)}
    
    def save_document(self, doc: Any = None, file_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Save a document
        
        Args:
            doc: Document to save (None for active document)
            file_path: Path to save to (None to save to current location)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to save"}
            
            if file_path:
                # Save as new file
                url = uno.systemPathToFileUrl(file_path)
                doc.storeAsURL(url, ())
                logger.info(f"Saved document to {file_path}")
                return {"success": True, "message": f"Document saved to {file_path}"}
            else:
                # Save to current location
                if doc.hasLocation():
                    doc.store()
                    logger.info("Saved document to current location")
                    return {"success": True, "message": "Document saved"}
                else:
                    return {"success": False, "error": "Document has no location, specify file_path"}
                    
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return {"success": False, "error": str(e)}
    
    def export_document(self, export_format: str, file_path: str, doc: Any = None) -> Dict[str, Any]:
        """
        Export document to different format
        
        Args:
            export_format: Target format ('pdf', 'docx', 'odt', 'txt', etc.)
            file_path: Path to export to
            doc: Document to export (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to export"}
            
            # Filter map for different formats
            filter_map = {
                'pdf': 'writer_pdf_Export',
                'docx': 'MS Word 2007 XML',
                'doc': 'MS Word 97',
                'odt': 'writer8',
                'txt': 'Text',
                'rtf': 'Rich Text Format',
                'html': 'HTML (StarWriter)'
            }
            
            filter_name = filter_map.get(export_format.lower())
            if not filter_name:
                return {"success": False, "error": f"Unsupported export format: {export_format}"}
            
            # Prepare export properties
            properties = (
                PropertyValue("FilterName", 0, filter_name, 0),
                PropertyValue("Overwrite", 0, True, 0),
            )
            
            # Export document
            url = uno.systemPathToFileUrl(file_path)
            doc.storeToURL(url, properties)
            
            logger.info(f"Exported document to {file_path} as {export_format}")
            return {"success": True, "message": f"Document exported to {file_path}"}
            
        except Exception as e:
            logger.error(f"Failed to export document: {e}")
            return {"success": False, "error": str(e)}
    
    def get_text_content(self, doc: Any = None) -> Dict[str, Any]:
        """Get text content from a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document available"}
            
            if isinstance(doc, XTextDocument):
                text = doc.getText().getString()
                return {"success": True, "content": text, "length": len(text)}
            else:
                return {"success": False, "error": f"Text extraction not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to get text content: {e}")
            return {"success": False, "error": str(e)}
    
    def _get_document_type(self, doc: Any) -> str:
        """Determine document type"""
        if isinstance(doc, XTextDocument):
            return "writer"
        elif isinstance(doc, XSpreadsheetDocument):
            return "calc"
        elif XPresentationDocument is not None and isinstance(doc, XPresentationDocument):
            return "impress"
        else:
            return "unknown"
    
    def _has_selection(self, doc: Any) -> bool:
        """Check if document has selected content"""
        try:
            if hasattr(doc, 'getCurrentController'):
                controller = doc.getCurrentController()
                if hasattr(controller, 'getSelection'):
                    selection = controller.getSelection()
                    return selection.getCount() > 0
        except:
            pass
        return False
