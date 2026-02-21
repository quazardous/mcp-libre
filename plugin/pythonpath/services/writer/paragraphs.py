"""
ParagraphService — paragraph CRUD operations.

Reading, inserting, deleting, modifying paragraph text and style.
"""

import logging
from typing import Any, Dict, Optional

from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

logger = logging.getLogger(__name__)


class ParagraphService:
    """Paragraph-level operations on Writer documents."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base

    # ==================================================================
    # Helpers
    # ==================================================================

    def _is_inside_index(self, paragraph) -> Optional[str]:
        """Check if paragraph is inside a document index (ToC, etc.)."""
        try:
            section = paragraph.getPropertyValue("TextSection")
            if section is None:
                return None
            si = section.supportsService
            if (si("com.sun.star.text.BaseIndex")
                    or si("com.sun.star.text.DocumentIndex")
                    or si("com.sun.star.text.ContentIndex")):
                return section.Name
            if hasattr(section, "ParentSection"):
                parent = section.ParentSection
                if parent is not None:
                    pi = parent.supportsService
                    if (pi("com.sun.star.text.BaseIndex")
                            or pi("com.sun.star.text.DocumentIndex")
                            or pi("com.sun.star.text.ContentIndex")):
                        return parent.Name
        except Exception:
            pass
        try:
            section = paragraph.getPropertyValue("TextSection")
            if section and section.IsProtected:
                return section.Name
        except Exception:
            pass
        return None

    def _resolve_style_name(self, doc, style_name: str) -> str:
        """Resolve a style name case-insensitively."""
        try:
            families = doc.getStyleFamilies()
            para_styles = families.getByName("ParagraphStyles")
            if para_styles.hasByName(style_name):
                return style_name
            lower = style_name.lower()
            for name in para_styles.getElementNames():
                if name.lower() == lower:
                    return name
        except Exception:
            pass
        return style_name

    # ==================================================================
    # Reading
    # ==================================================================

    def get_paragraph_count(self, file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            text = doc.getText()
            enum = text.createEnumeration()
            count = 0
            while enum.hasMoreElements():
                para = enum.nextElement()
                if para.supportsService("com.sun.star.text.Paragraph"):
                    count += 1
            return {"success": True, "paragraph_count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_paragraphs(self, start_index: int = None, count: int = 10,
                        locator: str = None,
                        file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and start_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                start_index = resolved.get("para_index", 0)
            if start_index is None:
                start_index = 0

            text = doc.getText()
            enum = text.createEnumeration()
            bookmark_map = self._writer.tree.get_mcp_bookmark_map(doc)
            paragraphs = []
            current_index = 0
            end_index = start_index + count

            while enum.hasMoreElements() and current_index < end_index:
                element = enum.nextElement()
                if current_index >= start_index:
                    if element.supportsService(
                            "com.sun.star.text.Paragraph"):
                        style_name = ""
                        outline_level = 0
                        try:
                            style_name = element.getPropertyValue(
                                "ParaStyleName")
                            outline_level = element.getPropertyValue(
                                "OutlineLevel")
                        except Exception:
                            pass
                        entry = {
                            "index": current_index,
                            "text": element.getString(),
                            "style_name": style_name,
                            "outline_level": outline_level,
                            "is_table": False,
                        }
                        if current_index in bookmark_map:
                            entry["bookmark"] = bookmark_map[current_index]
                        paragraphs.append(entry)
                    elif element.supportsService(
                            "com.sun.star.text.TextTable"):
                        info = self._writer.extract_table_info(element)
                        paragraphs.append({
                            "index": current_index,
                            "text": f"[Table: {info.get('name', '?')}, "
                                    f"{info.get('rows', '?')}x"
                                    f"{info.get('cols', '?')}]",
                            "style_name": "",
                            "outline_level": 0,
                            "is_table": True,
                            "table_info": info,
                        })
                current_index += 1
                self._base.yield_to_gui()

            return {
                "success": True,
                "paragraphs": paragraphs,
                "start_index": start_index,
                "count_returned": len(paragraphs),
                "has_more": enum.hasMoreElements(),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Editing
    # ==================================================================

    def insert_at_paragraph(self, paragraph_index: int = None,
                            text: str = "", position: str = "after",
                            locator: str = None, style: str = None,
                            file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if style:
                style = self._resolve_style_name(doc, style)
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            target, _ = self._writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            idx_name = self._is_inside_index(target)
            if idx_name:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} is inside "
                                 f"index '{idx_name}'. Use refresh_indexes()."}

            doc_text = doc.getText()
            cursor = doc_text.createTextCursorByRange(target)

            if position == "before":
                cursor.gotoStartOfParagraph(False)
                doc_text.insertString(cursor, text, False)
                doc_text.insertControlCharacter(
                    cursor, PARAGRAPH_BREAK, False)
                if style:
                    cursor.gotoPreviousParagraph(False)
                    cursor.gotoStartOfParagraph(False)
                    cursor.gotoEndOfParagraph(True)
                    cursor.setPropertyValue("ParaStyleName", style)
                    cursor.gotoNextParagraph(False)
            elif position == "after":
                cursor.gotoEndOfParagraph(False)
                doc_text.insertControlCharacter(
                    cursor, PARAGRAPH_BREAK, False)
                doc_text.insertString(cursor, text, False)
                if style:
                    cursor.gotoStartOfParagraph(False)
                    cursor.gotoEndOfParagraph(True)
                    cursor.setPropertyValue("ParaStyleName", style)
                    cursor.gotoEndOfParagraph(False)
            else:
                return {"success": False,
                        "error": f"Invalid position: {position}"}

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            result = {"success": True,
                      "message": f"Inserted text {position} paragraph "
                                 f"{paragraph_index}",
                      "text_length": len(text)}
            if style:
                result["style"] = style
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_paragraphs_batch(self, paragraphs: list,
                                paragraph_index: int = None,
                                position: str = "after",
                                locator: str = None,
                                file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            for item in paragraphs or []:
                if item.get("style"):
                    item["style"] = self._resolve_style_name(
                        doc, item["style"])
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}
            if not paragraphs:
                return {"success": False, "error": "Empty paragraphs list"}

            target, _ = self._writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            idx_name = self._is_inside_index(target)
            if idx_name:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} is inside "
                                 f"index '{idx_name}'. Use refresh_indexes()."}

            doc_text = doc.getText()
            cursor = doc_text.createTextCursorByRange(target)

            if position == "before":
                cursor.gotoStartOfParagraph(False)
                for item in paragraphs:
                    txt = item.get("text", "")
                    sty = item.get("style")
                    doc_text.insertString(cursor, txt, False)
                    doc_text.insertControlCharacter(
                        cursor, PARAGRAPH_BREAK, False)
                    if sty:
                        cursor.gotoPreviousParagraph(False)
                        cursor.gotoStartOfParagraph(False)
                        cursor.gotoEndOfParagraph(True)
                        cursor.setPropertyValue("ParaStyleName", sty)
                        cursor.gotoNextParagraph(False)
            elif position == "after":
                cursor.gotoEndOfParagraph(False)
                for item in paragraphs:
                    txt = item.get("text", "")
                    sty = item.get("style")
                    doc_text.insertControlCharacter(
                        cursor, PARAGRAPH_BREAK, False)
                    doc_text.insertString(cursor, txt, False)
                    if sty:
                        cursor.gotoStartOfParagraph(False)
                        cursor.gotoEndOfParagraph(True)
                        cursor.setPropertyValue("ParaStyleName", sty)
                        cursor.gotoEndOfParagraph(False)
            else:
                return {"success": False,
                        "error": f"Invalid position: {position}"}

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "message": f"Inserted {len(paragraphs)} paragraphs "
                               f"{position} paragraph {paragraph_index}",
                    "count": len(paragraphs)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_paragraph(self, paragraph_index: int = None,
                         locator: str = None,
                         file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            doc_text = doc.getText()
            enum = doc_text.createEnumeration()
            idx = 0
            target = None
            while enum.hasMoreElements():
                element = enum.nextElement()
                if idx == paragraph_index:
                    target = element
                    break
                idx += 1

            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            idx_name = self._is_inside_index(target)
            if idx_name:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} is inside "
                                 f"index '{idx_name}'. Use refresh_indexes()."}

            cursor = doc_text.createTextCursorByRange(target)
            cursor.gotoStartOfParagraph(False)
            cursor.gotoEndOfParagraph(True)
            if enum.hasMoreElements():
                cursor.goRight(1, True)
            cursor.setString("")

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "message": f"Deleted paragraph {paragraph_index}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_text(self, text: str = "",
                           paragraph_index: int = None,
                           locator: str = None,
                           file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            target, _ = self._writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            idx_name = self._is_inside_index(target)
            if idx_name:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} is inside "
                                 f"index '{idx_name}'. Use refresh_indexes()."}

            old_text = target.getString()
            target.setString(text)

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "paragraph_index": paragraph_index,
                    "old_length": len(old_text), "new_length": len(text)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_style(self, style_name: str,
                            paragraph_index: int = None,
                            locator: str = None,
                            file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            target, _ = self._writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            idx_name = self._is_inside_index(target)
            if idx_name:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} is inside "
                                 f"index '{idx_name}'. Use refresh_indexes()."}

            old_style = target.getPropertyValue("ParaStyleName")
            target.setPropertyValue("ParaStyleName", style_name)

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "paragraph_index": paragraph_index,
                    "old_style": old_style, "new_style": style_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def duplicate_paragraph(self, paragraph_index: int = None,
                            locator: str = None, count: int = 1,
                            file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            doc_text = doc.getText()
            enum = doc_text.createEnumeration()
            elements = []
            idx = 0
            while enum.hasMoreElements():
                el = enum.nextElement()
                if paragraph_index <= idx < paragraph_index + count:
                    elements.append(el)
                if idx >= paragraph_index + count - 1:
                    break
                idx += 1

            if not elements:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            last = elements[-1]
            cursor = doc_text.createTextCursorByRange(last)
            cursor.gotoEndOfParagraph(False)

            for el in elements:
                txt = el.getString()
                sty = el.getPropertyValue("ParaStyleName")
                doc_text.insertControlCharacter(
                    cursor, PARAGRAPH_BREAK, False)
                doc_text.insertString(cursor, txt, False)
                cursor.gotoStartOfParagraph(False)
                cursor.gotoEndOfParagraph(True)
                cursor.setPropertyValue("ParaStyleName", sty)
                cursor.gotoEndOfParagraph(False)

            self._writer.invalidate_caches(doc)
            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "message": f"Duplicated {count} paragraph(s) "
                               f"at {paragraph_index}",
                    "duplicated_count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}
