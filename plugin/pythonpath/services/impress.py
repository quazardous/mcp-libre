"""
ImpressService — Impress presentation operations.
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class ImpressService:
    """Presentation slide operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    def list_slides(self, file_path: str = None) -> Dict[str, Any]:
        """List slides: count, titles, layout names."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_impress(doc):
                return {"success": False,
                        "error": "Not an Impress document"}

            pages = doc.getDrawPages()
            count = pages.getCount()
            slides = []
            for i in range(count):
                page = pages.getByIndex(i)
                name = (page.Name if hasattr(page, 'Name')
                        else f"Slide {i + 1}")
                layout = ""
                try:
                    layout = str(page.Layout)
                except Exception:
                    pass
                title = ""
                for s in range(page.getCount()):
                    shape = page.getByIndex(s)
                    if hasattr(shape, 'getString'):
                        txt = shape.getString().strip()
                        if txt:
                            title = txt[:100]
                            break
                slides.append({
                    "index": i,
                    "name": name,
                    "layout": layout,
                    "title": title,
                })
            return {"success": True, "slides": slides, "count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_slide_text(self, slide_index: int,
                        file_path: str = None) -> Dict[str, Any]:
        """Get all text from a slide + notes page."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_impress(doc):
                return {"success": False,
                        "error": "Not an Impress document"}

            pages = doc.getDrawPages()
            if slide_index < 0 or slide_index >= pages.getCount():
                return {"success": False,
                        "error": f"Slide index {slide_index} out of range "
                                 f"(0..{pages.getCount() - 1})"}

            page = pages.getByIndex(slide_index)
            texts = []
            for s in range(page.getCount()):
                shape = page.getByIndex(s)
                if hasattr(shape, 'getString'):
                    txt = shape.getString()
                    if txt.strip():
                        texts.append(txt)

            notes_text = ""
            try:
                notes_page = page.getNotesPage()
                for s in range(notes_page.getCount()):
                    shape = notes_page.getByIndex(s)
                    if hasattr(shape, 'getString'):
                        txt = shape.getString().strip()
                        if txt:
                            notes_text += txt + "\n"
            except Exception:
                pass

            return {
                "success": True,
                "slide_index": slide_index,
                "name": page.Name if hasattr(page, 'Name') else "",
                "texts": texts,
                "notes": notes_text.strip(),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_presentation_info(self,
                              file_path: str = None) -> Dict[str, Any]:
        """Slide count, dimensions, master page names."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_impress(doc):
                return {"success": False,
                        "error": "Not an Impress document"}

            pages = doc.getDrawPages()
            slide_count = pages.getCount()

            width = height = 0
            if slide_count > 0:
                first = pages.getByIndex(0)
                width = first.Width
                height = first.Height

            masters = doc.getMasterPages()
            master_names = []
            for i in range(masters.getCount()):
                mp = masters.getByIndex(i)
                master_names.append(
                    mp.Name if hasattr(mp, 'Name')
                    else f"Master {i + 1}")

            return {
                "success": True,
                "slide_count": slide_count,
                "width": width,
                "height": height,
                "master_pages": master_names,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
