"""
WriterService — facade for Writer document operations.

Delegates to sub-services for specific domains while providing
shared helpers used across sub-services.
"""

import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class WriterService:
    """Facade for all Writer document operations.

    Sub-services:
        tree        — heading navigation, bookmarks, content strategies
        paragraphs  — paragraph CRUD, reading, editing
        search      — search & replace
        structural  — sections, pages, indexes, locator resolution
        proximity   — local heading navigation, surroundings discovery
    """

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

        from .tree import TreeService
        from .paragraphs import ParagraphService
        from .search import SearchService
        from .structural import StructuralService
        from .proximity import ProximityService
        from .index import IndexService

        self.tree = TreeService(self)
        self.paragraphs = ParagraphService(self)
        self.search = SearchService(self)
        self.structural = StructuralService(self)
        self.proximity = ProximityService(self)
        self.index = IndexService(self)

        logger.info("WriterService ready (6 sub-services)")

    # ==================================================================
    # Shared helpers (used across sub-services)
    # ==================================================================

    def get_paragraph_ranges(self, doc) -> List[Any]:
        """Get list of paragraph elements for range comparison."""
        text = doc.getText()
        enum = text.createEnumeration()
        ranges = []
        while enum.hasMoreElements():
            ranges.append(enum.nextElement())
            self._base.yield_to_gui()
        return ranges

    def find_paragraph_for_range(self, match_range, para_ranges: List,
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

    def find_paragraph_element(self, doc, para_index: int):
        """Find a paragraph element by index. Returns (element, max_index)."""
        doc_text = doc.getText()
        enum = doc_text.createEnumeration()
        idx = 0
        while enum.hasMoreElements():
            element = enum.nextElement()
            if idx == para_index:
                return element, idx
            idx += 1
        return None, idx

    def extract_table_info(self, table) -> Dict[str, Any]:
        """Extract basic info from a TextTable element."""
        try:
            name = table.getName() if hasattr(table, "getName") else "unnamed"
            rows = table.getRows().getCount()
            cols = table.getColumns().getCount()
            return {"name": name, "rows": rows, "cols": cols}
        except Exception:
            return {"name": "unknown", "rows": 0, "cols": 0}

    def invalidate_caches(self, doc=None):
        """Invalidate all per-document caches after an edit."""
        if self._registry.batch_mode:
            return  # deferred — batch will call once at end
        self.tree.invalidate_cache(doc)
        self.proximity.invalidate_cache(doc)
        self.index.invalidate_cache(doc)
        self._base.invalidate_page_cache()

    # ==================================================================
    # Delegated public API (tools call these — no change needed)
    # ==================================================================

    # -- Tree / Navigation --
    def get_document_tree(self, *a, **kw):
        return self.tree.get_document_tree(*a, **kw)

    def get_heading_children(self, *a, **kw):
        return self.tree.get_heading_children(*a, **kw)

    # -- Paragraph reading --
    def get_paragraph_count(self, *a, **kw):
        return self.paragraphs.get_paragraph_count(*a, **kw)

    def read_paragraphs(self, *a, **kw):
        return self.paragraphs.read_paragraphs(*a, **kw)

    # -- Paragraph editing --
    def insert_at_paragraph(self, *a, **kw):
        return self.paragraphs.insert_at_paragraph(*a, **kw)

    def insert_paragraphs_batch(self, *a, **kw):
        return self.paragraphs.insert_paragraphs_batch(*a, **kw)

    def delete_paragraph(self, *a, **kw):
        return self.paragraphs.delete_paragraph(*a, **kw)

    def set_paragraph_text(self, *a, **kw):
        return self.paragraphs.set_paragraph_text(*a, **kw)

    def set_paragraph_style(self, *a, **kw):
        return self.paragraphs.set_paragraph_style(*a, **kw)

    def duplicate_paragraph(self, *a, **kw):
        return self.paragraphs.duplicate_paragraph(*a, **kw)

    def clone_heading_block(self, *a, **kw):
        return self.paragraphs.clone_heading_block(*a, **kw)

    # -- Search --
    def search_document(self, *a, **kw):
        return self.search.search_document(*a, **kw)

    def replace_in_document(self, *a, **kw):
        return self.search.replace_in_document(*a, **kw)

    # -- Full-text index --
    def search_boolean(self, *a, **kw):
        return self.index.search_boolean(*a, **kw)

    def get_index_stats(self, *a, **kw):
        return self.index.get_index_stats(*a, **kw)

    # -- AI annotations --
    def add_ai_summary(self, *a, **kw):
        return self.tree.add_ai_summary(*a, **kw)

    def get_ai_summaries(self, *a, **kw):
        return self.tree.get_ai_summaries(*a, **kw)

    def remove_ai_summary(self, *a, **kw):
        return self.tree.remove_ai_summary(*a, **kw)

    # -- Structural --
    def resolve_writer_locator(self, *a, **kw):
        return self.structural.resolve_writer_locator(*a, **kw)

    def resolve_bookmark(self, *a, **kw):
        return self.structural.resolve_bookmark(*a, **kw)

    def list_sections(self, *a, **kw):
        return self.structural.list_sections(*a, **kw)

    def read_section(self, *a, **kw):
        return self.structural.read_section(*a, **kw)

    def list_bookmarks(self, *a, **kw):
        return self.structural.list_bookmarks(*a, **kw)

    def get_page_count(self, *a, **kw):
        return self.structural.get_page_count(*a, **kw)

    def goto_page(self, *a, **kw):
        return self.structural.goto_page(*a, **kw)

    def get_page_objects(self, *a, **kw):
        return self.structural.get_page_objects(*a, **kw)

    def refresh_indexes(self, *a, **kw):
        return self.structural.refresh_indexes(*a, **kw)

    def update_fields(self, *a, **kw):
        return self.structural.update_fields(*a, **kw)

    # -- Proximity --
    def navigate_heading(self, *a, **kw):
        return self.proximity.navigate_heading(*a, **kw)

    def get_surroundings(self, *a, **kw):
        return self.proximity.get_surroundings(*a, **kw)

    # -- Document ops (delegated to BaseService) --
    def get_document_properties(self, *a, **kw):
        return self._base.get_document_properties(*a, **kw)

    def set_document_properties(self, *a, **kw):
        return self._base.set_document_properties(*a, **kw)

    def save_document_as(self, *a, **kw):
        return self._base.save_document_as(*a, **kw)

    def save_document(self, *a, **kw):
        return self._base.save_document(*a, **kw)

    def set_document_protection(self, *a, **kw):
        return self._base.set_document_protection(*a, **kw)

    def get_recent_documents(self, *a, **kw):
        return self._base.get_recent_documents(*a, **kw)
