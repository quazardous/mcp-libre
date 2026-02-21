"""
SearchService — search and replace in Writer documents.
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class SearchService:
    """Search and replace operations on Writer documents."""

    def __init__(self, writer):
        self._writer = writer
        self._base = writer._base

    def search_document(self, pattern: str, regex: bool = False,
                        case_sensitive: bool = False,
                        max_results: int = 20,
                        context_paragraphs: int = 1,
                        file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            search_desc = doc.createSearchDescriptor()
            search_desc.SearchString = pattern
            search_desc.SearchRegularExpression = regex
            search_desc.SearchCaseSensitive = case_sensitive

            found = doc.findAll(search_desc)
            if found is None or found.getCount() == 0:
                return {"success": True, "matches": [],
                        "total_found": 0}

            total_found = found.getCount()
            para_ranges = self._writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()
            bookmark_map = self._writer.tree.get_mcp_bookmark_map(doc)

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
                match_para_idx = self._writer.find_paragraph_for_range(
                    match_range, para_ranges, text_obj)
                ctx_start = max(0, match_para_idx - context_paragraphs)
                ctx_end = min(len(para_texts),
                              match_para_idx + context_paragraphs + 1)
                context = [{"index": j, "text": para_texts[j]}
                           for j in range(ctx_start, ctx_end)]
                entry = {"match_index": i, "match_text": match_text,
                         "paragraph_index": match_para_idx,
                         "context": context}
                nearest = self._writer.tree.find_nearest_heading_bookmark(
                    match_para_idx, bookmark_map)
                if nearest:
                    entry["nearest_heading"] = nearest
                results.append(entry)

            return {"success": True, "matches": results,
                    "total_found": total_found,
                    "returned": len(results)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def replace_in_document(self, search: str, replace: str,
                            regex: bool = False,
                            case_sensitive: bool = False,
                            file_path: str = None) -> Dict[str, Any]:
        try:
            doc = self._base.resolve_document(file_path)
            replace_desc = doc.createReplaceDescriptor()
            replace_desc.SearchString = search
            replace_desc.ReplaceString = replace
            replace_desc.SearchRegularExpression = regex
            replace_desc.SearchCaseSensitive = case_sensitive
            count = doc.replaceAll(replace_desc)

            if count > 0:
                self._writer.invalidate_caches(doc)
                if doc.hasLocation():
                    self._base.store_doc(doc)

            return {"success": True, "replacements_made": count,
                    "search": search, "replace": replace}
        except Exception as e:
            return {"success": False, "error": str(e)}
