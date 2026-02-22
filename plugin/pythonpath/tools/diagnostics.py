"""Diagnostic tools — document health checks."""

from .base import McpTool


class DocumentHealthCheck(McpTool):
    name = "document_health_check"
    description = (
        "Run diagnostics on a document: detect empty headings, "
        "broken bookmarks, orphan images, inconsistent heading levels, "
        "and large unstructured sections. Returns issues sorted by severity."
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
        doc = self.services.writer._base.resolve_document(file_path)
        issues = []

        # 1. Empty headings
        doc_text = doc.getText()
        enum = doc_text.createEnumeration()
        idx = 0
        prev_level = 0
        headings = []
        while enum.hasMoreElements():
            el = enum.nextElement()
            if not el.supportsService("com.sun.star.text.Paragraph"):
                idx += 1
                continue
            try:
                level = el.getPropertyValue("OutlineLevel")
            except Exception:
                level = 0
            if level > 0:
                text = el.getString().strip()
                headings.append({
                    "para_index": idx, "level": level, "text": text})
                if not text:
                    issues.append({
                        "severity": "warning",
                        "type": "empty_heading",
                        "para_index": idx,
                        "level": level,
                        "message": f"Empty heading (level {level}) "
                                   f"at paragraph {idx}",
                    })
                # Check heading level jumps (e.g. H1 -> H3, skipping H2)
                if prev_level > 0 and level > prev_level + 1:
                    issues.append({
                        "severity": "info",
                        "type": "heading_level_skip",
                        "para_index": idx,
                        "message": f"Heading level jumps from {prev_level} "
                                   f"to {level} at paragraph {idx}: "
                                   f"'{text[:50]}'",
                    })
                prev_level = level
            idx += 1

        # 2. Broken bookmarks (point to nonexistent positions)
        if hasattr(doc, 'getBookmarks'):
            bookmarks = doc.getBookmarks()
            total_paras = idx
            for i in range(bookmarks.getCount()):
                try:
                    bm = bookmarks.getByIndex(i)
                    name = bm.getName()
                    anchor = bm.getAnchor()
                    text = anchor.getString()[:50] if anchor else ""
                    # Check if bookmark text is empty (might be orphan)
                    if name.startswith("_mcp_") and not text:
                        issues.append({
                            "severity": "info",
                            "type": "empty_bookmark",
                            "bookmark": name,
                            "message": f"Bookmark '{name}' has no text "
                                       f"(heading may have been deleted)",
                        })
                except Exception:
                    pass

        # 3. Orphan images (no anchor or broken reference)
        if hasattr(doc, 'getGraphicObjects'):
            graphics = doc.getGraphicObjects()
            for name in graphics.getElementNames():
                try:
                    img = graphics.getByName(name)
                    url = img.getPropertyValue("GraphicURL")
                    if not url:
                        issues.append({
                            "severity": "warning",
                            "type": "broken_image",
                            "image": name,
                            "message": f"Image '{name}' has no graphic URL",
                        })
                except Exception:
                    pass

        # 4. Large unstructured sections (>50 body paragraphs without heading)
        if len(headings) >= 2:
            for i in range(len(headings) - 1):
                gap = headings[i + 1]["para_index"] - headings[i]["para_index"]
                if gap > 50:
                    issues.append({
                        "severity": "info",
                        "type": "large_unstructured_block",
                        "para_index": headings[i]["para_index"],
                        "heading": headings[i]["text"][:50],
                        "paragraph_count": gap,
                        "message": f"{gap} paragraphs under "
                                   f"'{headings[i]['text'][:40]}' — "
                                   f"consider splitting with sub-headings",
                    })

        return {
            "success": True,
            "total_paragraphs": idx,
            "total_headings": len(headings),
            "issues_count": len(issues),
            "issues": sorted(issues,
                             key=lambda x: {"warning": 0, "info": 1}
                             .get(x["severity"], 2)),
        }
