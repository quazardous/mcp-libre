"""
CommentService — comments, annotations, and track changes.
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class CommentService:
    """Comment and track-changes operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    # ==================================================================
    # Comments
    # ==================================================================

    def list_comments(self, file_path: str = None) -> Dict[str, Any]:
        """List all comments (excluding MCP-AI summaries)."""
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            para_ranges = self._registry.writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()

            comments = []
            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    author = ""
                if author == "MCP-AI":
                    continue

                content = field.getPropertyValue("Content")
                name = ""
                parent_name = ""
                resolved = False
                try:
                    name = field.getPropertyValue("Name")
                except Exception:
                    pass
                try:
                    parent_name = field.getPropertyValue("ParentName")
                except Exception:
                    pass
                try:
                    resolved = field.getPropertyValue("Resolved")
                except Exception:
                    pass

                date_str = ""
                try:
                    dt = field.getPropertyValue("DateTimeValue")
                    date_str = (f"{dt.Year:04d}-{dt.Month:02d}-{dt.Day:02d} "
                                f"{dt.Hours:02d}:{dt.Minutes:02d}")
                except Exception:
                    pass

                anchor = field.getAnchor()
                para_idx = self._registry.writer.find_paragraph_for_range(
                    anchor, para_ranges, text_obj)
                anchor_preview = anchor.getString()[:80]

                entry = {
                    "author": author,
                    "content": content,
                    "date": date_str,
                    "resolved": resolved,
                    "paragraph_index": para_idx,
                    "anchor_preview": anchor_preview,
                }
                if name:
                    entry["name"] = name
                if parent_name:
                    entry["parent_name"] = parent_name
                    entry["is_reply"] = True
                else:
                    entry["is_reply"] = False
                comments.append(entry)

            return {"success": True, "comments": comments,
                    "count": len(comments)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_comment(self, content: str, author: str = "AI Agent",
                    paragraph_index: int = None,
                    locator: str = None,
                    file_path: str = None) -> Dict[str, Any]:
        """Add a comment at a paragraph."""
        try:
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

            annotation = doc.createInstance(
                "com.sun.star.text.textfield.Annotation")
            annotation.setPropertyValue("Author", author)
            annotation.setPropertyValue("Content", content)

            doc_text = doc.getText()
            cursor = doc_text.createTextCursorByRange(target.getStart())
            doc_text.insertTextContent(cursor, annotation, False)

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "message": f"Comment added at paragraph {paragraph_index}",
                    "author": author}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def resolve_comment(self, comment_name: str,
                        resolution: str = "",
                        author: str = "AI Agent",
                        file_path: str = None) -> Dict[str, Any]:
        """Resolve a comment with an optional reason."""
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            target = None

            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    name = field.getPropertyValue("Name")
                except Exception:
                    continue
                if name == comment_name:
                    target = field
                    break

            if target is None:
                return {"success": False,
                        "error": f"Comment '{comment_name}' not found"}

            if resolution:
                reply = doc.createInstance(
                    "com.sun.star.text.textfield.Annotation")
                reply.setPropertyValue("Author", author)
                reply.setPropertyValue("Content", resolution)
                try:
                    reply.setPropertyValue("ParentName", comment_name)
                except Exception:
                    pass
                anchor = target.getAnchor()
                cursor = doc.getText().createTextCursorByRange(anchor)
                doc.getText().insertTextContent(cursor, reply, False)

            try:
                target.setPropertyValue("Resolved", True)
            except Exception:
                pass

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True,
                    "comment": comment_name,
                    "resolved": True,
                    "resolution": resolution}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_comment(self, comment_name: str,
                       file_path: str = None) -> Dict[str, Any]:
        """Delete a comment and all its replies."""
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            text_obj = doc.getText()
            deleted = 0

            to_delete = []
            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    name = field.getPropertyValue("Name")
                    parent = field.getPropertyValue("ParentName")
                except Exception:
                    continue
                if name == comment_name or parent == comment_name:
                    to_delete.append(field)

            for field in to_delete:
                text_obj.removeTextContent(field)
                deleted += 1

            if deleted > 0 and doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "deleted": deleted,
                    "comment": comment_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Task scanning (comment-as-ticket workflow)
    # ==================================================================

    _TASK_PREFIXES = ("TODO-AI", "FIX", "QUESTION", "VALIDATION", "NOTE")

    def scan_tasks(self, unresolved_only: bool = True,
                   prefix_filter: str = None,
                   file_path: str = None) -> Dict[str, Any]:
        """Scan comments for actionable task prefixes."""
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            para_ranges = self._registry.writer.get_paragraph_ranges(doc)
            text_obj = doc.getText()

            tasks = []
            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue

                # Skip replies
                try:
                    parent = field.getPropertyValue("ParentName")
                    if parent:
                        continue
                except Exception:
                    pass

                content = field.getPropertyValue("Content").strip()
                if not content:
                    continue

                # Match prefix
                matched_prefix = None
                task_body = content
                for pfx in self._TASK_PREFIXES:
                    if content.upper().startswith(pfx):
                        matched_prefix = pfx
                        # Strip prefix and separator (colon, dash, space)
                        rest = content[len(pfx):].lstrip(":- ")
                        task_body = rest
                        break

                if matched_prefix is None:
                    continue

                if prefix_filter and matched_prefix != prefix_filter.upper():
                    continue

                resolved = False
                try:
                    resolved = field.getPropertyValue("Resolved")
                except Exception:
                    pass

                if unresolved_only and resolved:
                    continue

                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    author = ""

                name = ""
                try:
                    name = field.getPropertyValue("Name")
                except Exception:
                    pass

                anchor = field.getAnchor()
                para_idx = self._registry.writer.find_paragraph_for_range(
                    anchor, para_ranges, text_obj)

                tasks.append({
                    "prefix": matched_prefix,
                    "body": task_body,
                    "author": author,
                    "resolved": resolved,
                    "comment_name": name,
                    "paragraph_index": para_idx,
                })

            return {"success": True, "tasks": tasks,
                    "count": len(tasks)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Workflow status (master dashboard comment)
    # ==================================================================

    _WORKFLOW_AUTHOR = "MCP-WORKFLOW"

    def get_workflow_status(self,
                            file_path: str = None) -> Dict[str, Any]:
        """Read the master workflow dashboard comment."""
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()

            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    continue
                if author != self._WORKFLOW_AUTHOR:
                    continue

                content = field.getPropertyValue("Content")
                name = ""
                try:
                    name = field.getPropertyValue("Name")
                except Exception:
                    pass

                # Parse key: value lines
                status = {}
                for line in content.splitlines():
                    line = line.strip()
                    if ":" in line:
                        k, v = line.split(":", 1)
                        status[k.strip()] = v.strip()

                return {"success": True, "found": True,
                        "comment_name": name, "raw": content,
                        "status": status}

            return {"success": True, "found": False,
                    "message": "No workflow dashboard comment found. "
                               "Use set_workflow_status to create one."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_workflow_status(self, content: str,
                            file_path: str = None) -> Dict[str, Any]:
        """Create or update the master workflow dashboard comment.

        Content should be key: value lines, e.g.:
            Phase: Rédaction
            Images: 3/10 insérées
            Annexes: En attente
        """
        try:
            doc = self._base.resolve_document(file_path)
            fields = doc.getTextFields()
            enum = fields.createEnumeration()
            existing = None

            while enum.hasMoreElements():
                field = enum.nextElement()
                if not field.supportsService(
                        "com.sun.star.text.textfield.Annotation"):
                    continue
                try:
                    author = field.getPropertyValue("Author")
                except Exception:
                    continue
                if author == self._WORKFLOW_AUTHOR:
                    existing = field
                    break

            if existing:
                existing.setPropertyValue("Content", content)
                fields.refresh()
                if doc.hasLocation():
                    self._base.store_doc(doc)
                return {"success": True, "action": "updated",
                        "content": content}

            # Create new at paragraph 0
            annotation = doc.createInstance(
                "com.sun.star.text.textfield.Annotation")
            annotation.setPropertyValue("Author", self._WORKFLOW_AUTHOR)
            annotation.setPropertyValue("Content", content)

            doc_text = doc.getText()
            para_enum = doc_text.createEnumeration()
            if para_enum.hasMoreElements():
                first_para = para_enum.nextElement()
                cursor = doc_text.createTextCursorByRange(
                    first_para.getStart())
                doc_text.insertTextContent(cursor, annotation, False)

            if doc.hasLocation():
                self._base.store_doc(doc)
            return {"success": True, "action": "created",
                    "content": content}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ==================================================================
    # Track Changes
    # ==================================================================

    def set_track_changes(self, enabled: bool,
                          file_path: str = None) -> Dict[str, Any]:
        """Enable or disable change tracking."""
        try:
            doc = self._base.resolve_document(file_path)
            doc.setPropertyValue("RecordChanges", enabled)
            return {"success": True, "record_changes": enabled}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_tracked_changes(self,
                            file_path: str = None) -> Dict[str, Any]:
        """List all tracked changes (redlines)."""
        try:
            doc = self._base.resolve_document(file_path)
            recording = doc.getPropertyValue("RecordChanges")

            if not hasattr(doc, 'getRedlines'):
                return {"success": False,
                        "error": "Document does not support redlines"}

            redlines = doc.getRedlines()
            enum = redlines.createEnumeration()
            changes = []
            while enum.hasMoreElements():
                redline = enum.nextElement()
                entry = {}
                for prop in ("RedlineType", "RedlineAuthor",
                             "RedlineComment", "RedlineIdentifier"):
                    try:
                        entry[prop] = redline.getPropertyValue(prop)
                    except Exception:
                        pass
                try:
                    dt = redline.getPropertyValue("RedlineDateTime")
                    entry["date"] = (f"{dt.Year:04d}-{dt.Month:02d}-"
                                     f"{dt.Day:02d} {dt.Hours:02d}:"
                                     f"{dt.Minutes:02d}")
                except Exception:
                    pass
                changes.append(entry)

            return {"success": True, "recording": recording,
                    "changes": changes, "count": len(changes)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def accept_all_changes(self,
                           file_path: str = None) -> Dict[str, Any]:
        """Accept all tracked changes."""
        try:
            doc = self._base.resolve_document(file_path)
            dispatcher = self._base.smgr.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", self._base.ctx)
            frame = doc.getCurrentController().getFrame()
            dispatcher.executeDispatch(
                frame, ".uno:AcceptAllTrackedChanges", "", 0, ())
            if doc.hasLocation():
                self._base.store_doc(doc)
            return {"success": True, "message": "All changes accepted"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def reject_all_changes(self,
                           file_path: str = None) -> Dict[str, Any]:
        """Reject all tracked changes."""
        try:
            doc = self._base.resolve_document(file_path)
            dispatcher = self._base.smgr.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", self._base.ctx)
            frame = doc.getCurrentController().getFrame()
            dispatcher.executeDispatch(
                frame, ".uno:RejectAllTrackedChanges", "", 0, ())
            if doc.hasLocation():
                self._base.store_doc(doc)
            return {"success": True, "message": "All changes rejected"}
        except Exception as e:
            return {"success": False, "error": str(e)}
