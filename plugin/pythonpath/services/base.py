"""
BaseService — shared UNO infrastructure for all domain services.

Holds the UNO component context, desktop, toolkit, and provides
cross-cutting helpers: document resolution, locator parsing,
GUI yield, page caching.
"""

import logging
import os
import uuid
from typing import Any, Dict, List, Optional, Tuple

import uno
from com.sun.star.beans import PropertyValue

logger = logging.getLogger(__name__)


class BaseService:
    """Shared UNO infrastructure injected into every domain service."""

    def __init__(self):
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        self._toolkit = self.smgr.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", self.ctx)
        self._page_cache: Dict[Tuple[str, str], int] = {}
        self._yield_counter = 0
        self._registry = None   # set by ServiceRegistry after construction
        logger.info("BaseService initialized")

    # ------------------------------------------------------------------
    # GUI yield — keep LO responsive during long operations
    # ------------------------------------------------------------------

    def yield_to_gui(self, every: int = 50):
        """Process pending VCL events to keep GUI responsive.

        Call this inside tight loops.  The actual reschedule only fires
        every *every* calls to amortise overhead.
        """
        self._yield_counter += 1
        if self._yield_counter % every != 0:
            return
        try:
            if hasattr(self._toolkit, "processEventsToIdle"):
                self._toolkit.processEventsToIdle()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Document resolution
    # ------------------------------------------------------------------

    def get_active_document(self) -> Optional[Any]:
        """Get the currently active document (may be None)."""
        try:
            return self.desktop.getCurrentComponent()
        except Exception:
            return None

    def find_open_document(self, file_url: str) -> Optional[Any]:
        """Find an already-open document by its URL (normalized)."""
        try:
            components = self.desktop.getComponents()
            if components is None:
                return None
            import urllib.parse
            norm = urllib.parse.unquote(file_url).lower().rstrip("/")
            enum = components.createEnumeration()
            while enum.hasMoreElements():
                doc = enum.nextElement()
                if not hasattr(doc, "getURL"):
                    continue
                doc_url = urllib.parse.unquote(
                    doc.getURL()).lower().rstrip("/")
                if doc_url == norm:
                    return doc
            return None
        except Exception:
            return None

    def open_document(self, file_path: str,
                      force: bool = False) -> Dict[str, Any]:
        """Open a document by path, or return it if already open."""
        try:
            file_url = uno.systemPathToFileUrl(file_path)
            existing = self.find_open_document(file_url)
            if existing is not None:
                return {"success": True, "doc": existing,
                        "url": file_url, "already_open": True}

            # Same filename at different path → informational warning
            same_name_url = None
            target_name = os.path.basename(file_path).lower()
            try:
                components = self.desktop.getComponents()
                if components:
                    import urllib.parse
                    enum = components.createEnumeration()
                    while enum.hasMoreElements():
                        doc = enum.nextElement()
                        if hasattr(doc, "getURL"):
                            doc_name = os.path.basename(
                                urllib.parse.unquote(
                                    doc.getURL())).lower()
                            if doc_name == target_name:
                                same_name_url = doc.getURL()
                                break
            except Exception:
                pass

            props = (
                PropertyValue("Hidden", 0, False, 0),
                PropertyValue("ReadOnly", 0, False, 0),
            )
            target = "_blank" if force else "_default"
            doc = self.desktop.loadComponentFromURL(
                file_url, target, 0, props)
            if doc is None:
                return {"success": False,
                        "error": f"Failed to load document: {file_path}"}

            result = {"success": True, "doc": doc,
                      "url": file_url, "already_open": False}
            if same_name_url:
                result["warning"] = (
                    f"Another '{target_name}' is already open "
                    f"from a different path: {same_name_url}")
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def close_document(self, file_path: str) -> Dict[str, Any]:
        """Close a document by file path. Does not save."""
        try:
            file_url = uno.systemPathToFileUrl(file_path)
            doc = self.find_open_document(file_url)
            if doc is None:
                return {"success": True,
                        "message": "Document was not open"}
            doc.setModified(False)
            doc.close(True)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_document(self, doc_type: str = "writer") -> Any:
        """Create a new document (writer, calc, impress, draw)."""
        url_map = {
            "writer": "private:factory/swriter",
            "calc": "private:factory/scalc",
            "impress": "private:factory/simpress",
            "draw": "private:factory/sdraw",
        }
        url = url_map.get(doc_type, "private:factory/swriter")
        return self.desktop.loadComponentFromURL(url, "_blank", 0, ())

    def resolve_document(self, file_path: str = None) -> Any:
        """Open by path or return the active document. Raises on failure."""
        if file_path:
            result = self.open_document(file_path)
            if not result["success"]:
                raise RuntimeError(result["error"])
            return result["doc"]
        doc = self.get_active_document()
        if doc is None:
            raise RuntimeError(
                "No active document and no file path provided")
        return doc

    # ------------------------------------------------------------------
    # Locator resolution
    # ------------------------------------------------------------------

    def resolve_locator(self, doc, locator: str) -> Dict[str, Any]:
        """Parse 'type:value' locator and resolve to document position.

        Simple locators handled here; Writer-specific ones are
        delegated to WriterService via the registry.
        """
        loc_type, sep, loc_value = locator.partition(":")
        if not sep:
            raise ValueError(
                f"Invalid locator format: '{locator}'. "
                "Expected 'type:value'.")

        # -- Simple locators --
        if loc_type == "paragraph":
            return {"para_index": int(loc_value)}
        if loc_type in ("cell", "range", "sheet"):
            return {"loc_type": loc_type, "loc_value": loc_value}
        if loc_type == "slide":
            return {"slide_index": int(loc_value)}

        # -- Writer-specific locators (delegate) --
        if loc_type in ("bookmark", "page", "section", "heading",
                        "heading_text"):
            return self._registry.writer.resolve_writer_locator(
                doc, loc_type, loc_value)

        raise ValueError(f"Unknown locator type: '{loc_type}'")

    # ------------------------------------------------------------------
    # Page caching
    # ------------------------------------------------------------------

    def doc_key(self, doc) -> str:
        """Stable key for a document (URL or id)."""
        try:
            return doc.getURL() or str(id(doc))
        except Exception:
            return str(id(doc))

    def resolve_page(self, doc, obj_name: str, anchor) -> Optional[int]:
        """Cached page number for a named object."""
        key = (self.doc_key(doc), obj_name)
        if key in self._page_cache:
            return self._page_cache[key]
        try:
            page = self.get_page_for_range(doc, anchor)
            self._page_cache[key] = page
            return page
        except Exception:
            return None

    def invalidate_page_cache(self, file_path: str = None):
        """Clear page cache (all or for a specific document)."""
        if file_path is None:
            self._page_cache.clear()
        else:
            self._page_cache = {
                k: v for k, v in self._page_cache.items()
                if k[0] != file_path}

    def store_doc(self, doc):
        """Save document and invalidate its page cache."""
        doc.store()
        dk = self.doc_key(doc)
        self._page_cache = {
            k: v for k, v in self._page_cache.items()
            if k[0] != dk}

    def get_page_for_range(self, doc, text_range) -> int:
        """Page number for a text range using ViewCursor."""
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        vc.gotoRange(text_range, False)
        return vc.getPage()

    def get_page_for_paragraph(self, doc, para_index: int) -> int:
        """Page number for a paragraph by index."""
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        for _ in range(para_index):
            if not cursor.gotoNextParagraph(False):
                break
        return self.get_page_for_range(doc, cursor)

    def anchor_para_index(self, doc, anchor) -> Optional[int]:
        """Paragraph index for a text anchor (handles frame-anchored objects)."""
        main_text = doc.getText()
        rng = anchor
        try:
            anchor_text = anchor.getText()
            if anchor_text != main_text:
                if hasattr(doc, "getTextFrames"):
                    frames = doc.getTextFrames()
                    for fname in frames.getElementNames():
                        frame = frames.getByName(fname)
                        if frame.getText() == anchor_text:
                            rng = frame.getAnchor()
                            break
        except Exception:
            pass
        try:
            tc = main_text.createTextCursorByRange(rng)
            idx = 0
            while tc.gotoPreviousParagraph(False):
                idx += 1
            return idx
        except Exception:
            return None

    def annotate_pages(self, nodes: list, doc):
        """Recursively add 'page' field to heading tree nodes.

        Uses lockControllers + cursor save/restore to prevent
        visible viewport jumping while resolving page numbers.
        """
        controller = doc.getCurrentController()
        vc = controller.getViewCursor()
        saved = doc.getText().createTextCursorByRange(vc.getStart())
        doc.lockControllers()
        try:
            self._annotate_pages_inner(nodes, doc)
        finally:
            vc.gotoRange(saved, False)
            doc.unlockControllers()

    def _annotate_pages_inner(self, nodes: list, doc):
        for node in nodes:
            try:
                pi = node.get("para_index")
                if pi is not None:
                    node["page"] = self.get_page_for_paragraph(doc, pi)
            except Exception:
                pass
            if "children" in node:
                self._annotate_pages_inner(node["children"], doc)

    # ------------------------------------------------------------------
    # Document type helpers
    # ------------------------------------------------------------------

    @staticmethod
    def is_writer(doc) -> bool:
        return doc.supportsService("com.sun.star.text.TextDocument")

    @staticmethod
    def is_calc(doc) -> bool:
        return doc.supportsService(
            "com.sun.star.sheet.SpreadsheetDocument")

    @staticmethod
    def is_impress(doc) -> bool:
        return doc.supportsService(
            "com.sun.star.presentation.PresentationDocument")

    @staticmethod
    def get_document_type(doc) -> str:
        if BaseService.is_writer(doc):
            return "writer"
        if BaseService.is_calc(doc):
            return "calc"
        if BaseService.is_impress(doc):
            return "impress"
        return "unknown"

    # ------------------------------------------------------------------
    # Document metadata (type-agnostic)
    # ------------------------------------------------------------------

    def get_document_properties(self,
                                file_path: str = None) -> Dict[str, Any]:
        """Read document metadata (title, author, subject, etc.)."""
        try:
            doc = self.resolve_document(file_path)
            props = doc.getDocumentProperties()
            result = {
                "success": True,
                "title": props.Title,
                "author": props.Author,
                "subject": props.Subject,
                "description": props.Description,
                "keywords": list(props.Keywords) if props.Keywords else [],
                "generator": props.Generator,
            }
            try:
                cd = props.CreationDate
                result["creation_date"] = (
                    f"{cd.Year:04d}-{cd.Month:02d}-{cd.Day:02d}")
            except Exception:
                pass
            try:
                md = props.ModificationDate
                result["modification_date"] = (
                    f"{md.Year:04d}-{md.Month:02d}-{md.Day:02d}")
            except Exception:
                pass
            try:
                user_props = props.getUserDefinedProperties()
                info = user_props.getPropertySetInfo()
                custom = {}
                for prop in info.getProperties():
                    try:
                        custom[prop.Name] = str(
                            user_props.getPropertyValue(prop.Name))
                    except Exception:
                        pass
                if custom:
                    result["custom_properties"] = custom
            except Exception:
                pass
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_document_properties(self, title: str = None,
                                author: str = None,
                                subject: str = None,
                                description: str = None,
                                keywords: list = None,
                                file_path: str = None) -> Dict[str, Any]:
        """Update document metadata."""
        try:
            doc = self.resolve_document(file_path)
            props = doc.getDocumentProperties()
            updated = []
            if title is not None:
                props.Title = title
                updated.append("title")
            if author is not None:
                props.Author = author
                updated.append("author")
            if subject is not None:
                props.Subject = subject
                updated.append("subject")
            if description is not None:
                props.Description = description
                updated.append("description")
            if keywords is not None:
                props.Keywords = tuple(keywords)
                updated.append("keywords")
            if updated and doc.hasLocation():
                self.store_doc(doc)
            return {"success": True, "updated_fields": updated}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Document protection (type-agnostic)
    # ------------------------------------------------------------------

    def set_document_protection(self, enabled: bool,
                                file_path: str = None) -> Dict[str, Any]:
        """Lock/unlock the document UI. MCP/UNO can still edit."""
        try:
            doc = self.resolve_document(file_path)
            if enabled:
                if not doc.isProtected():
                    doc.protect("")
            else:
                if doc.isProtected():
                    doc.unprotect("")
            return {"success": True, "protected": enabled}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # File operations (type-agnostic)
    # ------------------------------------------------------------------

    def save_document_as(self, target_path: str,
                         file_path: str = None) -> Dict[str, Any]:
        """Save/duplicate a document under a new name."""
        try:
            doc = self.resolve_document(file_path)
            target_url = uno.systemPathToFileUrl(target_path)
            doc.storeToURL(target_url, ())
            return {"success": True,
                    "message": f"Saved to {target_path}",
                    "url": target_url}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_document(self, file_path: str = None) -> Dict[str, Any]:
        """Save the active document."""
        try:
            doc = self.resolve_document(file_path)
            if doc.hasLocation():
                self.store_doc(doc)
                return {"success": True, "message": "Document saved"}
            return {"success": False,
                    "error": "Document has no location (use save_as)"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Recent documents
    # ------------------------------------------------------------------

    def get_recent_documents(self, max_count: int = 20) -> Dict[str, Any]:
        """Get recently opened documents from LO history."""
        try:
            import urllib.parse
            config_provider = self.smgr.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider",
                self.ctx)
            node_path = PropertyValue()
            node_path.Name = "nodepath"
            node_path.Value = "/org.openoffice.Office.Common/History"
            access = config_provider.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess",
                (node_path,))
            pick_list = access.getByName("PickList")
            names = pick_list.getElementNames()

            docs = []
            for name in names[:max_count]:
                try:
                    item = pick_list.getByName(name)
                    url = name
                    title = ""
                    try:
                        title = item.getByName("Title")
                    except Exception:
                        pass
                    path = ""
                    try:
                        path = urllib.parse.unquote(
                            url.replace("file:///", ""))
                    except Exception:
                        pass
                    docs.append({"url": url, "title": title, "path": path})
                except Exception:
                    pass

            return {"success": True, "documents": docs,
                    "count": len(docs)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Tracked changes — author switching & redline comments
    # ------------------------------------------------------------------

    def _user_profile_access(self, writable=False):
        """Get ConfigurationAccess for the LO user profile."""
        config_provider = self.smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider",
            self.ctx)
        node_path = PropertyValue()
        node_path.Name = "nodepath"
        node_path.Value = "/org.openoffice.UserProfile/Data"
        mode = ("com.sun.star.configuration.ConfigurationUpdateAccess"
                if writable else
                "com.sun.star.configuration.ConfigurationAccess")
        return config_provider.createInstanceWithArguments(
            mode, (node_path,))

    def get_lo_author_parts(self):
        """Return (givenname, sn, initials) from the LO user profile."""
        try:
            access = self._user_profile_access()
            return (access.getByName("givenname") or "",
                    access.getByName("sn") or "",
                    access.getByName("initials") or "")
        except Exception:
            return ("", "", "")

    def set_lo_author(self, name, initials=None):
        """Set the LO user identity (affects tracked-change authorship).

        Splits *name* on first space into givenname + sn.
        """
        try:
            parts = name.split(" ", 1)
            first = parts[0]
            last = parts[1] if len(parts) > 1 else ""
            access = self._user_profile_access(writable=True)
            access.replaceByName("givenname", first)
            access.replaceByName("sn", last)
            if initials is not None:
                access.replaceByName("initials", initials)
            access.commitChanges()
            return True
        except Exception as e:
            logger.warning("set_lo_author failed: %s", e)
            return False

    def restore_lo_author(self, saved):
        """Restore author from a (givenname, sn, initials) tuple."""
        try:
            access = self._user_profile_access(writable=True)
            access.replaceByName("givenname", saved[0])
            access.replaceByName("sn", saved[1])
            access.replaceByName("initials", saved[2])
            access.commitChanges()
        except Exception as e:
            logger.warning("restore_lo_author failed: %s", e)

    def is_recording_changes(self, doc=None):
        """Check if tracked changes are being recorded."""
        try:
            if doc is None:
                doc = self.get_active_document()
            if doc is None:
                return False
            return bool(doc.getPropertyValue("RecordChanges"))
        except Exception:
            return False

    def get_redline_ids(self, doc):
        """Return a set of redline identifiers currently in the document."""
        ids = set()
        try:
            redlines = doc.getRedlines()
            enum = redlines.createEnumeration()
            while enum.hasMoreElements():
                rl = enum.nextElement()
                try:
                    ids.add(rl.getPropertyValue("RedlineIdentifier"))
                except Exception:
                    pass
        except Exception:
            pass
        return ids

    def set_new_redline_comments(self, doc, old_ids, comment):
        """Set comment on redlines that were NOT in *old_ids*."""
        try:
            redlines = doc.getRedlines()
            enum = redlines.createEnumeration()
            tagged = 0
            while enum.hasMoreElements():
                rl = enum.nextElement()
                try:
                    rl_id = rl.getPropertyValue("RedlineIdentifier")
                    if rl_id not in old_ids:
                        rl.setPropertyValue("RedlineComment", comment)
                        tagged += 1
                except Exception:
                    pass
            return tagged
        except Exception:
            return 0
