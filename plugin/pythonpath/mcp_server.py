"""
LibreOffice MCP Extension - MCP Server Module

Auto-discovers McpTool subclasses and dispatches tool calls to them.
Each tool class IS its own MCP definition (name, description, schema)
and delegates to the appropriate UNO service.
"""

import logging
import threading
import time
from typing import Dict, Any, List

from services import ServiceRegistry
from tools import discover_tools

logger = logging.getLogger(__name__)


# ── Document event listener (prewarm index on open) ─────────────────

_PREWARM_EVENTS = frozenset({
    "OnLoadFinished", "OnNew", "OnViewCreated",
})


def _register_doc_listener_on_main(server):
    """Register a GlobalEventBroadcaster listener + prewarm.

    MUST run on the VCL main thread.
    """
    import uno
    import unohelper
    from com.sun.star.document import XDocumentEventListener

    class _DocOpenListener(unohelper.Base, XDocumentEventListener):
        """Prewarms full-text index when a Writer document opens."""

        def documentEventOccured(self, event):
            if event.EventName not in _PREWARM_EVENTS:
                return
            doc = event.Source
            try:
                if not doc.supportsService(
                        "com.sun.star.text.TextDocument"):
                    return
            except Exception:
                return
            try:
                server.registry.writer.index.prewarm()
                logger.info("Prewarm on %s complete", event.EventName)
            except Exception as e:
                logger.warning("Prewarm on %s failed: %s",
                               event.EventName, e)

        def disposing(self, source):
            pass

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager

    # Try multiple service names — LO versions differ
    broadcaster = None
    for svc_name in (
        "com.sun.star.frame.theGlobalEventBroadcaster",
        "com.sun.star.frame.GlobalEventBroadcaster",
    ):
        broadcaster = smgr.createInstanceWithContext(svc_name, ctx)
        if broadcaster is not None:
            logger.info("Got broadcaster via %s", svc_name)
            break

    if broadcaster is None:
        raise RuntimeError("GlobalEventBroadcaster returned None "
                           "(smgr=%s)" % type(smgr).__name__)

    listener = _DocOpenListener()
    broadcaster.addDocumentEventListener(listener)
    server._doc_listener = listener
    logger.info("Document event listener registered (prewarm on open)")

    # Prewarm the already-open document right now (we're on main thread)
    try:
        server.registry.writer.index.prewarm()
        logger.info("Initial prewarm complete")
    except Exception as e:
        logger.info("Initial prewarm skipped (no document?): %s", e)


def _register_doc_listener(server):
    """Dispatch listener registration to the VCL main thread (with retries).

    The GlobalEventBroadcaster may not exist yet when LO is still
    initialising (Start Center loading).  Retry a few times with delay.
    """
    def _try_register():
        from main_thread_executor import execute_on_main_thread
        for attempt in range(5):
            if attempt > 0:
                time.sleep(3)
            try:
                execute_on_main_thread(
                    _register_doc_listener_on_main, server, timeout=15.0)
                return  # success
            except Exception as e:
                logger.info("Doc listener attempt %d/5 failed: %s",
                            attempt + 1, e)
        logger.warning("Could not register document listener after 5 attempts")

    threading.Thread(
        target=_try_register, daemon=True,
        name="mcp-doc-listener-init",
    ).start()


class LibreOfficeMCPServer:
    """MCP server with auto-discovered tools."""

    def __init__(self):
        self.registry = ServiceRegistry()

        # Discover and instantiate every McpTool subclass
        self.tools: Dict[str, Any] = {}
        for tool_cls in discover_tools():
            tool = tool_cls(self.registry)
            self.tools[tool.name] = tool

        logger.info("LibreOffice MCP Server ready — %d tools", len(self.tools))

    def execute_tool_sync(self, tool_name: str,
                          parameters: Dict[str, Any],
                          revision_comment: str = None) -> Dict[str, Any]:
        """Execute a tool on the VCL main thread.

        Called by ai_interface via MainThreadExecutor.

        If tracked changes are active, temporarily switches the LO
        author to "MCP" so revisions are attributed correctly.
        If *revision_comment* is provided, it is set on any new
        redlines created during this tool call.
        """
        tool = self.tools.get(tool_name)
        if tool is None:
            return {
                "success": False,
                "error": f"Unknown tool: {tool_name}",
                "available_tools": list(self.tools.keys()),
            }

        # -- Track-changes author switching --
        base = self.registry.base
        switched_author = False
        saved_author = None
        redline_ids_before = None
        try:
            doc = base.get_active_document()
            if doc and base.is_recording_changes(doc):
                saved_author = base.get_lo_author_parts()
                base.set_lo_author("MCP")
                switched_author = True
                if revision_comment:
                    redline_ids_before = base.get_redline_ids(doc)
        except Exception:
            pass

        t0 = time.perf_counter()
        try:
            result = tool.execute(**parameters)
        except Exception as e:
            logger.error("Tool '%s' error: %s", tool_name, e, exc_info=True)
            result = {"success": False, "error": str(e), "tool": tool_name}
        elapsed = time.perf_counter() - t0
        if isinstance(result, dict):
            result["_elapsed_ms"] = round(elapsed * 1000, 1)

        # -- Restore author & tag new redlines --
        if switched_author:
            try:
                base.restore_lo_author(saved_author)
            except Exception:
                pass
            if revision_comment and redline_ids_before is not None:
                try:
                    doc = base.get_active_document()
                    if doc:
                        tagged = base.set_new_redline_comments(
                            doc, redline_ids_before, revision_comment)
                        if isinstance(result, dict) and tagged:
                            result["_redlines_tagged"] = tagged
                except Exception:
                    pass

        # Prewarm full-text index (skip during batch)
        if not self.registry.batch_mode:
            try:
                self.registry.writer.index.prewarm()
            except Exception:
                pass

        return result

    def get_tool_list(self) -> List[Dict[str, Any]]:
        """Return MCP tools/list payload."""
        return [
            {
                "name": t.name,
                "description": t.description,
                "parameters": t.parameters,
            }
            for t in self.tools.values()
        ]


# Global singleton
mcp_server = None


def get_mcp_server() -> LibreOfficeMCPServer:
    """Get or create the global MCP server instance."""
    global mcp_server
    if mcp_server is None:
        mcp_server = LibreOfficeMCPServer()
        _register_doc_listener(mcp_server)
    return mcp_server
