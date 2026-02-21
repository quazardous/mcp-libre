"""
MCP Tool auto-discovery.

All McpTool subclasses in this package are automatically discovered
and registered by the MCP server.
"""

import importlib
import inspect
import logging
import pkgutil

logger = logging.getLogger(__name__)


def discover_tools():
    """Find all McpTool subclasses in the tools package."""
    from .base import McpTool

    tool_classes = []
    pkg = importlib.import_module(__name__)
    for _, mod_name, _ in pkgutil.iter_modules(pkg.__path__):
        if mod_name == "base":
            continue
        try:
            mod = importlib.import_module(f".{mod_name}", __name__)
        except Exception as exc:
            logger.warning("Failed to import tools.%s: %s", mod_name, exc)
            continue
        for obj in vars(mod).values():
            if (inspect.isclass(obj)
                    and issubclass(obj, McpTool)
                    and obj is not McpTool
                    and getattr(obj, "name", None) is not None):
                tool_classes.append(obj)
    logger.info("Discovered %d MCP tools", len(tool_classes))
    return tool_classes
