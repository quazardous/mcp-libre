"""
LibreOffice MCP Extension - MCP Server Module

Auto-discovers McpTool subclasses and dispatches tool calls to them.
Each tool class IS its own MCP definition (name, description, schema)
and delegates to the appropriate UNO service.
"""

import logging
from typing import Dict, Any, List

from services import ServiceRegistry
from tools import discover_tools

logger = logging.getLogger(__name__)


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
                          parameters: Dict[str, Any]) -> Dict[str, Any]:
        """Execute a tool on the VCL main thread.

        Called by ai_interface via MainThreadExecutor.
        """
        tool = self.tools.get(tool_name)
        if tool is None:
            return {
                "success": False,
                "error": f"Unknown tool: {tool_name}",
                "available_tools": list(self.tools.keys()),
            }
        try:
            return tool.execute(**parameters)
        except Exception as e:
            logger.error("Tool '%s' error: %s", tool_name, e, exc_info=True)
            return {"success": False, "error": str(e), "tool": tool_name}

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
    return mcp_server
