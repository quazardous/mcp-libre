"""
McpTool — base class for all MCP tools.

Each tool class IS its own MCP definition:
  - name, description, parameters (JSON Schema)
  - execute(**kwargs) → dict

The MCP server auto-discovers all subclasses and registers them.
No separate handler or registration step needed.
"""

from abc import ABC, abstractmethod
from typing import Any, Dict


class McpTool(ABC):
    """Contract for an MCP tool.

    Subclass attributes (required):
        name:        MCP tool name (e.g. "get_document_tree")
        description: Human-readable description for the AI agent
        parameters:  JSON Schema dict (inputSchema)

    The constructor receives the ServiceRegistry so every tool
    can access any UNO service it needs.
    """

    name: str = None
    description: str = ""
    parameters: dict = {"type": "object", "properties": {}}

    def __init__(self, services):
        self.services = services

    @abstractmethod
    def execute(self, **kwargs) -> Dict[str, Any]:
        """Run the tool and return a JSON-serialisable dict."""
        ...
