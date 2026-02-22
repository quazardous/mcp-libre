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

    # JSON Schema type → Python type(s)
    _TYPE_MAP = {
        "string": str,
        "integer": int,
        "number": (int, float),
        "boolean": bool,
        "array": list,
        "object": dict,
    }

    def validate(self, **kwargs):
        """Fast parameter check against the JSON schema.

        Returns ``(True, None)`` on success or ``(False, message)``
        on the first error found.  No UNO calls — pure Python.
        """
        schema = self.parameters
        required = schema.get("required", [])
        properties = schema.get("properties", {})

        for field in required:
            if field not in kwargs:
                return (False, "Missing required parameter: %s" % field)

        for key, value in kwargs.items():
            if key not in properties or value is None:
                continue
            expected = properties[key].get("type")
            if expected and expected in self._TYPE_MAP:
                if not isinstance(value, self._TYPE_MAP[expected]):
                    return (False,
                            "Parameter '%s' should be %s, got %s"
                            % (key, expected, type(value).__name__))
        return (True, None)

    @abstractmethod
    def execute(self, **kwargs) -> Dict[str, Any]:
        """Run the tool and return a JSON-serialisable dict."""
        ...
