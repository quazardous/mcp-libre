"""
LibreOffice MCP Server — orchestrator.

Creates the FastMCP instance, initialises the backend, registers all tools
and resources, and exposes ``main()`` as the entry point.
"""

import sys
from pathlib import Path
from typing import List

from mcp.server.fastmcp import FastMCP

from .plugin import call_plugin
from .backends import get_backend
from .tools import register_all

# ---------------------------------------------------------------------------
# FastMCP instance + backend
# ---------------------------------------------------------------------------

mcp = FastMCP("LibreOffice MCP Server")
backend = get_backend(call_plugin)

# ---------------------------------------------------------------------------
# Register all tools
# ---------------------------------------------------------------------------

register_all(mcp, backend, call_plugin)

# ---------------------------------------------------------------------------
# Resources
# ---------------------------------------------------------------------------

@mcp.resource("documents://")
def list_documents() -> List[str]:
    """List all LibreOffice documents in common locations"""
    documents: List[str] = []
    search_paths = [
        Path.home() / "Documents",
        Path.home() / "Desktop",
        Path.cwd(),
    ]
    extensions = {
        '.odt', '.ods', '.odp', '.odg',
        '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
    }
    for sp in search_paths:
        if sp.exists():
            for ext in extensions:
                for doc in sp.rglob(f'*{ext}'):
                    if doc.is_file():
                        documents.append(str(doc))
    return sorted(documents)


@mcp.resource("document://{path}")
def get_document_content(path: str) -> str:
    """Get the text content of a specific document"""
    try:
        actual_path = path if path.startswith('/') else '/' + path
        content = backend.read_document_text(actual_path)
        return (f"Document: {Path(actual_path).name}\n"
                f"Words: {content.word_count}, "
                f"Characters: {content.char_count}\n\n"
                + content.content)
    except Exception as e:
        return f"Error reading document {path}: {e}"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Run the LibreOffice MCP server."""
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg in ("--help", "-h"):
            print("LibreOffice MCP Server")
            print("=" * 30)
            print("Usage:")
            print("  python src/main.py        # Start MCP server (stdio)")
            print("  python src/main.py --help  # Show this help")
            return
        if arg == "--version":
            print("LibreOffice MCP Server v1.0.0")
            return

    print("Starting LibreOffice MCP Server...", file=sys.stderr)
    print("Running in MCP protocol mode (stdio)", file=sys.stderr)

    try:
        mcp.run()
    except KeyboardInterrupt:
        print("\nLibreOffice MCP Server stopped", file=sys.stderr)
    except Exception as e:
        print(f"\nServer error: {e}", file=sys.stderr)
        sys.exit(1)
