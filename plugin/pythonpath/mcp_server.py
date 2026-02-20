"""
LibreOffice MCP Extension - MCP Server Module

This module implements an embedded MCP server that integrates with LibreOffice
via the UNO API, providing real-time document manipulation capabilities.
"""

import asyncio
import json
import logging
from typing import Dict, Any, Optional, List
from .uno_bridge import UNOBridge

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class LibreOfficeMCPServer:
    """Embedded MCP server for LibreOffice plugin"""
    
    def __init__(self):
        """Initialize the MCP server"""
        self.uno_bridge = UNOBridge()
        self.tools = {}
        self._register_tools()
        logger.info("LibreOffice MCP Server initialized")
    
    def _register_tools(self):
        """Register all available MCP tools"""
        
        # Document creation tools
        self.tools["create_document_live"] = {
            "description": "Create a new document in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["writer", "calc", "impress", "draw"],
                        "description": "Type of document to create",
                        "default": "writer"
                    }
                }
            },
            "handler": self.create_document_live
        }
        
        # Text manipulation tools
        self.tools["insert_text_live"] = {
            "description": "Insert text into the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "Text to insert"
                    },
                    "position": {
                        "type": "integer",
                        "description": "Position to insert at (optional, defaults to cursor position)"
                    }
                },
                "required": ["text"]
            },
            "handler": self.insert_text_live
        }
        
        # Document info tools
        self.tools["get_document_info_live"] = {
            "description": "Get information about the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_document_info_live
        }
        
        # Text formatting tools
        self.tools["format_text_live"] = {
            "description": "Apply formatting to selected text in active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "bold": {
                        "type": "boolean",
                        "description": "Apply bold formatting"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Apply italic formatting"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Apply underline formatting"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name"
                    }
                }
            },
            "handler": self.format_text_live
        }
        
        # Document saving tools
        self.tools["save_document_live"] = {
            "description": "Save the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to save document to (optional, saves to current location if not specified)"
                    }
                }
            },
            "handler": self.save_document_live
        }
        
        # Document export tools
        self.tools["export_document_live"] = {
            "description": "Export the currently active document to a different format",
            "parameters": {
                "type": "object",
                "properties": {
                    "export_format": {
                        "type": "string",
                        "enum": ["pdf", "docx", "doc", "odt", "txt", "rtf", "html"],
                        "description": "Format to export to"
                    },
                    "file_path": {
                        "type": "string",
                        "description": "Path to export document to"
                    }
                },
                "required": ["export_format", "file_path"]
            },
            "handler": self.export_document_live
        }
        
        # Content reading tools
        self.tools["get_text_content_live"] = {
            "description": "Get the text content of the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_text_content_live
        }
        
        # Document list tools
        self.tools["list_open_documents"] = {
            "description": "List all currently open documents in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_open_documents
        }
        
        # -------------------------------------------------------
        # Context-efficient document tools (tree navigation)
        # -------------------------------------------------------

        self.tools["open_document"] = {
            "description": "Open a document by file path in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "Absolute file path"}
                },
                "required": ["file_path"]
            },
            "handler": self._h_open_document
        }

        self.tools["get_document_tree"] = {
            "description": "Get document heading tree. Use depth to control "
                           "how many levels are returned.",
            "parameters": {
                "type": "object",
                "properties": {
                    "content_strategy": {
                        "type": "string",
                        "enum": ["none", "first_lines",
                                 "ai_summary_first", "full"],
                        "description": "What to show for body text "
                                       "(default: first_lines)"
                    },
                    "depth": {
                        "type": "integer",
                        "description": "Levels to return: 1=direct children, "
                                       "2=two levels, 0=unlimited (default: 1)"
                    },
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_get_document_tree
        }

        self.tools["get_heading_children"] = {
            "description": "Get children of a heading (drill down into tree). "
                           "Use heading_bookmark for stable navigation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "heading_para_index": {
                        "type": "integer",
                        "description": "Paragraph index of the heading"
                    },
                    "heading_bookmark": {
                        "type": "string",
                        "description": "Bookmark name (alternative to "
                                       "para_index, more stable)"
                    },
                    "content_strategy": {
                        "type": "string",
                        "enum": ["none", "first_lines",
                                 "ai_summary_first", "full"],
                        "description": "default: first_lines"
                    },
                    "depth": {
                        "type": "integer",
                        "description": "Sub-levels to include (default: 1)"
                    },
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_get_heading_children
        }

        self.tools["read_paragraphs"] = {
            "description": "Read a range of paragraphs by index",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_index": {"type": "integer",
                                    "description": "First paragraph index"},
                    "count": {"type": "integer",
                              "description": "Number to read (default: 10)"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["start_index"]
            },
            "handler": self._h_read_paragraphs
        }

        self.tools["get_paragraph_count"] = {
            "description": "Get total paragraph count",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_get_paragraph_count
        }

        self.tools["search_in_document"] = {
            "description": "Search text with paragraph context",
            "parameters": {
                "type": "object",
                "properties": {
                    "pattern": {"type": "string",
                                "description": "Search string or regex"},
                    "regex": {"type": "boolean",
                              "description": "Use regex (default: false)"},
                    "case_sensitive": {"type": "boolean",
                                      "description": "default: false"},
                    "max_results": {"type": "integer",
                                    "description": "default: 20"},
                    "context_paragraphs": {
                        "type": "integer",
                        "description": "Paragraphs of context (default: 1)"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["pattern"]
            },
            "handler": self._h_search_in_document
        }

        self.tools["replace_in_document"] = {
            "description": "Find and replace text (preserves formatting)",
            "parameters": {
                "type": "object",
                "properties": {
                    "search": {"type": "string", "description": "Text to find"},
                    "replace": {"type": "string",
                                "description": "Replacement text"},
                    "regex": {"type": "boolean", "description": "default: false"},
                    "case_sensitive": {"type": "boolean",
                                      "description": "default: false"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["search", "replace"]
            },
            "handler": self._h_replace_in_document
        }

        self.tools["insert_at_paragraph"] = {
            "description": "Insert text before or after a paragraph",
            "parameters": {
                "type": "object",
                "properties": {
                    "paragraph_index": {"type": "integer",
                                        "description": "Target paragraph"},
                    "text": {"type": "string", "description": "Text to insert"},
                    "position": {"type": "string", "enum": ["before", "after"],
                                 "description": "default: after"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["paragraph_index", "text"]
            },
            "handler": self._h_insert_at_paragraph
        }

        self.tools["add_ai_summary"] = {
            "description": "Add an AI annotation/summary to a heading",
            "parameters": {
                "type": "object",
                "properties": {
                    "para_index": {"type": "integer",
                                   "description": "Heading paragraph index"},
                    "summary": {"type": "string",
                                "description": "Summary text"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["para_index", "summary"]
            },
            "handler": self._h_add_ai_summary
        }

        self.tools["get_ai_summaries"] = {
            "description": "List all AI annotations in the document",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_get_ai_summaries
        }

        self.tools["remove_ai_summary"] = {
            "description": "Remove an AI annotation from a heading",
            "parameters": {
                "type": "object",
                "properties": {
                    "para_index": {"type": "integer",
                                   "description": "Heading paragraph index"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["para_index"]
            },
            "handler": self._h_remove_ai_summary
        }

        self.tools["list_sections"] = {
            "description": "List named text sections",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_list_sections
        }

        self.tools["read_section"] = {
            "description": "Read the content of a named section",
            "parameters": {
                "type": "object",
                "properties": {
                    "section_name": {"type": "string",
                                     "description": "Section name"},
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["section_name"]
            },
            "handler": self._h_read_section
        }

        self.tools["list_bookmarks"] = {
            "description": "List all bookmarks in the document",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_list_bookmarks
        }

        self.tools["resolve_bookmark"] = {
            "description": "Resolve a bookmark to its current paragraph index "
                           "(bookmarks are stable across edits)",
            "parameters": {
                "type": "object",
                "properties": {
                    "bookmark_name": {
                        "type": "string",
                        "description": "Bookmark name (e.g. _mcp_a1b2c3d4)"
                    },
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                },
                "required": ["bookmark_name"]
            },
            "handler": self._h_resolve_bookmark
        }

        self.tools["get_page_count"] = {
            "description": "Get the document page count",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string",
                                  "description": "File path (optional)"}
                }
            },
            "handler": self._h_get_page_count
        }

        logger.info(f"Registered {len(self.tools)} MCP tools")
    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute an MCP tool
        
        Args:
            tool_name: Name of the tool to execute
            parameters: Parameters for the tool
            
        Returns:
            Result dictionary
        """
        try:
            if tool_name not in self.tools:
                return {
                    "success": False,
                    "error": f"Unknown tool: {tool_name}",
                    "available_tools": list(self.tools.keys())
                }
            
            tool = self.tools[tool_name]
            handler = tool["handler"]
            
            # Execute the tool handler
            result = handler(**parameters)
            
            logger.info(f"Executed tool '{tool_name}' successfully")
            return result
            
        except Exception as e:
            logger.error(f"Error executing tool '{tool_name}': {e}")
            return {
                "success": False,
                "error": str(e),
                "tool": tool_name,
                "parameters": parameters
            }
    
    def get_tool_list(self) -> List[Dict[str, Any]]:
        """Get list of available tools with their descriptions"""
        return [
            {
                "name": name,
                "description": tool["description"],
                "parameters": tool["parameters"]
            }
            for name, tool in self.tools.items()
        ]
    
    # Tool handler methods
    
    def create_document_live(self, doc_type: str = "writer") -> Dict[str, Any]:
        """Create a new document in LibreOffice"""
        try:
            doc = self.uno_bridge.create_document(doc_type)
            doc_info = self.uno_bridge.get_document_info(doc)
            
            return {
                "success": True,
                "message": f"Created new {doc_type} document",
                "document_info": doc_info
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def insert_text_live(self, text: str, position: Optional[int] = None) -> Dict[str, Any]:
        """Insert text into the currently active document"""
        return self.uno_bridge.insert_text(text, position)
    
    def get_document_info_live(self) -> Dict[str, Any]:
        """Get information about the currently active document"""
        doc_info = self.uno_bridge.get_document_info()
        if "error" in doc_info:
            return {"success": False, **doc_info}
        else:
            return {"success": True, "document_info": doc_info}
    
    def format_text_live(self, **formatting) -> Dict[str, Any]:
        """Apply formatting to selected text"""
        return self.uno_bridge.format_text(formatting)
    
    def save_document_live(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the currently active document"""
        return self.uno_bridge.save_document(file_path=file_path)
    
    def export_document_live(self, export_format: str, file_path: str) -> Dict[str, Any]:
        """Export the currently active document"""
        return self.uno_bridge.export_document(export_format, file_path)
    
    def get_text_content_live(self) -> Dict[str, Any]:
        """Get text content of the currently active document"""
        return self.uno_bridge.get_text_content()
    
    # -- Context-efficient tool handlers --

    def _h_open_document(self, file_path: str) -> Dict[str, Any]:
        result = self.uno_bridge.open_document(file_path)
        return {k: v for k, v in result.items() if k != "doc"}

    def _h_get_document_tree(self, content_strategy: str = "first_lines",
                             depth: int = 1,
                             file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.get_document_tree(
            content_strategy, depth, file_path)

    def _h_get_heading_children(self, heading_para_index: int = None,
                                heading_bookmark: str = None,
                                content_strategy: str = "first_lines",
                                depth: int = 1,
                                file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.get_heading_children(
            heading_para_index, heading_bookmark,
            content_strategy, depth, file_path)

    def _h_read_paragraphs(self, start_index: int, count: int = 10,
                           file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.read_paragraphs(start_index, count, file_path)

    def _h_get_paragraph_count(self, file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.get_paragraph_count(file_path)

    def _h_search_in_document(self, pattern: str, regex: bool = False,
                              case_sensitive: bool = False,
                              max_results: int = 20,
                              context_paragraphs: int = 1,
                              file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.search_document(
            pattern, regex, case_sensitive, max_results,
            context_paragraphs, file_path)

    def _h_replace_in_document(self, search: str, replace: str,
                               regex: bool = False,
                               case_sensitive: bool = False,
                               file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.replace_in_document(
            search, replace, regex, case_sensitive, file_path)

    def _h_insert_at_paragraph(self, paragraph_index: int, text: str,
                               position: str = "after",
                               file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.insert_at_paragraph(
            paragraph_index, text, position, file_path)

    def _h_add_ai_summary(self, para_index: int, summary: str,
                           file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.add_ai_summary(para_index, summary, file_path)

    def _h_get_ai_summaries(self, file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.get_ai_summaries(file_path)

    def _h_remove_ai_summary(self, para_index: int,
                              file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.remove_ai_summary(para_index, file_path)

    def _h_list_sections(self, file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.list_sections(file_path)

    def _h_read_section(self, section_name: str,
                         file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.read_section(section_name, file_path)

    def _h_list_bookmarks(self, file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.list_bookmarks(file_path)

    def _h_resolve_bookmark(self, bookmark_name: str,
                             file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.resolve_bookmark(bookmark_name, file_path)

    def _h_get_page_count(self, file_path: str = None) -> Dict[str, Any]:
        return self.uno_bridge.get_page_count(file_path)

    def list_open_documents(self) -> Dict[str, Any]:
        """List all open documents in LibreOffice"""
        try:
            desktop = self.uno_bridge.desktop
            documents = []
            
            # Get all open documents
            frames = desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController()
                if controller:
                    doc = controller.getModel()
                    if doc:
                        doc_info = self.uno_bridge.get_document_info(doc)
                        documents.append(doc_info)
            
            return {
                "success": True,
                "documents": documents,
                "count": len(documents)
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}


# Global instance
mcp_server = None

def get_mcp_server() -> LibreOfficeMCPServer:
    """Get or create the global MCP server instance"""
    global mcp_server
    if mcp_server is None:
        mcp_server = LibreOfficeMCPServer()
    return mcp_server
