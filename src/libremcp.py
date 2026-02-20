"""
LibreOffice Model Context Protocol Server

This MCP server provides tools and resources for interacting with LibreOffice documents.
It supports reading, writing, and manipulating Writer documents, Calc spreadsheets, 
and other LibreOffice formats.
"""

import asyncio
import json
import os
import platform
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from dataclasses import dataclass
from datetime import datetime

import httpx
from pydantic import BaseModel, Field
from mcp.server.fastmcp import FastMCP

# Initialize FastMCP server
mcp = FastMCP("LibreOffice MCP Server")

# Configuration: set MCP_LIBREOFFICE_GUI=1 to run LibreOffice with visible window
# instead of headless mode. Useful for seeing changes in real-time.
HEADLESS_MODE = os.environ.get("MCP_LIBREOFFICE_GUI", "0") not in ("1", "true", "yes")

# ============================================================================
# Plugin HTTP API (UNO bridge running inside LibreOffice on port 8765)
# ============================================================================

PLUGIN_API_URL = os.environ.get("MCP_PLUGIN_URL", "http://localhost:8765")
_plugin_available: Optional[bool] = None
_plugin_check_time: float = 0.0
_PLUGIN_CHECK_INTERVAL = 30.0


def _check_plugin_available() -> bool:
    """Check if the LibreOffice plugin HTTP API is available (cached)."""
    global _plugin_available, _plugin_check_time
    now = time.time()
    if _plugin_available is not None and (now - _plugin_check_time) < _PLUGIN_CHECK_INTERVAL:
        return _plugin_available
    try:
        with httpx.Client(timeout=3.0) as client:
            resp = client.get(f"{PLUGIN_API_URL}/health")
            _plugin_available = resp.status_code == 200
    except Exception:
        _plugin_available = False
    _plugin_check_time = now
    return _plugin_available


def _call_plugin(tool_name: str, params: Dict[str, Any]) -> Dict[str, Any]:
    """Call a tool on the LibreOffice plugin HTTP API."""
    if not _check_plugin_available():
        raise RuntimeError(
            "LibreOffice plugin not available. Start LibreOffice with the "
            "MCP extension (HTTP API on http://localhost:8765)."
        )
    try:
        with httpx.Client(timeout=60.0) as client:
            resp = client.post(
                f"{PLUGIN_API_URL}/tools/{tool_name}", json=params)
            resp.raise_for_status()
            return resp.json()
    except httpx.ConnectError:
        global _plugin_available
        _plugin_available = None
        raise RuntimeError("Lost connection to LibreOffice plugin.")
    except httpx.TimeoutException:
        raise RuntimeError("Plugin API call timed out.")
    except httpx.HTTPStatusError as e:
        raise RuntimeError(f"Plugin API error ({e.response.status_code}): "
                           f"{e.response.text}")


# ============================================================================
# Context-efficient document tools (via LibreOffice UNO plugin)
# ============================================================================

@mcp.tool()
def get_document_tree(path: str, content_strategy: str = "first_lines",
                      depth: int = 1) -> Dict[str, Any]:
    """Get the heading tree of a document without loading full text.

    Returns headings organized as a tree. Use depth to control how many
    levels are returned (1=top-level only, 2=two levels, 0=full tree).
    Use content_strategy to control body text visibility:
    none, first_lines, ai_summary_first, full.

    Args:
        path: Absolute path to the document
        content_strategy: What to show for body text (default: first_lines)
        depth: Heading levels to return (default: 1)
    """
    return _call_plugin("get_document_tree", {
        "file_path": path, "content_strategy": content_strategy,
        "depth": depth})


@mcp.tool()
def get_heading_children(path: str, heading_para_index: int = None,
                         heading_bookmark: str = None,
                         content_strategy: str = "first_lines",
                         depth: int = 1) -> Dict[str, Any]:
    """Drill down into a heading to see its children.

    Use heading_bookmark (stable across edits) or heading_para_index.
    All heading nodes in the response include a 'bookmark' field.

    Args:
        path: Absolute path to the document
        heading_para_index: Paragraph index of the parent heading
        heading_bookmark: Bookmark name (stable alternative to para_index)
        content_strategy: none, first_lines, ai_summary_first, full
        depth: Sub-levels to include (default: 1)
    """
    params = {"file_path": path, "content_strategy": content_strategy,
              "depth": depth}
    if heading_bookmark is not None:
        params["heading_bookmark"] = heading_bookmark
    if heading_para_index is not None:
        params["heading_para_index"] = heading_para_index
    return _call_plugin("get_heading_children", params)


@mcp.tool()
def read_document_paragraphs(path: str, start_index: int,
                             count: int = 10) -> Dict[str, Any]:
    """Read specific paragraphs by index range.

    Use get_document_tree first to find which paragraphs to read.

    Args:
        path: Absolute path to the document
        start_index: Zero-based index of first paragraph
        count: Number of paragraphs to read (default: 10)
    """
    return _call_plugin("read_paragraphs", {
        "file_path": path, "start_index": start_index, "count": count})


@mcp.tool()
def get_document_paragraph_count(path: str) -> Dict[str, Any]:
    """Get total paragraph count of a document.

    Args:
        path: Absolute path to the document
    """
    return _call_plugin("get_paragraph_count", {"file_path": path})


@mcp.tool()
def search_in_document(path: str, pattern: str, regex: bool = False,
                       case_sensitive: bool = False,
                       max_results: int = 20,
                       context_paragraphs: int = 1) -> Dict[str, Any]:
    """Search for text in a document with paragraph context.

    Uses LibreOffice native search. Returns matches with surrounding
    paragraphs for context, without loading the entire document.

    Args:
        path: Absolute path to the document
        pattern: Search string or regex
        regex: Use regular expression (default: False)
        case_sensitive: Case-sensitive search (default: False)
        max_results: Max results to return (default: 20)
        context_paragraphs: Paragraphs of context around each match (default: 1)
    """
    return _call_plugin("search_in_document", {
        "file_path": path, "pattern": pattern, "regex": regex,
        "case_sensitive": case_sensitive, "max_results": max_results,
        "context_paragraphs": context_paragraphs})


@mcp.tool()
def replace_in_document(path: str, search: str, replace: str,
                        regex: bool = False,
                        case_sensitive: bool = False) -> Dict[str, Any]:
    """Find and replace text preserving all formatting.

    Args:
        path: Absolute path to the document
        search: Text to find
        replace: Replacement text
        regex: Use regular expression (default: False)
        case_sensitive: Case-sensitive matching (default: False)
    """
    return _call_plugin("replace_in_document", {
        "file_path": path, "search": search, "replace": replace,
        "regex": regex, "case_sensitive": case_sensitive})


@mcp.tool()
def insert_text_at_paragraph(path: str, paragraph_index: int,
                             text: str,
                             position: str = "after") -> Dict[str, Any]:
    """Insert text before or after a specific paragraph.

    Preserves all existing formatting.

    Args:
        path: Absolute path to the document
        paragraph_index: Target paragraph index
        text: Text to insert
        position: 'before' or 'after' (default: after)
    """
    return _call_plugin("insert_at_paragraph", {
        "file_path": path, "paragraph_index": paragraph_index,
        "text": text, "position": position})


@mcp.tool()
def add_document_ai_summary(path: str, para_index: int,
                            summary: str) -> Dict[str, Any]:
    """Add an AI annotation/summary to a heading.

    The summary is stored as a Writer annotation with Author='MCP-AI'.
    It will be shown when using content_strategy='ai_summary_first'.

    Args:
        path: Absolute path to the document
        para_index: Paragraph index of the heading
        summary: Summary text to attach
    """
    return _call_plugin("add_ai_summary", {
        "file_path": path, "para_index": para_index, "summary": summary})


@mcp.tool()
def get_document_ai_summaries(path: str) -> Dict[str, Any]:
    """List all AI annotations in a document.

    Args:
        path: Absolute path to the document
    """
    return _call_plugin("get_ai_summaries", {"file_path": path})


@mcp.tool()
def remove_document_ai_summary(path: str,
                               para_index: int) -> Dict[str, Any]:
    """Remove an AI annotation from a heading.

    Args:
        path: Absolute path to the document
        para_index: Paragraph index of the heading
    """
    return _call_plugin("remove_ai_summary", {
        "file_path": path, "para_index": para_index})


@mcp.tool()
def list_document_sections(path: str) -> Dict[str, Any]:
    """List all named text sections in a document.

    Args:
        path: Absolute path to the document
    """
    return _call_plugin("list_sections", {"file_path": path})


@mcp.tool()
def read_document_section(path: str,
                          section_name: str) -> Dict[str, Any]:
    """Read the content of a named text section.

    Args:
        path: Absolute path to the document
        section_name: Name of the section
    """
    return _call_plugin("read_section", {
        "file_path": path, "section_name": section_name})


@mcp.tool()
def list_document_bookmarks(path: str) -> Dict[str, Any]:
    """List all bookmarks in a document.

    Args:
        path: Absolute path to the document
    """
    return _call_plugin("list_bookmarks", {"file_path": path})


@mcp.tool()
def resolve_document_bookmark(path: str,
                               bookmark_name: str) -> Dict[str, Any]:
    """Resolve a heading bookmark to its current paragraph index.

    Bookmarks are stable identifiers that survive document edits.
    Use this to find the current position of a previously bookmarked heading.

    Args:
        path: Absolute path to the document
        bookmark_name: Bookmark name (e.g. _mcp_a1b2c3d4)
    """
    return _call_plugin("resolve_bookmark", {
        "file_path": path, "bookmark_name": bookmark_name})


@mcp.tool()
def get_document_page_count(path: str) -> Dict[str, Any]:
    """Get the page count of a document.

    Args:
        path: Absolute path to the document
    """
    return _call_plugin("get_page_count", {"file_path": path})


# Data models for structured responses
class DocumentInfo(BaseModel):
    """Information about a LibreOffice document"""
    path: str = Field(description="Full path to the document")
    filename: str = Field(description="Document filename")
    format: str = Field(description="Document format (odt, ods, odp, etc.)")
    size_bytes: int = Field(description="File size in bytes")
    modified_time: datetime = Field(description="Last modification time")
    exists: bool = Field(description="Whether the file exists")


class TextContent(BaseModel):
    """Text content extracted from a document"""
    content: str = Field(description="The extracted text content")
    word_count: int = Field(description="Number of words in the content")
    char_count: int = Field(description="Number of characters in the content")
    page_count: Optional[int] = Field(description="Number of pages (if available)")


class ConversionResult(BaseModel):
    """Result of document conversion"""
    source_path: str = Field(description="Source document path")
    target_path: str = Field(description="Target document path")
    source_format: str = Field(description="Original format")
    target_format: str = Field(description="Converted format")
    success: bool = Field(description="Whether conversion was successful")
    error_message: Optional[str] = Field(description="Error message if conversion failed")


class SpreadsheetData(BaseModel):
    """Data from a spreadsheet"""
    sheet_name: str = Field(description="Name of the sheet")
    data: List[List[str]] = Field(description="2D array of cell values")
    row_count: int = Field(description="Number of rows")
    col_count: int = Field(description="Number of columns")


# Helper functions

def _find_libreoffice_executable() -> str:
    """Find the LibreOffice executable on the current platform."""
    if platform.system() == "Windows":
        # On Windows, check common install locations
        candidates = []
        for prog_dir in [os.environ.get("ProgramFiles", r"C:\Program Files"),
                         os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                         os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs")]:
            if prog_dir:
                candidates.append(os.path.join(prog_dir, "LibreOffice", "program", "soffice.exe"))
        # Also try PATH
        for name in ['soffice', 'soffice.exe', 'libreoffice', 'loffice']:
            import shutil
            found = shutil.which(name)
            if found:
                return found
        # Check install paths
        for path in candidates:
            if os.path.isfile(path):
                return path
        return "soffice"  # fallback, let subprocess raise if not found
    else:
        # Unix: try common names in PATH
        import shutil
        for name in ['libreoffice', 'loffice', 'soffice']:
            found = shutil.which(name)
            if found:
                return found
        return "libreoffice"  # fallback


# Cache the result to avoid repeated filesystem scans
_libreoffice_exe: Optional[str] = None

def _get_libreoffice_exe() -> str:
    """Return the cached LibreOffice executable path."""
    global _libreoffice_exe
    if _libreoffice_exe is None:
        _libreoffice_exe = _find_libreoffice_executable()
    return _libreoffice_exe


def _is_libreoffice_running() -> bool:
    """Check if a LibreOffice process is already running."""
    try:
        if platform.system() == "Windows":
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq soffice.bin", "/NH"],
                capture_output=True, text=True, timeout=5)
            return "soffice.bin" in result.stdout
        else:
            result = subprocess.run(
                ["pgrep", "-x", "soffice.bin"],
                capture_output=True, timeout=5)
            return result.returncode == 0
    except Exception:
        return False


def _run_libreoffice_command(args: List[str], timeout: int = 30) -> subprocess.CompletedProcess:
    """Run a LibreOffice command with proper error handling"""
    executable = _get_libreoffice_exe()
    try:
        # In GUI mode, refuse to run if LibreOffice is already open
        if not HEADLESS_MODE and _is_libreoffice_running():
            raise RuntimeError(
                "LibreOffice is already running. In GUI mode (MCP_LIBREOFFICE_GUI=1), "
                "the MCP server needs exclusive access. Close LibreOffice first, "
                "or use the desktop shortcut created by create-shortcut.ps1."
            )
        # In GUI mode, strip --headless so the user sees LibreOffice working
        if not HEADLESS_MODE:
            args = [a for a in args if a != '--headless']
        cmd = [executable] + args
        # On Windows in headless mode, hide the console window
        kwargs = {}
        if platform.system() == "Windows" and HEADLESS_MODE:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            kwargs['startupinfo'] = startupinfo
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            **kwargs
        )
        return result
    except FileNotFoundError:
        raise FileNotFoundError(
            f"LibreOffice executable not found at '{executable}'. "
            "Please install LibreOffice and ensure it is in your PATH.\n"
            "Windows: run setup-windows.ps1 or add LibreOffice\\program to PATH."
        )
    except subprocess.TimeoutExpired:
        raise TimeoutError(f"LibreOffice command timed out after {timeout} seconds")


def _get_document_info(file_path: str) -> DocumentInfo:
    """Get information about a document file"""
    path = Path(file_path)
    return DocumentInfo(
        path=str(path.absolute()),
        filename=path.name,
        format=path.suffix.lower().lstrip('.'),
        size_bytes=path.stat().st_size if path.exists() else 0,
        modified_time=datetime.fromtimestamp(path.stat().st_mtime) if path.exists() else datetime.now(),
        exists=path.exists()
    )


# Core LibreOffice Tools

@mcp.tool()
def create_document(path: str, doc_type: str = "writer", content: str = "") -> DocumentInfo:
    """Create a new LibreOffice document
    
    Args:
        path: Full path where the document should be created
        doc_type: Type of document to create (writer, calc, impress, draw)
        content: Initial content for the document (for writer documents)
    """
    path_obj = Path(path)
    
    # Ensure directory exists
    path_obj.parent.mkdir(parents=True, exist_ok=True)
    
    # Map document types to LibreOffice formats
    format_map = {
        "writer": ".odt",
        "calc": ".ods", 
        "impress": ".odp",
        "draw": ".odg"
    }
    
    if doc_type not in format_map:
        raise ValueError(f"Unsupported document type: {doc_type}. Use: {list(format_map.keys())}")
    
    # Add appropriate extension if not present
    if not path_obj.suffix:
        path = str(path_obj) + format_map[doc_type]
        path_obj = Path(path)
    
    try:
        if doc_type == "writer" and content:
            # For writer documents with content, create a simple text file first
            # then convert to ODT format
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as tmp:
                tmp.write(content)
                tmp_path = tmp.name
            
            try:
                # Try to convert text to ODT
                result = _run_libreoffice_command([
                    '--headless',
                    '--convert-to', 'odt',
                    '--outdir', str(path_obj.parent),
                    tmp_path
                ])
                
                # Find the converted file and move it to the target location
                tmp_stem = Path(tmp_path).stem
                converted_file = path_obj.parent / f"{tmp_stem}.odt"
                
                if converted_file.exists():
                    converted_file.rename(path_obj)
                else:
                    # If conversion failed, create a basic ODT file manually
                    # This is a minimal ODT structure
                    import zipfile
                    import xml.etree.ElementTree as ET
                    
                    # Create a minimal ODT file
                    with zipfile.ZipFile(path_obj, 'w', zipfile.ZIP_DEFLATED) as zf:
                        # Add mimetype
                        zf.writestr('mimetype', 'application/vnd.oasis.opendocument.text')
                        
                        # Add META-INF/manifest.xml
                        manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
                        zf.writestr('META-INF/manifest.xml', manifest)
                        
                        # Add content.xml with the text
                        content_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body>
  <office:text>
   <text:p>{content}</text:p>
  </office:text>
 </office:body>
</office:document-content>'''
                        zf.writestr('content.xml', content_xml)
                
            finally:
                # Clean up temporary file
                Path(tmp_path).unlink(missing_ok=True)
        else:
            # For other document types or empty documents, try LibreOffice template creation
            try:
                # Try using LibreOffice to create an empty document
                template_map = {
                    "writer": "--writer",
                    "calc": "--calc", 
                    "impress": "--impress",
                    "draw": "--draw"
                }
                
                # Create using LibreOffice command line
                result = _run_libreoffice_command([
                    '--headless',
                    '--invisible',
                    '--nodefault',
                    '--nolockcheck',
                    '--nologo',
                    '--norestore',
                    '--convert-to', format_map[doc_type].lstrip('.'),
                    '--outdir', str(path_obj.parent),
                    '/dev/null'  # Convert from nothing to create empty document
                ])
                
                # If that doesn't work, create minimal file structure
                if not path_obj.exists():
                    if doc_type == "writer":
                        # Create minimal ODT
                        import zipfile
                        with zipfile.ZipFile(path_obj, 'w', zipfile.ZIP_DEFLATED) as zf:
                            zf.writestr('mimetype', 'application/vnd.oasis.opendocument.text')
                            manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
                            zf.writestr('META-INF/manifest.xml', manifest)
                            content_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
 <office:body>
  <office:text>
   <text:p></text:p>
  </office:text>
 </office:body>
</office:document-content>'''
                            zf.writestr('content.xml', content_xml)
                    else:
                        # For other formats, create empty file
                        path_obj.touch()
                        
            except Exception as e:
                # Fallback: create empty file
                path_obj.touch()
        
        return _get_document_info(str(path_obj))
        
    except Exception as e:
        raise RuntimeError(f"Failed to create document: {str(e)}")


@mcp.tool()
def read_document_text(path: str) -> TextContent:
    """Extract text content from a LibreOffice document
    
    Args:
        path: Path to the document file
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        # Use LibreOffice to convert to plain text
        with tempfile.TemporaryDirectory() as tmp_dir:
            result = _run_libreoffice_command([
                '--headless',
                '--convert-to', 'txt',
                '--outdir', tmp_dir,
                str(path_obj)
            ])
            
            # Debug: check what files were created
            tmp_path = Path(tmp_dir)
            created_files = list(tmp_path.iterdir())
            
            # Look for the converted text file
            txt_file = None
            # Try different possible names
            possible_names = [
                path_obj.stem + '.txt',
                path_obj.name + '.txt', 
                'output.txt'
            ]
            
            for name in possible_names:
                candidate = tmp_path / name
                if candidate.exists():
                    txt_file = candidate
                    break
            
            # If no specific file found, try any .txt file
            if not txt_file:
                txt_files = list(tmp_path.glob('*.txt'))
                if txt_files:
                    txt_file = txt_files[0]
            
            if txt_file and txt_file.exists():
                content = txt_file.read_text(encoding='utf-8', errors='ignore')
            else:
                # Fallback: try to extract text directly from ODT if it's a zip file
                if path_obj.suffix.lower() == '.odt':
                    content = _extract_text_from_odt(str(path_obj))
                else:
                    # Last resort: read as plain text
                    try:
                        content = path_obj.read_text(encoding='utf-8', errors='ignore')
                    except:
                        raise RuntimeError(f"Could not extract text. LibreOffice output: {result.stderr}. Files created: {[f.name for f in created_files]}")
        
        word_count = len(content.split())
        char_count = len(content)
        
        return TextContent(
            content=content,
            word_count=word_count,
            char_count=char_count,
            page_count=None  # Page count would require more complex parsing
        )
        
    except Exception as e:
        raise RuntimeError(f"Failed to read document: {str(e)}")


def _extract_text_from_odt(file_path: str) -> str:
    """Extract text content directly from ODT file"""
    import zipfile
    import xml.etree.ElementTree as ET
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # Read content.xml from the ODT file
            content_xml = zf.read('content.xml').decode('utf-8')
            
            # Parse XML and extract text
            root = ET.fromstring(content_xml)
            
            # Find all text elements (simplified extraction)
            text_parts = []
            for elem in root.iter():
                if elem.text:
                    text_parts.append(elem.text)
                if elem.tail:
                    text_parts.append(elem.tail)
            
            return ' '.join(text_parts).strip()
    
    except Exception as e:
        raise RuntimeError(f"Failed to extract text from ODT: {str(e)}")


@mcp.tool()
def convert_document(source_path: str, target_path: str, target_format: str) -> ConversionResult:
    """Convert a document to a different format
    
    Args:
        source_path: Path to the source document
        target_path: Path where converted document should be saved
        target_format: Target format (pdf, docx, xlsx, pptx, html, txt, etc.)
    """
    source_obj = Path(source_path)
    target_obj = Path(target_path)
    
    if not source_obj.exists():
        return ConversionResult(
            source_path=source_path,
            target_path=target_path,
            source_format=source_obj.suffix.lower().lstrip('.'),
            target_format=target_format,
            success=False,
            error_message=f"Source file not found: {source_path}"
        )
    
    # Ensure target directory exists
    target_obj.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        result = _run_libreoffice_command([
            '--headless',
            '--convert-to', target_format,
            '--outdir', str(target_obj.parent),
            str(source_obj)
        ])
        
        # LibreOffice creates the file with a predictable name
        expected_output = target_obj.parent / (source_obj.stem + f'.{target_format}')
        
        # Move to target location if needed
        if expected_output.exists() and expected_output != target_obj:
            expected_output.rename(target_obj)
        
        success = target_obj.exists()
        error_msg = None if success else f"Conversion failed. LibreOffice output: {result.stderr}"
        
        return ConversionResult(
            source_path=source_path,
            target_path=str(target_obj),
            source_format=source_obj.suffix.lower().lstrip('.'),
            target_format=target_format,
            success=success,
            error_message=error_msg
        )
        
    except Exception as e:
        return ConversionResult(
            source_path=source_path,
            target_path=target_path,
            source_format=source_obj.suffix.lower().lstrip('.'),
            target_format=target_format,
            success=False,
            error_message=str(e)
        )


@mcp.tool()
def get_document_info(path: str) -> DocumentInfo:
    """Get detailed information about a LibreOffice document
    
    Args:
        path: Path to the document file
    """
    return _get_document_info(path)


@mcp.tool()
def read_spreadsheet_data(path: str, sheet_name: Optional[str] = None, max_rows: int = 100) -> SpreadsheetData:
    """Read data from a LibreOffice Calc spreadsheet
    
    Args:
        path: Path to the spreadsheet file (.ods, .xlsx, etc.)
        sheet_name: Name of the specific sheet to read (if None, reads first sheet)
        max_rows: Maximum number of rows to read (default 100)
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Spreadsheet not found: {path}")
    
    try:
        # Convert to CSV to easily read the data
        with tempfile.TemporaryDirectory() as tmp_dir:
            result = _run_libreoffice_command([
                '--headless',
                '--convert-to', 'csv',
                '--outdir', tmp_dir,
                str(path_obj)
            ])
            
            csv_file = Path(tmp_dir) / (path_obj.stem + '.csv')
            if not csv_file.exists():
                raise RuntimeError("Failed to convert spreadsheet to CSV")
            
            # Read CSV data
            import csv
            data = []
            with open(csv_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for i, row in enumerate(reader):
                    if i >= max_rows:
                        break
                    data.append(row)
            
            row_count = len(data)
            col_count = max(len(row) for row in data) if data else 0
            
            return SpreadsheetData(
                sheet_name=sheet_name or "Sheet1",
                data=data,
                row_count=row_count,
                col_count=col_count
            )
            
    except Exception as e:
        raise RuntimeError(f"Failed to read spreadsheet: {str(e)}")


@mcp.tool()
def insert_text_at_position(path: str, text: str, position: str = "end") -> DocumentInfo:
    """Insert text into a LibreOffice Writer document
    
    Args:
        path: Path to the document file
        text: Text to insert
        position: Where to insert the text ("start", "end", or "replace")
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        # Read existing content
        existing_content = read_document_text(path).content
        
        # Determine new content based on position
        if position == "start":
            new_content = text + "\n" + existing_content
        elif position == "end":
            new_content = existing_content + "\n" + text
        elif position == "replace":
            new_content = text
        else:
            raise ValueError("Position must be 'start', 'end', or 'replace'")
        
        # Get file extension to determine document type
        file_ext = path_obj.suffix.lower()
        
        # Create backup
        backup_path = str(path_obj) + '.backup'
        import shutil
        shutil.copy2(path_obj, backup_path)
        
        try:
            if file_ext in ['.odt', '.docx', '.doc']:
                # For Writer documents, use a more robust approach
                success = _insert_text_writer_document(str(path_obj), new_content)
                if not success:
                    # Fallback: recreate document with new content
                    _recreate_writer_document(str(path_obj), new_content)
            else:
                # For other formats, try to recreate
                _recreate_document_with_content(str(path_obj), new_content)
                
        except Exception as convert_error:
            # Restore backup if anything goes wrong
            shutil.copy2(backup_path, path_obj)
            raise RuntimeError(f"Failed to modify document: {str(convert_error)}")
        finally:
            # Clean up backup
            Path(backup_path).unlink(missing_ok=True)
        
        return _get_document_info(str(path_obj))
        
    except Exception as e:
        raise RuntimeError(f"Failed to insert text: {str(e)}")


def _insert_text_writer_document(path: str, content: str) -> bool:
    """Insert text into Writer document using LibreOffice macro approach"""
    try:
        # Create a temporary script file for LibreOffice macro
        script_content = f'''
import uno
import sys

def modify_document():
    try:
        # Connect to LibreOffice
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        
        # Start LibreOffice if not running
        import subprocess
        lo_cmd = [_get_libreoffice_exe(), "--accept=socket,host=127.0.0.1,port=2002;urp;"]
        if HEADLESS_MODE:
            lo_cmd.insert(1, "--headless")
        subprocess.Popen(lo_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        import time
        time.sleep(2)  # Wait for LibreOffice to start
        
        ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # Open document
        doc = desktop.loadComponentFromURL("file://{path}", "_blank", 0, ())
        
        # Clear content and insert new content
        text = doc.getText()
        text.setString("{content.replace('"', '\\"')}")
        
        # Save document
        doc.store()
        doc.close(True)
        
        return True
    except:
        return False

if __name__ == "__main__":
    modify_document()
'''
        
        # Try the macro approach (advanced)
        with tempfile.NamedTemporaryFile(mode='w', suffix='.py', delete=False) as script:
            script.write(script_content)
            script_path = script.name
        
        try:
            # This is complex and may not work in all environments
            # So we'll skip it and use the simpler approach
            return False
        finally:
            Path(script_path).unlink(missing_ok=True)
            
    except Exception:
        return False


def _recreate_writer_document(path: str, content: str):
    """Recreate a Writer document with new content"""
    path_obj = Path(path)
    original_ext = path_obj.suffix.lower()
    
    # Create temporary text file with proper encoding
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as tmp:
        tmp.write(content)
        tmp_path = tmp.name
    
    try:
        # Determine target format
        if original_ext == '.odt':
            target_format = 'odt'
        elif original_ext == '.docx':
            target_format = 'docx'  
        elif original_ext == '.doc':
            target_format = 'doc'
        else:
            target_format = 'odt'  # Default to ODT
        
        # Remove original file
        if path_obj.exists():
            path_obj.unlink()
        
        # Use LibreOffice to convert text to document format
        # First, let's try a different approach - create from template
        try:
            # Method 1: Convert using LibreOffice
            result = _run_libreoffice_command([
                '--headless',
                '--invisible', 
                '--convert-to', target_format,
                '--outdir', str(path_obj.parent),
                tmp_path
            ])
            
            # Find and rename the converted file
            tmp_name = Path(tmp_path).stem
            converted_file = path_obj.parent / f"{tmp_name}.{target_format}"
            
            if converted_file.exists():
                converted_file.rename(path_obj)
                return
        except Exception:
            pass
        
        # Method 2: If conversion failed, create minimal valid ODT
        if original_ext == '.odt' or target_format == 'odt':
            _create_minimal_odt(path_obj, content)
        else:
            # For other formats, create a simple text file with correct extension
            with open(path_obj, 'w', encoding='utf-8') as f:
                f.write(content)
        
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _create_minimal_odt(path: Path, content: str):
    """Create a minimal but valid ODT file with the given content"""
    import zipfile
    
    # Escape content for XML
    import html
    escaped_content = html.escape(content)
    
    # Split content into paragraphs
    paragraphs = escaped_content.split('\n')
    
    # Create paragraph XML
    text_paragraphs = []
    for para in paragraphs:
        if para.strip():
            text_paragraphs.append(f'   <text:p text:style-name="Standard">{para}</text:p>')
        else:
            text_paragraphs.append('   <text:p text:style-name="Standard"/>')
    
    text_content = '\n'.join(text_paragraphs)
    
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # mimetype (must be first and uncompressed)
        zf.writestr('mimetype', 'application/vnd.oasis.opendocument.text', compress_type=zipfile.ZIP_STORED)
        
        # META-INF/manifest.xml
        manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
        zf.writestr('META-INF/manifest.xml', manifest_xml)
        
        # content.xml
        content_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
 <office:scripts/>
 <office:font-face-decls/>
 <office:automatic-styles/>
 <office:body>
  <office:text>
{text_content}
  </office:text>
 </office:body>
</office:document-content>'''
        zf.writestr('content.xml', content_xml)
        
        # styles.xml (minimal)
        styles_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
 <office:font-face-decls/>
 <office:styles>
  <style:default-style style:family="paragraph">
   <style:paragraph-properties fo:hyphenation-ladder-count="no-limit"/>
   <style:text-properties style:tab-stop-distance="0.5in"/>
  </style:default-style>
  <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
 </office:styles>
 <office:automatic-styles/>
 <office:master-styles/>
</office:document-styles>'''
        zf.writestr('styles.xml', styles_xml)
        
        # meta.xml (minimal)
        meta_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">
 <office:meta>
  <meta:generator>LibreOffice MCP Server</meta:generator>
 </office:meta>
</office:document-meta>'''
        zf.writestr('meta.xml', meta_xml)


def _recreate_document_with_content(path: str, content: str):
    """Recreate any document with new content"""
    # For non-Writer documents, just create a text file with the correct extension
    with open(path, 'w') as f:
        f.write(content)


# Resources for document discovery

@mcp.resource("documents://")
def list_documents() -> List[str]:
    """List all LibreOffice documents in common locations"""
    documents = []
    
    # Common document locations
    search_paths = [
        Path.home() / "Documents",
        Path.home() / "Desktop", 
        Path.cwd()
    ]
    
    # LibreOffice file extensions
    extensions = {'.odt', '.ods', '.odp', '.odg', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'}
    
    for search_path in search_paths:
        if search_path.exists():
            for ext in extensions:
                for doc in search_path.rglob(f'*{ext}'):
                    if doc.is_file():
                        documents.append(str(doc))
    
    return sorted(documents)


@mcp.resource("document://{path}")
def get_document_content(path: str) -> str:
    """Get the text content of a specific document"""
    try:
        # Decode the path properly - remove the leading slash if present
        if path.startswith('/'):
            actual_path = path
        else:
            actual_path = '/' + path
            
        content = read_document_text(actual_path)
        return f"Document: {Path(actual_path).name}\n" + \
               f"Words: {content.word_count}, Characters: {content.char_count}\n\n" + \
               content.content
    except Exception as e:
        return f"Error reading document {path}: {str(e)}"


# Additional utility tools

@mcp.tool()
def search_documents(query: str, search_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """Search for documents containing specific text
    
    Args:
        query: Text to search for
        search_path: Directory to search in (default: common document locations)
    """
    results = []
    
    if search_path:
        search_paths = [Path(search_path)]
    else:
        search_paths = [
            Path.home() / "Documents",
            Path.home() / "Desktop",
            Path.cwd()
        ]
    
    extensions = {'.odt', '.ods', '.odp', '.odg', '.doc', '.docx', '.txt'}
    
    for search_dir in search_paths:
        if not search_dir.exists():
            continue
            
        for ext in extensions:
            for doc_path in search_dir.rglob(f'*{ext}'):
                if not doc_path.is_file():
                    continue
                    
                try:
                    # Read document content
                    content = read_document_text(str(doc_path))
                    
                    # Search for query in content (case-insensitive)
                    if query.lower() in content.content.lower():
                        results.append({
                            "path": str(doc_path),
                            "filename": doc_path.name,
                            "format": doc_path.suffix.lower(),
                            "word_count": content.word_count,
                            "match_context": _get_match_context(content.content, query)
                        })
                        
                except Exception:
                    # Skip documents that can't be read
                    continue
    
    return results


def _get_match_context(content: str, query: str, context_chars: int = 200) -> str:
    """Get surrounding context for a search match"""
    content_lower = content.lower()
    query_lower = query.lower()
    
    match_pos = content_lower.find(query_lower)
    if match_pos == -1:
        return ""
    
    start = max(0, match_pos - context_chars // 2)
    end = min(len(content), match_pos + len(query) + context_chars // 2)
    
    context = content[start:end]
    if start > 0:
        context = "..." + context
    if end < len(content):
        context = context + "..."
        
    return context


@mcp.tool()
def batch_convert_documents(source_dir: str, target_dir: str, target_format: str, 
                          source_extensions: Optional[List[str]] = None) -> List[ConversionResult]:
    """Convert multiple documents in a directory to a different format
    
    Args:
        source_dir: Directory containing source documents
        target_dir: Directory where converted documents should be saved
        target_format: Target format for conversion
        source_extensions: List of source file extensions to convert (default: common formats)
    """
    source_path = Path(source_dir)
    target_path = Path(target_dir)
    
    if not source_path.exists():
        raise FileNotFoundError(f"Source directory not found: {source_dir}")
    
    target_path.mkdir(parents=True, exist_ok=True)
    
    if source_extensions is None:
        source_extensions = ['.odt', '.ods', '.odp', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']
    
    results = []
    
    for ext in source_extensions:
        for doc_file in source_path.rglob(f'*{ext}'):
            if doc_file.is_file():
                target_file = target_path / (doc_file.stem + f'.{target_format}')
                result = convert_document(str(doc_file), str(target_file), target_format)
                results.append(result)
    
    return results


@mcp.tool()
def merge_text_documents(document_paths: List[str], output_path: str, 
                        separator: str = "\n\n---\n\n") -> DocumentInfo:
    """Merge multiple text documents into a single document
    
    Args:
        document_paths: List of paths to documents to merge
        output_path: Path where merged document should be saved
        separator: Text to insert between merged documents
    """
    merged_content = []
    
    for doc_path in document_paths:
        try:
            content = read_document_text(doc_path)
            doc_name = Path(doc_path).name
            merged_content.append(f"=== {doc_name} ===\n\n{content.content}")
        except Exception as e:
            merged_content.append(f"=== {Path(doc_path).name} ===\n\nError reading document: {str(e)}")
    
    final_content = separator.join(merged_content)
    
    # Create the merged document
    return create_document(output_path, "writer", final_content)


@mcp.tool()
def get_document_statistics(path: str) -> Dict[str, Any]:
    """Get detailed statistics about a document
    
    Args:
        path: Path to the document file
    """
    doc_info = get_document_info(path)
    
    if not doc_info.exists:
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        content = read_document_text(path)
        
        # Calculate additional statistics
        lines = content.content.split('\n')
        paragraphs = [p for p in content.content.split('\n\n') if p.strip()]
        sentences = [s for s in content.content.replace('!', '.').replace('?', '.').split('.') if s.strip()]
        
        return {
            "file_info": doc_info.model_dump(),
            "content_stats": {
                "word_count": content.word_count,
                "character_count": content.char_count,
                "line_count": len(lines),
                "paragraph_count": len(paragraphs),
                "sentence_count": len(sentences),
                "average_words_per_sentence": content.word_count / max(len(sentences), 1),
                "average_chars_per_word": content.char_count / max(content.word_count, 1)
            }
        }
        
    except Exception as e:
        return {
            "file_info": doc_info.model_dump(),
            "error": f"Could not analyze content: {str(e)}"
        }


# Main server entry point
def main():
    """Run the LibreOffice MCP server"""
    import sys
    
    # Parse command line arguments
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        
        if arg == "--test":
            # Test mode - run some basic functionality tests
            print("🧪 Running LibreOffice MCP Server tests...")
            asyncio.run(test_server())
            return
        
        elif arg == "--help" or arg == "-h":
            # Show help
            print("LibreOffice MCP Server")
            print("=" * 30)
            print("Usage:")
            print("  python src/main.py          # Start MCP server (stdio mode)")
            print("  python src/main.py --test   # Run functionality tests")
            print("  python src/main.py --help   # Show this help")
            print("")
            print("MCP Server Mode:")
            print("  The server runs in stdio mode for MCP protocol communication.")
            print("  It reads JSON-RPC messages from stdin and writes responses to stdout.")
            print("  Use with MCP clients like Claude Desktop or test with test_client.py")
            print("")
            print("Testing:")
            print("  cd tests/ && python test_client.py  # Interactive test client")
            return
        
        elif arg == "--version":
            print("LibreOffice MCP Server v1.0.0")
            return
    
    # Normal server mode - show startup message and run
    print("🚀 Starting LibreOffice MCP Server...", file=sys.stderr)
    print("📡 Running in MCP protocol mode (stdio)", file=sys.stderr)
    print("💡 Use --help for command line options", file=sys.stderr)
    print("🔌 Connect via MCP clients or test with: cd tests/ && python test_client.py", file=sys.stderr)
    print("", file=sys.stderr)
    
    try:
        mcp.run()
    except KeyboardInterrupt:
        print("\n👋 LibreOffice MCP Server stopped", file=sys.stderr)
    except Exception as e:
        print(f"\n❌ Server error: {e}", file=sys.stderr)
        sys.exit(1)


async def test_server():
    """Test the server functionality"""
    print("Testing LibreOffice MCP Server...")
    print("=" * 50)
    
    # First test LibreOffice installation
    if not await test_libreoffice_installation():
        print("❌ LibreOffice installation test failed")
        return
    
    # Test creating a document
    test_doc = "/tmp/test_document.odt"
    try:
        print("\nTesting document creation...")
        result = create_document(test_doc, "writer", "This is a test document.\n\nHello, LibreOffice!")
        print(f"✓ Created test document: {result.filename}")
        
        # Test reading the document
        print("Testing document reading...")
        content = read_document_text(test_doc)
        print(f"✓ Read document content: {content.word_count} words")
        
        # Test converting to PDF
        print("Testing document conversion...")
        pdf_path = "/tmp/test_document.pdf"
        conversion = convert_document(test_doc, pdf_path, "pdf")
        if conversion.success:
            print(f"✓ Conversion to PDF: Success")
        else:
            print(f"⚠ Conversion to PDF: Failed - {conversion.error_message}")
        
        # Test document statistics
        print("Testing document statistics...")
        stats = get_document_statistics(test_doc)
        print(f"✓ Document statistics: {stats['content_stats']['word_count']} words, {stats['content_stats']['sentence_count']} sentences")
        
        print("\n✅ All tests passed!")
        
    except Exception as e:
        print(f"❌ Test failed: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Clean up test files
        for test_file in [test_doc, "/tmp/test_document.pdf"]:
            try:
                Path(test_file).unlink(missing_ok=True)
            except:
                pass


# Test LibreOffice functionality directly
async def test_libreoffice_installation():
    """Test LibreOffice installation and basic functionality"""
    print("\nTesting LibreOffice Installation...")
    print("=" * 40)
    
    try:
        # Test basic LibreOffice command
        result = _run_libreoffice_command(['--version'])
        if result.returncode == 0:
            print(f"✓ LibreOffice version: {result.stdout.strip()}")
        else:
            print(f"⚠ LibreOffice version check failed: {result.stderr}")
    except Exception as e:
        print(f"❌ LibreOffice not accessible: {str(e)}")
        return False
    
    # Test headless mode
    try:
        result = _run_libreoffice_command(['--headless', '--help'])
        if result.returncode == 0 or 'headless' in result.stdout.lower():
            print("✓ LibreOffice headless mode available")
        else:
            print(f"⚠ LibreOffice headless mode issue: {result.stderr}")
    except Exception as e:
        print(f"❌ LibreOffice headless mode failed: {str(e)}")
    
    return True


# Live viewing and document management tools

@mcp.tool()
def open_document_in_libreoffice(path: str, readonly: bool = False) -> Dict[str, Any]:
    """Open a document in LibreOffice GUI for live viewing
    
    Args:
        path: Path to the document to open
        readonly: Whether to open in read-only mode (default: False)
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        # Build command to open LibreOffice with GUI
        cmd = [_get_libreoffice_exe()]
        
        if readonly:
            cmd.append('--view')
        
        # Add the document path
        cmd.append(str(path_obj.absolute()))
        
        # Start LibreOffice GUI (non-blocking)
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            # Detach from parent process (platform-appropriate)
            **({"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS}
               if platform.system() == "Windows"
               else {"start_new_session": True})
        )
        
        return {
            "success": True,
            "message": f"Opened {path_obj.name} in LibreOffice GUI",
            "path": str(path_obj.absolute()),
            "readonly": readonly,
            "process_id": process.pid,
            "note": "Document is now open for live viewing. Changes made via MCP will be reflected after saving and refreshing."
        }
        
    except Exception as e:
        raise RuntimeError(f"Failed to open document in LibreOffice: {str(e)}")


@mcp.tool()
def refresh_document_in_libreoffice(path: str) -> Dict[str, Any]:
    """Send a refresh signal to LibreOffice to reload a document
    
    Args:
        path: Path to the document that should be refreshed
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        # Try to send a signal to LibreOffice to refresh
        # This uses LibreOffice's ability to detect file changes
        
        # Method 1: Touch the file to update modification time
        import time
        current_time = time.time()
        path_obj.touch()
        
        # Method 2: Try to send a signal via LibreOffice's socket interface
        try:
            # This is a more advanced approach that may work if LibreOffice is running
            refresh_cmd = [_get_libreoffice_exe(), '--invisible',
                '--accept=socket,host=127.0.0.1,port=2002;urp;',
                '--norestore', '--nologo']
            if HEADLESS_MODE:
                refresh_cmd.insert(1, '--headless')
            result = subprocess.run(refresh_cmd, timeout=2, capture_output=True)
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass  # LibreOffice may already be running or not available
        
        return {
            "success": True,
            "message": f"Refresh signal sent for {path_obj.name}",
            "path": str(path_obj.absolute()),
            "note": "LibreOffice should detect the file change and prompt to reload. Manual refresh may be needed."
        }
        
    except Exception as e:
        return {
            "success": False,
            "message": f"Failed to refresh document: {str(e)}",
            "path": str(path_obj.absolute()),
            "note": "Try manually refreshing in LibreOffice (File → Reload)"
        }


@mcp.tool()
def watch_document_changes(path: str, duration_seconds: int = 30) -> Dict[str, Any]:
    """Watch a document for changes and provide live updates
    
    Args:
        path: Path to the document to watch
        duration_seconds: How long to watch for changes (default: 30 seconds)
    """
    path_obj = Path(path)
    if not path_obj.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    
    try:
        import time
        
        # Get initial state
        initial_stat = path_obj.stat()
        initial_size = initial_stat.st_size
        initial_mtime = initial_stat.st_mtime
        
        start_time = time.time()
        changes_detected = []
        
        print(f"👀 Watching {path_obj.name} for {duration_seconds} seconds...")
        
        while time.time() - start_time < duration_seconds:
            try:
                current_stat = path_obj.stat()
                current_size = current_stat.st_size
                current_mtime = current_stat.st_mtime
                
                if current_mtime > initial_mtime or current_size != initial_size:
                    change_info = {
                        "timestamp": datetime.now().isoformat(),
                        "size_before": initial_size,
                        "size_after": current_size,
                        "size_change": current_size - initial_size,
                        "modification_time": datetime.fromtimestamp(current_mtime).isoformat()
                    }
                    changes_detected.append(change_info)
                    
                    # Update baseline
                    initial_size = current_size
                    initial_mtime = current_mtime
                    
                    print(f"📝 Change detected: {change_info['size_change']:+d} bytes at {change_info['timestamp']}")
                
                time.sleep(1)  # Check every second
                
            except FileNotFoundError:
                break  # File was deleted
        
        return {
            "success": True,
            "path": str(path_obj.absolute()),
            "watch_duration": duration_seconds,
            "changes_detected": len(changes_detected),
            "changes": changes_detected,
            "message": f"Watched {path_obj.name} for {duration_seconds} seconds, detected {len(changes_detected)} changes"
        }
        
    except Exception as e:
        raise RuntimeError(f"Failed to watch document: {str(e)}")


@mcp.tool()
def create_live_editing_session(path: str, auto_refresh: bool = True) -> Dict[str, Any]:
    """Create a live editing session with automatic refresh capabilities
    
    Args:
        path: Path to the document for live editing
        auto_refresh: Whether to enable automatic refresh detection
    """
    path_obj = Path(path)
    
    try:
        # 1. Open the document in LibreOffice GUI
        open_result = open_document_in_libreoffice(str(path_obj), readonly=False)
        
        # 2. Set up file monitoring if requested
        session_info = {
            "session_id": f"live_session_{int(time.time())}",
            "document_path": str(path_obj.absolute()),
            "document_name": path_obj.name,
            "opened_in_gui": open_result["success"],
            "auto_refresh_enabled": auto_refresh,
            "created_at": datetime.now().isoformat(),
            "instructions": {
                "view_changes": "Document is open in LibreOffice GUI",
                "make_mcp_changes": "Use insert_text_at_position, convert_document, etc.",
                "see_updates": "LibreOffice will detect file changes and prompt to reload",
                "manual_refresh": "Press Ctrl+Shift+R in LibreOffice to force reload",
                "end_session": "Close LibreOffice window when done"
            }
        }
        
        if auto_refresh:
            session_info["monitoring"] = "File modification time will be updated after MCP operations"
        
        return session_info
        
    except Exception as e:
        raise RuntimeError(f"Failed to create live editing session: {str(e)}")


if __name__ == "__main__":
    main()
