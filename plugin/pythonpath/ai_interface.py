"""
LibreOffice MCP Extension - AI Interface Module

MCP Streamable HTTP server (POST /mcp) speaking JSON-RPC 2.0.
Handles backpressure: one tool execution at a time on the VCL main
thread, with a short wait timeout for queued requests and a longer
processing timeout for active execution.
"""

import json
import logging
import os
import queue
import socketserver
import threading
import uuid
from typing import Dict, Any, List, Optional
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

from .mcp_server import get_mcp_server
from .main_thread_executor import execute_on_main_thread


def _get_version() -> str:
    try:
        from .version import EXTENSION_VERSION
        return EXTENSION_VERSION
    except Exception:
        return "unknown"


_agent_instructions_cache: Optional[str] = None


def _load_agent_instructions() -> str:
    """Load AGENT.md from the extension directory (cached after first read)."""
    global _agent_instructions_cache
    if _agent_instructions_cache is not None:
        return _agent_instructions_cache
    ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for candidate in (
        os.path.join(ext_dir, "AGENT.md"),
        os.path.join(ext_dir, "..", "AGENT.md"),
    ):
        if os.path.isfile(candidate):
            with open(candidate, "r", encoding="utf-8") as f:
                _agent_instructions_cache = f.read()
            return _agent_instructions_cache
    _agent_instructions_cache = ""
    return _agent_instructions_cache


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# MCP protocol version we advertise
MCP_PROTOCOL_VERSION = "2025-11-25"


# ===================================================================
# Backpressure — one tool execution at a time
# ===================================================================

# Only one tool/call can be dispatched to the VCL main thread at a
# time.  Other requests wait up to _WAIT_TIMEOUT seconds; if the
# semaphore is still held they get a "busy" JSON-RPC error.
_tool_semaphore = threading.Semaphore(1)
_WAIT_TIMEOUT = 5.0       # seconds to wait in queue
_PROCESS_TIMEOUT = 60.0   # seconds for VCL main-thread execution


class BusyError(Exception):
    """The VCL main thread is already processing another tool call."""
    pass


# ===================================================================
# SSE transport (GET /sse + POST /sse) — ChatGPT compatibility
# ===================================================================
# Flow:
#   1. ChatGPT POSTs initialize → inline 200 (SSE not open yet)
#   2. ChatGPT opens GET /sse → SSE stream ready
#   3. Subsequent POSTs (tools/list etc) → 202, response via SSE
_sse_response_queue: Optional[queue.Queue] = None  # None = no stream
_sse_queue_lock = threading.Lock()

# ===================================================================
# Streamable HTTP session management
# ===================================================================
_mcp_session_id: Optional[str] = None


# ===================================================================
# Threaded HTTP server (one thread per request)
# ===================================================================

class _ThreadedHTTPServer(socketserver.ThreadingMixIn, HTTPServer):
    """HTTP server that handles each request in its own thread.

    This is essential for backpressure: multiple MCP clients can
    connect concurrently; the semaphore serialises actual tool
    execution while the HTTP connections stay alive.
    """
    daemon_threads = True


# ===================================================================
# MCP JSON-RPC helpers
# ===================================================================

def _jsonrpc_ok(req_id, result: Any) -> Dict[str, Any]:
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def _jsonrpc_error(req_id, code: int, message: str,
                   data: Any = None) -> Dict[str, Any]:
    err = {"code": code, "message": message}
    if data is not None:
        err["data"] = data
    return {"jsonrpc": "2.0", "id": req_id, "error": err}


# Standard JSON-RPC error codes
_PARSE_ERROR = -32700
_INVALID_REQUEST = -32600
_METHOD_NOT_FOUND = -32601
_INVALID_PARAMS = -32602
_INTERNAL_ERROR = -32603

# Implementation-defined codes
_SERVER_BUSY = -32000
_EXECUTION_TIMEOUT = -32001


def _build_mcp_tool_list(mcp_server) -> List[Dict[str, Any]]:
    """Convert our tool registry to MCP tools/list format."""
    tools = []
    for name, tool_obj in mcp_server.tools.items():
        schema = dict(getattr(tool_obj, "parameters", {}))
        schema.setdefault("type", "object")
        tools.append({
            "name": name,
            "description": getattr(tool_obj, "description", ""),
            "inputSchema": schema,
        })
    return tools


# ===================================================================
# HTTP request handler
# ===================================================================

class MCPRequestHandler(BaseHTTPRequestHandler):
    """HTTP handler for the MCP streamable-http protocol."""

    mcp_server = None          # injected by AIInterface.start()

    # ---------------------------------------------------------------
    # GET
    # ---------------------------------------------------------------
    def do_GET(self):
        try:
            path = urlparse(self.path).path
            if path == '/health':
                self._send_json(200, {
                    "status": "healthy",
                    "server": "LibreOffice MCP Extension",
                    "version": _get_version(),
                })
            elif path == '/':
                self._send_json(200, self._get_server_info())
            elif path == '/mcp':
                # MCP Streamable HTTP spec: GET opens an SSE stream for
                # server-to-client notifications.  We don't push events,
                # so we just keep the connection open until the client
                # disconnects.
                accept = self.headers.get("Accept", "")
                if "text/event-stream" not in accept:
                    self._send_json(406, {
                        "error": "Not Acceptable: must Accept text/event-stream"})
                    return
                self.send_response(200)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Cache-Control", "no-cache")
                self._send_cors_headers()
                self.end_headers()
                # Keep connection alive until client disconnects
                try:
                    while True:
                        # Send keepalive comment every 15s
                        self.wfile.write(b": keepalive\n\n")
                        self.wfile.flush()
                        import time
                        time.sleep(15)
                except (BrokenPipeError, ConnectionResetError, OSError):
                    pass  # client disconnected
            elif path == '/sse':
                self._handle_sse_stream()
            else:
                self._send_json(404, {"error": "Not found"})
        except Exception as e:
            logger.error("GET %s error: %s", self.path, e)
            self._send_json(500, {"error": str(e)})

    # ---------------------------------------------------------------
    # POST  (/mcp only)
    # ---------------------------------------------------------------
    def do_POST(self):
        try:
            path = urlparse(self.path).path
            if path == '/messages':
                body = self._read_body()
                if body is not None:
                    self._handle_sse_message(body)
                return
            if path == '/sse':
                body = self._read_body()
                if body is not None:
                    self._handle_sse_post(body)
                return
            if path == '/mcp':
                body = self._read_body()
                if body is not None:
                    self._handle_mcp(body)
                return
            self._send_json(404, {"error": "Not found"})
        except Exception as e:
            logger.error("POST %s error: %s", self.path, e)
            self._send_json(500, {"error": str(e)})

    # ---------------------------------------------------------------
    # DELETE  (MCP session teardown — accept it)
    # ---------------------------------------------------------------
    def do_DELETE(self):
        path = urlparse(self.path).path
        if path == '/mcp':
            self.send_response(200)
            self._send_cors_headers()
            self.end_headers()
        else:
            self._send_json(404, {"error": "Not found"})

    # ---------------------------------------------------------------
    # OPTIONS (CORS preflight)
    # ---------------------------------------------------------------
    def do_OPTIONS(self):
        self.send_response(204)
        self._send_cors_headers()
        self.end_headers()

    # ===============================================================
    # MCP protocol handler
    # ===============================================================

    def _handle_mcp(self, msg):
        """Route MCP JSON-RPC request(s) — single or batch (Claude Code)."""
        global _mcp_session_id

        method = msg.get("method", "?") if isinstance(msg, dict) else "batch"
        req_id = msg.get("id") if isinstance(msg, dict) else None
        logger.info("[MCP] <<< %s (id=%s)", method, req_id)

        is_initialize = (isinstance(msg, dict)
                         and msg.get("method") == "initialize")

        # Batch request (array of JSON-RPC messages)
        if isinstance(msg, list):
            responses = []
            for item in msg:
                result = self._process_jsonrpc(item)
                if result is not None:
                    _status, response = result
                    responses.append(response)
            if responses:
                self._send_json(200, responses)
            else:
                self.send_response(202)
                self._send_cors_headers()
                self.end_headers()
            return

        # Single request
        result = self._process_jsonrpc(msg)
        if result is None:
            self.send_response(202)
            self._send_cors_headers()
            if _mcp_session_id:
                self.send_header('Mcp-Session-Id', _mcp_session_id)
            self.end_headers()
            return
        status, response = result

        # On initialize: create session ID
        if is_initialize and status == 200:
            _mcp_session_id = str(uuid.uuid4())

        self.send_response(status)
        self._send_cors_headers()
        self.send_header('Content-Type', 'application/json')
        if _mcp_session_id:
            self.send_header('Mcp-Session-Id', _mcp_session_id)
        self.end_headers()
        body = json.dumps(response, ensure_ascii=False, default=str)
        logger.info("[MCP] >>> %s (id=%s) -> %d",
                    method, req_id, status)
        self.wfile.write(body.encode('utf-8'))

    # --- MCP method handlers ---

    def _mcp_initialize(self, params: Dict) -> Dict:
        # Negotiate protocol version: use client's version if we recognise it
        client_version = (params.get("protocolVersion") or MCP_PROTOCOL_VERSION)
        return {
            "protocolVersion": client_version,
            "capabilities": {
                "tools": {"listChanged": False},
                "resources": {"listChanged": False},
                "prompts": {"listChanged": False},
            },
            "serverInfo": {
                "name": "LibreOffice MCP",
                "version": _get_version(),
            },
            "instructions": (
                "LibreOffice MCP — AI document workspace. "
                "WORKFLOW: 1) list_open_documents 2) get_document_tree(depth=1) "
                "to get heading structure with bookmark IDs (_mcp_*). "
                "3) Use 'bookmark:_mcp_xxx' as locator in ALL tools — "
                "bookmarks are stable anchors that survive edits. "
                "NEVER read the whole document. Navigate by headings, "
                "drill down with get_heading_children, read with read_paragraphs. "
                "insert_image accepts URLs (auto-download)."
            ),
        }

    def _mcp_ping(self, params: Dict) -> Dict:
        return {}

    def _mcp_tools_list(self, params: Dict) -> Dict:
        return {"tools": _build_mcp_tool_list(self.mcp_server)}

    def _mcp_resources_list(self, params: Dict) -> Dict:
        return {"resources": []}

    def _mcp_prompts_list(self, params: Dict) -> Dict:
        return {"prompts": []}

    def _mcp_tools_call(self, params: Dict) -> Dict:
        tool_name = params.get("name")
        arguments = params.get("arguments", {})
        if not tool_name:
            raise ValueError("Missing 'name' in tools/call params")

        result = self._execute_with_backpressure(tool_name, arguments)

        is_error = (isinstance(result, dict)
                    and not result.get("success", True))
        return {
            "content": [
                {
                    "type": "text",
                    "text": json.dumps(result, ensure_ascii=False,
                                       default=str),
                }
            ],
            "isError": is_error,
        }

    # ===============================================================
    # JSON-RPC processing (shared by /mcp and /sse transports)
    # ===============================================================

    def _process_jsonrpc(self, msg):
        """Process a JSON-RPC message.

        Returns (http_status, response_dict) or None for notifications.
        """
        if not isinstance(msg, dict) or msg.get("jsonrpc") != "2.0":
            logger.warning("Invalid JSON-RPC: type=%s, body=%s",
                           type(msg).__name__, str(msg)[:300])
            return (400, _jsonrpc_error(
                None, _INVALID_REQUEST, "Invalid JSON-RPC 2.0 request"))

        method = msg.get("method", "")
        params = msg.get("params", {})
        req_id = msg.get("id")

        if req_id is None:
            return None  # notification

        handler = {
            "initialize":      self._mcp_initialize,
            "ping":            self._mcp_ping,
            "tools/list":      self._mcp_tools_list,
            "tools/call":      self._mcp_tools_call,
            "resources/list":  self._mcp_resources_list,
            "prompts/list":    self._mcp_prompts_list,
        }.get(method)

        if handler is None:
            return (400, _jsonrpc_error(
                req_id, _METHOD_NOT_FOUND,
                f"Unknown method: {method}"))

        try:
            result = handler(params)
            return (200, _jsonrpc_ok(req_id, result))
        except BusyError as e:
            logger.warning("MCP %s: busy (%s)", method, e)
            return (429, _jsonrpc_error(
                req_id, _SERVER_BUSY, str(e),
                {"retryable": True}))
        except TimeoutError as e:
            logger.error("MCP %s: timeout (%s)", method, e)
            return (504, _jsonrpc_error(
                req_id, _EXECUTION_TIMEOUT, str(e)))
        except Exception as e:
            logger.error("MCP %s error: %s", method, e, exc_info=True)
            return (500, _jsonrpc_error(
                req_id, _INTERNAL_ERROR, str(e)))

    # ===============================================================
    # /sse endpoint (ChatGPT compatibility)
    #
    # ChatGPT uses streamable HTTP via POST /sse — same as /mcp
    # but on the /sse path.  Direct JSON responses, no SSE stream.
    # GET /sse kept as notification stream (keepalive only).
    # ===============================================================

    def _handle_sse_stream(self):
        """GET /sse — notification stream (keepalive only)."""
        try:
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Connection", "keep-alive")
            self.send_header("X-Accel-Buffering", "no")
            self._send_cors_headers()
            self.end_headers()
            logger.info("[SSE] GET stream opened (notification keepalive)")
            while True:
                self.wfile.write(b": keepalive\n\n")
                self.wfile.flush()
                import time
                time.sleep(15)
        except (BrokenPipeError, ConnectionResetError, OSError):
            logger.info("[SSE] GET stream disconnected")

    def _handle_sse_post(self, msg):
        """POST /sse — streamable HTTP (same as /mcp, inline JSON)."""
        method = msg.get("method", "?") if isinstance(msg, dict) else "batch"
        req_id = msg.get("id") if isinstance(msg, dict) else None
        logger.info("[SSE] POST <<< %s (id=%s)", method, req_id)

        result = self._process_jsonrpc(msg)
        if result is None:
            self.send_response(202)
            self._send_cors_headers()
            self.end_headers()
            return

        status, response = result
        self.send_response(status)
        self._send_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        body = json.dumps(response, ensure_ascii=False, default=str)
        logger.info("[SSE] POST >>> %s (id=%s) -> %d (%d bytes)",
                    method, req_id, status, len(body))
        self.wfile.write(body.encode('utf-8'))

    def _handle_sse_message(self, msg: Dict):
        """POST /messages — same as POST /sse (legacy endpoint)."""
        self._handle_sse_post(msg)

    # ===============================================================
    # Backpressure execution
    # ===============================================================

    def _execute_with_backpressure(self, tool_name: str,
                                   arguments: Dict) -> Any:
        """Execute a tool on the VCL main thread with backpressure.

        1. Wait up to _WAIT_TIMEOUT to acquire the semaphore.
           If another tool is already running and doesn't finish in
           time, raise BusyError (-> HTTP 429).
        2. Once acquired, dispatch to the VCL main thread with
           _PROCESS_TIMEOUT.  If it doesn't finish, raise
           TimeoutError (-> HTTP 504).
        """
        acquired = _tool_semaphore.acquire(timeout=_WAIT_TIMEOUT)
        if not acquired:
            raise BusyError(
                "LibreOffice is busy processing another tool call. "
                "Please wait a moment and retry.")
        try:
            return execute_on_main_thread(
                self.mcp_server.execute_tool_sync,
                tool_name, arguments,
                timeout=_PROCESS_TIMEOUT,
            )
        finally:
            _tool_semaphore.release()

    # ===============================================================
    # Helpers
    # ===============================================================

    def _get_server_info(self) -> Dict[str, Any]:
        return {
            "name": "LibreOffice MCP Extension",
            "version": _get_version(),
            "description": "MCP server integrated into LibreOffice",
            "mcp_endpoint": "/mcp",
            "endpoints": {
                "POST /mcp": "MCP streamable-http (JSON-RPC 2.0)",
                "GET /mcp": "MCP streamable-http SSE notifications",
                "GET /sse": "MCP SSE transport (legacy, ChatGPT)",
                "POST /messages": "MCP SSE transport messages",
                "GET /": "Server info",
                "GET /health": "Health check",
            },
            "tools_count": len(self.mcp_server.tools),
        }

    def _read_body(self) -> Any:
        """Read and parse JSON body. Sends 400 on failure, returns None."""
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length == 0:
            return {}
        raw = self.rfile.read(content_length).decode('utf-8')
        try:
            parsed = json.loads(raw)
            logger.debug("Request body: %s", raw[:500])
            return parsed
        except json.JSONDecodeError:
            logger.warning("Invalid JSON body: %s", raw[:200])
            self._send_json(400, {"error": "Invalid JSON"})
            return None

    def _send_json(self, status: int, data: Any):
        self.send_response(status)
        self._send_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(
            data, ensure_ascii=False, default=str).encode('utf-8'))

    def _send_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods',
                         'GET, POST, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers',
                         'Content-Type, Authorization, Mcp-Session-Id')
        self.send_header('Access-Control-Expose-Headers',
                         'Mcp-Session-Id')

    def log_message(self, fmt, *args):
        logger.info("%s - %s", self.client_address[0], fmt % args)


# ===================================================================
# AIInterface (server lifecycle)
# ===================================================================

class AIInterface:
    """Threaded HTTP server hosting the MCP protocol."""

    def __init__(self, port: int = 8765, host: str = "localhost",
                 use_ssl: bool = True):
        self.port = port
        self.host = host
        self.use_ssl = use_ssl
        self.server = None
        self.server_thread = None
        self.running = False
        logger.info("AI Interface initialized for %s:%s (ssl=%s)",
                     host, port, use_ssl)

    def start(self):
        try:
            if self.running:
                logger.warning("Server is already running")
                return

            MCPRequestHandler.mcp_server = get_mcp_server()

            self.server = _ThreadedHTTPServer(
                (self.host, self.port), MCPRequestHandler)

            if self.use_ssl:
                from .ssl_certs import ensure_certs, create_ssl_context
                cert_path, key_path = ensure_certs()
                ssl_ctx = create_ssl_context(cert_path, key_path)
                self.server.socket = ssl_ctx.wrap_socket(
                    self.server.socket, server_side=True)
                logger.info("TLS enabled with cert %s", cert_path)

            self.running = True

            scheme = "https" if self.use_ssl else "http"
            logger.info("Started MCP %s server on %s:%s",
                        scheme.upper(), self.host, self.port)

            self.server_thread = threading.Thread(
                target=self._run_server, daemon=True)
            self.server_thread.start()

            scheme = "https" if self.use_ssl else "http"
            logger.info("MCP %s server ready - "
                        "MCP endpoint: %s://%s:%s/mcp",
                        scheme.upper(), scheme, self.host, self.port)

        except OSError as e:
            logger.error("Failed to bind %s:%s: %s",
                         self.host, self.port, e)
            self.running = False
            raise
        except Exception as e:
            logger.error("Failed to start HTTP server: %s", e)
            self.running = False
            raise

    def stop(self):
        try:
            if not self.running:
                return
            self.running = False
            if self.server:
                self.server.shutdown()
                self.server.server_close()
                logger.info("MCP HTTP server stopped")
        except Exception as e:
            logger.error("Error stopping HTTP server: %s", e)

    def _run_server(self):
        try:
            self.server.serve_forever()
        except Exception as e:
            if self.running:
                logger.error("HTTP server error: %s", e)
        finally:
            self.running = False

    def is_running(self) -> bool:
        return self.running

    def get_status(self) -> Dict[str, Any]:
        scheme = "https" if self.use_ssl else "http"
        return {
            "running": self.running,
            "host": self.host,
            "port": self.port,
            "ssl": self.use_ssl,
            "mcp_url": f"{scheme}://{self.host}:{self.port}/mcp",
            "thread_alive": (self.server_thread.is_alive()
                             if self.server_thread else False),
        }


# ===================================================================
# Module-level helpers
# ===================================================================

ai_interface = None

def get_ai_interface(port: int = 8765,
                     host: str = "localhost",
                     use_ssl: bool = True) -> AIInterface:
    global ai_interface
    if ai_interface is None:
        ai_interface = AIInterface(port, host, use_ssl=use_ssl)
    return ai_interface

def start_ai_interface(port: int = 8765,
                       host: str = "localhost",
                       use_ssl: bool = True) -> AIInterface:
    interface = get_ai_interface(port, host, use_ssl=use_ssl)
    if not interface.is_running():
        interface.start()
    return interface

def stop_ai_interface():
    global ai_interface
    if ai_interface:
        ai_interface.stop()
