"""
LibreOffice MCP Extension - Registration Module

Implements XDispatchProvider/XDispatch so menu items (Addons.xcu)
actually call our code when the user clicks them.
"""

import uno
import unohelper
import os
import logging
import threading
import traceback
import socket as _socket
import subprocess as _subprocess

from com.sun.star.frame import XDispatch, XDispatchProvider
from com.sun.star.lang import XServiceInfo, XInitialization
from com.sun.star.task import XJob
from com.sun.star.beans import PropertyValue

# ── File logger (so we can debug even when LO console is hidden) ─────────────

_log_path = os.path.join(os.path.expanduser("~"), "mcp-extension.log")
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(_log_path, mode="w", encoding="utf-8"),
    ],
)
logger = logging.getLogger("mcp-extension")

IMPLEMENTATION_NAME = "org.mcp.libreoffice.MCPExtension"
SERVICE_NAMES = ("com.sun.star.frame.ProtocolHandler",)

from .version import EXTENSION_VERSION

EXTENSION_NAME = "LibreOffice MCP"
EXTENSION_URL = "https://github.com/quazardous/mcp-libre"

logger.info("=== registration.py loaded — %s v%s ===", EXTENSION_NAME, EXTENSION_VERSION)

# ── Extension config ─────────────────────────────────────────────────────────
# Uses LibreOffice native configuration (MCPServerConfig.xcs/xcu)
# with fallback to hardcoded defaults if UNO config is not yet available.

import json as _json
from com.sun.star.beans import PropertyValue as _PropertyValue

_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".mcp-libre")
_PID_FILE = os.path.join(_CONFIG_DIR, "server.pid")

_CONFIG_NODE_PATH = "/org.mcp.libreoffice.Settings/Server"

_DEFAULT_CONFIG = {
    "autostart": True,
    "port": 8765,
    "host": "localhost",
    "enable_ssl": True,
    "enable_tunnel": False,
    "tunnel_provider": "tailscale",
    "tunnel_server": "bore.pub",
    "cf_tunnel_name": "",
    "cf_public_url": "",
    "ngrok_authtoken": "",
}

_TUNNEL_NODE_PATH = _CONFIG_NODE_PATH + "/Tunnel"


def _read_lo_config():
    """Read config from LibreOffice native configuration registry."""
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        provider = smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        node_arg = _PropertyValue()
        node_arg.Name = "nodepath"
        node_arg.Value = _CONFIG_NODE_PATH
        access = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (node_arg,))
        cfg = {
            "autostart": bool(access.getPropertyValue("AutoStart")),
            "port": int(access.getPropertyValue("Port")),
            "host": str(access.getPropertyValue("Host")),
        }
        try:
            cfg["enable_ssl"] = bool(access.getPropertyValue("EnableSSL"))
        except Exception:
            cfg["enable_ssl"] = True

        # Read nested Tunnel config
        try:
            tunnel = access.getByName("Tunnel")
            cfg["enable_tunnel"] = bool(tunnel.getPropertyValue("Enabled"))
            cfg["tunnel_provider"] = str(tunnel.getPropertyValue("Provider"))
            try:
                bore = tunnel.getByName("Bore")
                cfg["tunnel_server"] = str(bore.getPropertyValue("Server"))
            except Exception:
                cfg["tunnel_server"] = "bore.pub"
            try:
                cf = tunnel.getByName("Cloudflared")
                cfg["cf_tunnel_name"] = str(cf.getPropertyValue("TunnelName"))
                cfg["cf_public_url"] = str(cf.getPropertyValue("PublicUrl"))
            except Exception:
                cfg["cf_tunnel_name"] = ""
                cfg["cf_public_url"] = ""
            try:
                ngrok = tunnel.getByName("Ngrok")
                cfg["ngrok_authtoken"] = str(ngrok.getPropertyValue("Authtoken"))
            except Exception:
                cfg["ngrok_authtoken"] = ""
        except Exception:
            # Fallback for old flat schema
            cfg["enable_tunnel"] = False
            cfg["tunnel_provider"] = "tailscale"
            cfg["tunnel_server"] = "bore.pub"
            cfg["cf_tunnel_name"] = ""
            cfg["cf_public_url"] = ""
            cfg["ngrok_authtoken"] = ""

        access.dispose()
        logger.info(f"Config loaded from LO registry: {cfg}")
        return cfg
    except Exception as e:
        logger.warning(f"LO config read failed (using defaults): {e}")
        return None


def _write_lo_config(values):
    """Write config values to LibreOffice native configuration registry.

    Args:
        values: dict with keys like "autostart", "port", "host",
                "enable_tunnel", "tunnel_provider", "tunnel_server",
                "ngrok_authtoken".
    """
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        provider = smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        node_arg = _PropertyValue()
        node_arg.Name = "nodepath"
        node_arg.Value = _CONFIG_NODE_PATH
        update = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (node_arg,))

        # Server-level properties
        server_map = {"autostart": "AutoStart", "port": "Port", "host": "Host",
                      "enable_ssl": "EnableSSL"}
        for key, value in values.items():
            lo_key = server_map.get(key)
            if lo_key:
                update.setPropertyValue(lo_key, value)

        # Tunnel sub-group
        tunnel_keys = {"enable_tunnel", "tunnel_provider", "tunnel_server",
                       "cf_tunnel_name", "cf_public_url", "ngrok_authtoken"}
        if tunnel_keys & values.keys():
            tunnel = update.getByName("Tunnel")
            if "enable_tunnel" in values:
                tunnel.setPropertyValue("Enabled", values["enable_tunnel"])
            if "tunnel_provider" in values:
                tunnel.setPropertyValue("Provider", values["tunnel_provider"])
            if "tunnel_server" in values:
                bore = tunnel.getByName("Bore")
                bore.setPropertyValue("Server", values["tunnel_server"])
            if "cf_tunnel_name" in values or "cf_public_url" in values:
                cf = tunnel.getByName("Cloudflared")
                if "cf_tunnel_name" in values:
                    cf.setPropertyValue("TunnelName", values["cf_tunnel_name"])
                if "cf_public_url" in values:
                    cf.setPropertyValue("PublicUrl", values["cf_public_url"])
            if "ngrok_authtoken" in values:
                ngrok = tunnel.getByName("Ngrok")
                ngrok.setPropertyValue("Authtoken", values["ngrok_authtoken"])

        update.commitChanges()
        update.dispose()
        logger.info(f"Config written to LO registry: {values}")
    except Exception as e:
        logger.error(f"LO config write failed: {e}")


def _load_config():
    """Load config: try LO native registry first, then fall back to defaults."""
    cfg = _read_lo_config()
    if cfg is not None:
        return cfg
    return dict(_DEFAULT_CONFIG)


_config = _load_config()
logger.info(f"Config: autostart={_config['autostart']}, port={_config['port']}, ssl={_config.get('enable_ssl', True)}")

# ── Tunnel state ──────────────────────────────────────────────────────────────

_tunnel_state = {
    "process": None,     # subprocess.Popen
    "public_url": None,  # "bore.pub:43210"
}


def _stop_tunnel_process():
    """Stop tunnel process if running (module-level helper)."""
    proc = _tunnel_state["process"]
    if proc is None:
        return
    try:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except _subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=3)
    except Exception as e:
        logger.warning(f"Error stopping tunnel: {e}")
    # Reset Tailscale Funnel config (persists in daemon after process dies)
    if _config.get("tunnel_provider") == "tailscale":
        try:
            _subprocess.run(
                ["tailscale", "funnel", "reset"],
                capture_output=True, text=True, timeout=5,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            pass
    _tunnel_state["process"] = None
    _tunnel_state["public_url"] = None


# ── Port / zombie management ─────────────────────────────────────────────────


def _probe_health(host, port, timeout=2):
    """Probe the health endpoint. Returns True if OUR server responds.

    Tries HTTPS first (default), falls back to HTTP.
    """
    try:
        import http.client
        import ssl
        use_ssl = _config.get("enable_ssl", True)
        if use_ssl:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            conn = http.client.HTTPSConnection(
                host, port, timeout=timeout, context=ctx)
        else:
            conn = http.client.HTTPConnection(host, port, timeout=timeout)
        conn.request("GET", "/health")
        resp = conn.getresponse()
        body = resp.read().decode("utf-8", errors="replace")
        conn.close()
        return "LibreOffice MCP" in body
    except Exception:
        return False


def _is_port_bound(host, port, timeout=1):
    """Check if a port is bound (something is listening)."""
    try:
        s = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
        s.settimeout(timeout)
        s.connect((host, port))
        s.close()
        return True
    except Exception:
        return False


def _get_pids_on_port(port):
    """Get PIDs of processes listening on a port (Windows)."""
    pids = set()
    try:
        result = _subprocess.run(
            ["netstat", "-ano"],
            capture_output=True, text=True, timeout=5,
            creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
        )
        for line in result.stdout.splitlines():
            if f":{port}" in line and "LISTENING" in line:
                parts = line.split()
                pid = parts[-1]
                if pid.isdigit():
                    pids.add(int(pid))
    except Exception as e:
        logger.warning(f"Failed to enumerate PIDs on port {port}: {e}")
    return pids


def _kill_zombies_on_port(host, port):
    """Kill zombie processes on the port that aren't serving our MCP health."""
    if not _is_port_bound(host, port):
        logger.debug(f"Port {port} is free")
        return True  # Port is free

    if _probe_health(host, port):
        logger.info(f"Port {port} already has a healthy MCP server")
        return False  # Our server is already running, don't kill

    # Port is bound but not our server -> zombie
    pids = _get_pids_on_port(port)
    my_pid = os.getpid()
    for pid in pids:
        if pid == my_pid:
            continue
        logger.warning(f"Killing zombie process PID {pid} on port {port}")
        try:
            _subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True, timeout=5,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            logger.warning(f"Failed to kill PID {pid}: {e}")

    # Wait and verify
    import time
    time.sleep(1)
    if _is_port_bound(host, port):
        logger.error(f"Port {port} still bound after killing zombies")
        return False
    logger.info(f"Zombies cleared from port {port}")
    return True


def _write_pid_file():
    """Write current PID to server.pid."""
    try:
        os.makedirs(_CONFIG_DIR, exist_ok=True)
        with open(_PID_FILE, "w") as f:
            f.write(str(os.getpid()))
    except Exception:
        pass


def _remove_pid_file():
    """Remove server.pid."""
    try:
        if os.path.exists(_PID_FILE):
            os.unlink(_PID_FILE)
    except Exception:
        pass


# ── Shared state (singleton across dispatch calls) ───────────────────────────

_mcp_state = {
    "started": False,
    "mcp_server": None,
    "ai_interface": None,
}

# Cross-load guard: prevent two module loads from both auto-starting

# ── Server state & dynamic menu icons ─────────────────────────────────────

_STATE_STOPPED = "stopped"
_STATE_STARTING = "starting"
_STATE_RUNNING = "running"
_server_state = _STATE_STOPPED

_status_listeners_lock = threading.Lock()
_status_listeners_list = []  # [(listener, parsed_url), ...]


def _get_menu_text(url_path):
    """Return dynamic menu-item text based on current server state."""
    if url_path == "toggle_mcp_server":
        if _server_state == _STATE_RUNNING:
            return "Stop Server"
        if _server_state == _STATE_STARTING:
            return "Starting\u2026"
        return "Start Server"
    if url_path == "toggle_ssl":
        ssl_on = _config.get("enable_ssl", True)
        return "HTTPS: On" if ssl_on else "HTTPS: Off"
    if url_path == "toggle_tunnel":
        if _tunnel_state["process"] is not None:
            url = _tunnel_state["public_url"] or "..."
            display = url.replace("https://", "").replace("http://", "")
            return f"Stop Tunnel ({display})"
        return "Start Tunnel"
    return None


def _icon_name_for_state():
    """Return the icon filename prefix for the current server state."""
    if _server_state == _STATE_RUNNING:
        return "running"
    if _server_state == _STATE_STARTING:
        return "starting"
    return "stopped"


def _get_extension_url():
    """Get the base URL of the installed extension package."""
    try:
        ctx = uno.getComponentContext()
        pip = ctx.getByName(
            "/singletons/com.sun.star.deployment.PackageInformationProvider")
        return pip.getPackageLocation("org.mcp.libreoffice.extension")
    except Exception:
        return None


def _load_icon_graphic(icon_filename):
    """Load a PNG icon from the extension as XGraphic."""
    try:
        ext_url = _get_extension_url()
        logger.debug(f"Extension URL: {ext_url}")
        if not ext_url:
            logger.warning("Extension URL is None — icons won't load")
            return None
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        gp = smgr.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", ctx)
        pv = PropertyValue()
        pv.Name = "URL"
        icon_url = ext_url + "/icons/" + icon_filename
        pv.Value = icon_url
        logger.debug(f"Loading icon: {icon_url}")
        result = gp.queryGraphic((pv,))
        logger.debug(f"Icon loaded: {result is not None}")
        return result
    except Exception as e:
        logger.debug(f"Failed to load icon {icon_filename}: {e}")
        return None


# Command URLs that get dynamic icons
_ICON_CMDS = (
    "org.mcp.libreoffice:toggle_mcp_server",
    "org.mcp.libreoffice:get_status",
)

_MODULE_IDS = (
    "com.sun.star.text.TextDocument",
    "com.sun.star.sheet.SpreadsheetDocument",
    "com.sun.star.presentation.PresentationDocument",
    "com.sun.star.drawing.DrawingDocument",
)


def _update_menu_icons():
    """Push the current-state icon into every module's ImageManager."""
    return  # DISABLED: suspected cause of black menu rendering
    try:
        prefix = _icon_name_for_state()
        graphic = _load_icon_graphic(f"{prefix}_16.png")
        if graphic is None:
            logger.warning("Icon graphic is None for %s", prefix)
            return

        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        supplier = smgr.createInstanceWithContext(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier", ctx)

        ok_count = 0
        for mod_id in _MODULE_IDS:
            try:
                cfg_mgr = supplier.getUIConfigurationManager(mod_id)
                img_mgr = cfg_mgr.getImageManager()
                for cmd in _ICON_CMDS:
                    try:
                        if img_mgr.hasImage(0, cmd):
                            img_mgr.replaceImages(0, (cmd,), (graphic,))
                        else:
                            img_mgr.insertImages(0, (cmd,), (graphic,))
                        ok_count += 1
                    except Exception as e:
                        logger.info("ImageManager %s cmd %s: %s", mod_id, cmd, e)
            except Exception as e:
                logger.info("ImageManager skip %s: %s", mod_id, e)
        logger.info("Menu icons updated -> %s (%d insertions)", prefix, ok_count)
    except Exception as e:
        logger.warning("Dynamic icon update failed: %s", e)


def _fire_status_event(listener, url, text):
    """Send a FeatureStateEvent to one listener."""
    ev = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
    ev.FeatureURL = url
    ev.IsEnabled = True
    ev.Requery = False
    if text is not None:
        ev.State = text
    listener.statusChanged(ev)


def _set_server_state(new_state):
    """Update global server state and push it to every registered listener."""
    global _server_state
    _server_state = new_state
    logger.info(f"Server state -> {new_state}")
    _notify_all_listeners()
    # Update graphical icons in a thread (avoids blocking UI)
    threading.Thread(target=_update_menu_icons, daemon=True).start()


def _notify_all_listeners():
    """Push the current state to every registered status listener."""
    with _status_listeners_lock:
        alive = []
        for listener, url in _status_listeners_list:
            text = _get_menu_text(url.Path)
            try:
                _fire_status_event(listener, url, text)
                alive.append((listener, url))
            except Exception as e:
                logger.debug(f"Dropping dead status listener: {e}")
        _status_listeners_list[:] = alive


def _start_mcp_server():
    """Start the MCP HTTP server (called in a background thread)."""
    if _mcp_state["started"]:
        logger.warning("MCP server is already running (in-memory state)")
        _set_server_state(_STATE_RUNNING)
        return

    port = _config.get("port", 8765)
    host = _config.get("host", "localhost")

    # Check if our server is already running (e.g. from another module load)
    if _probe_health(host, port):
        logger.info("MCP server already healthy on port, marking as started")
        _mcp_state["started"] = True
        _set_server_state(_STATE_RUNNING)
        return

    _set_server_state(_STATE_STARTING)

    # Kill zombies on the port
    port_free = _kill_zombies_on_port(host, port)
    if not port_free:
        _set_server_state(_STATE_STOPPED)
        raise RuntimeError(f"Cannot free port {port}")

    try:
        from ai_interface import start_ai_interface
        from mcp_server import get_mcp_server

        use_ssl = _config.get("enable_ssl", True)
        _mcp_state["mcp_server"] = get_mcp_server()
        _mcp_state["ai_interface"] = start_ai_interface(
            port=port, host=host, use_ssl=use_ssl)
        _mcp_state["started"] = True
        _write_pid_file()
        _set_server_state(_STATE_RUNNING)
        scheme = "https" if use_ssl else "http"
        logger.info(f"MCP Extension started -- {scheme}://{host}:{port}")
    except Exception as e:
        _set_server_state(_STATE_STOPPED)
        logger.error(f"Failed to start MCP server: {e}")
        logger.error(traceback.format_exc())
        raise


def _stop_mcp_server():
    """Stop the MCP HTTP server."""
    if not _mcp_state["started"]:
        logger.warning("MCP server is not running")
        _set_server_state(_STATE_STOPPED)
        return
    try:
        from ai_interface import stop_ai_interface
        stop_ai_interface()
    except Exception:
        pass
    _mcp_state["ai_interface"] = None
    _mcp_state["mcp_server"] = None
    _mcp_state["started"] = False
    _remove_pid_file()
    # Auto-stop tunnel when server stops
    _stop_tunnel_process()
    _set_server_state(_STATE_STOPPED)
    logger.info("MCP Extension stopped")


# ── Message box helper ───────────────────────────────────────────────────────

def _msgbox(ctx, title, message):
    """Show an info message box in LibreOffice."""
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        frame = desktop.getCurrentFrame()
        if frame is None:
            logger.info(f"MSGBOX (no frame) - {title}: {message}")
            return
        window = frame.getContainerWindow()
        toolkit = smgr.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        box = toolkit.createMessageBox(
            window, 1, 1, title, message)  # INFOBOX, OK button
        box.execute()
    except Exception as e:
        logger.info(f"MSGBOX fallback - {title}: {message} (error: {e})")


def _copy_to_clipboard(ctx, text):
    """Copy text to system clipboard via LO API."""
    try:
        from com.sun.star.datatransfer import DataFlavor
        smgr = ctx.ServiceManager
        clip = smgr.createInstanceWithContext(
            "com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)

        class _TextTransferable(unohelper.Base,
                                __import__('com.sun.star.datatransfer',
                                           fromlist=['XTransferable'])
                                .XTransferable):
            def __init__(self, txt):
                self._text = txt

            def getTransferData(self, flavor):
                return self._text

            def getTransferDataFlavors(self):
                f = DataFlavor()
                f.MimeType = "text/plain;charset=utf-16"
                f.HumanPresentableName = "Unicode Text"
                f.DataType = uno.getTypeByName("string")
                return (f,)

            def isDataFlavorSupported(self, flavor):
                return "text/plain" in flavor.MimeType

        clip.setContents(_TextTransferable(text), None)
        logger.info(f"Copied to clipboard: {text}")
        return True
    except Exception as e:
        logger.error(f"Clipboard copy failed: {e}")
        return False


def _msgbox_with_copy(ctx, title, message, copy_text):
    """Show a dialog with message + Copy URL button."""
    try:
        smgr = ctx.ServiceManager

        dlg_model = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx)
        dlg_model.Title = title
        dlg_model.Width = 250
        dlg_model.Height = 80

        lbl = dlg_model.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel")
        lbl.Name = "Msg"
        lbl.PositionX = 10
        lbl.PositionY = 6
        lbl.Width = 230
        lbl.Height = 42
        lbl.MultiLine = True
        lbl.Label = message
        dlg_model.insertByName("Msg", lbl)

        copy_btn = dlg_model.createInstance(
            "com.sun.star.awt.UnoControlButtonModel")
        copy_btn.Name = "CopyBtn"
        copy_btn.PositionX = 10
        copy_btn.PositionY = 56
        copy_btn.Width = 75
        copy_btn.Height = 14
        copy_btn.Label = "Copy MCP URL"
        dlg_model.insertByName("CopyBtn", copy_btn)

        ok_btn = dlg_model.createInstance(
            "com.sun.star.awt.UnoControlButtonModel")
        ok_btn.Name = "OKBtn"
        ok_btn.PositionX = 190
        ok_btn.PositionY = 56
        ok_btn.Width = 50
        ok_btn.Height = 14
        ok_btn.Label = "OK"
        ok_btn.PushButtonType = 1  # OK
        dlg_model.insertByName("OKBtn", ok_btn)

        dlg = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dlg_model)
        toolkit = smgr.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        dlg.createPeer(toolkit, None)

        class _CopyListener(unohelper.Base,
                            __import__('com.sun.star.awt',
                                       fromlist=['XActionListener'])
                            .XActionListener):
            def __init__(self, dialog, context, text):
                self._dlg = dialog
                self._ctx = context
                self._text = text

            def actionPerformed(self, ev):
                if _copy_to_clipboard(self._ctx, self._text):
                    try:
                        self._dlg.getModel().getByName("CopyBtn").Label = \
                            "Copied!"
                    except Exception:
                        pass

            def disposing(self, ev):
                pass

        dlg.getControl("CopyBtn").addActionListener(
            _CopyListener(dlg, ctx, copy_text))

        dlg.execute()
        dlg.dispose()
    except Exception as e:
        logger.error(f"Copy dialog error: {e}")
        _msgbox(ctx, title, message)


# ── Protocol handler (dispatches menu clicks) ────────────────────────────────

class MCPExtension(unohelper.Base, XDispatch, XDispatchProvider,
                   XServiceInfo, XInitialization):
    """Protocol handler that dispatches menu actions."""

    def __init__(self, ctx):
        unohelper.Base.__init__(self)
        self.ctx = ctx
        self._frame = None
        logger.debug("MCPExtension.__init__ called")

    # -- XInitialization ------------------------------------------------------
    def initialize(self, args):
        try:
            if args:
                self._frame = args[0]
            logger.debug(f"MCPExtension.initialize called, frame={self._frame}")
            # Pre-load icons into ImageManager so first menu display has them
            threading.Thread(target=_update_menu_icons, daemon=True).start()
        except Exception as e:
            logger.error(f"initialize error: {e}")
            logger.error(traceback.format_exc())

    # -- XDispatchProvider ----------------------------------------------------
    def queryDispatch(self, url, target_frame_name, search_flags):
        try:
            logger.debug(f"queryDispatch: Complete={url.Complete} "
                         f"Protocol={url.Protocol} Path={url.Path}")
            if url.Protocol == "org.mcp.libreoffice:":
                return self
        except Exception as e:
            logger.error(f"queryDispatch error: {e}")
            logger.error(traceback.format_exc())
        return None

    def queryDispatches(self, requests):
        return [self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
                for r in requests]

    # -- XDispatch ------------------------------------------------------------
    def dispatch(self, url, args):
        """Called when the user clicks a menu item."""
        try:
            complete = url.Complete
            logger.info(f"dispatch called: {complete}")

            # URL format: org.mcp.libreoffice:<command>
            cmd = url.Path

            logger.info(f"command: {cmd}")

            if cmd == "start_mcp_server":
                _set_server_state(_STATE_STARTING)
                threading.Thread(target=self._do_start, daemon=True).start()
            elif cmd == "stop_mcp_server":
                self._do_stop()
            elif cmd == "restart_mcp_server":
                self._do_restart()
            elif cmd == "toggle_mcp_server":
                if _mcp_state["started"]:
                    self._do_stop()
                else:
                    _set_server_state(_STATE_STARTING)
                    threading.Thread(target=self._do_start, daemon=True).start()
            elif cmd == "toggle_ssl":
                self._do_toggle_ssl()
            elif cmd == "toggle_tunnel":
                self._do_toggle_tunnel()
            elif cmd == "get_status":
                self._do_status()
            elif cmd == "about":
                self._do_about()
            else:
                logger.warning(f"Unknown command: {cmd}")
                _msgbox(self.ctx, "MCP Extension",
                        f"Unknown command: {cmd}\nURL: {complete}")
        except Exception as e:
            logger.error(f"Error in dispatch: {e}")
            logger.error(traceback.format_exc())
            _msgbox(self.ctx, "MCP Error", str(e))

    def addStatusListener(self, listener, url):
        with _status_listeners_lock:
            _status_listeners_list.append((listener, url))
        # Send current state immediately so the menu renders correctly
        text = _get_menu_text(url.Path)
        if text is not None:
            try:
                _fire_status_event(listener, url, text)
            except Exception as e:
                logger.debug(f"Initial status event failed: {e}")

    def removeStatusListener(self, listener, url):
        with _status_listeners_lock:
            _status_listeners_list[:] = [
                (l, u) for l, u in _status_listeners_list
                if not (l is listener and u.Complete == url.Complete)
            ]

    # -- Actions --------------------------------------------------------------

    def _do_start(self):
        try:
            _start_mcp_server()
            host = _config.get("host", "localhost")
            port = _config.get("port", 8765)
            scheme = "https" if _config.get("enable_ssl", True) else "http"
            _msgbox(self.ctx, "MCP Server",
                    f"MCP server started.\n{scheme}://{host}:{port}")
            # Auto-start tunnel if enabled
            if _config.get("enable_tunnel", False) and _tunnel_state["process"] is None:
                self._do_start_tunnel()
        except Exception as e:
            _msgbox(self.ctx, "MCP Error", f"Failed to start:\n{e}")

    def _do_stop(self):
        _stop_mcp_server()
        _msgbox(self.ctx, "MCP Server", "MCP server stopped.")

    def _do_toggle_ssl(self):
        """Toggle HTTPS on/off, save config, restart server if running."""
        old_ssl = _config.get("enable_ssl", True)
        new_ssl = not old_ssl
        _config["enable_ssl"] = new_ssl
        _write_lo_config({"enable_ssl": new_ssl})
        _notify_all_listeners()  # update menu text
        scheme = "https" if new_ssl else "http"
        if _mcp_state["started"]:
            _stop_mcp_server()
            host = _config.get("host", "localhost")
            port = _config.get("port", 8765)
            _kill_zombies_on_port(host, port)
            _set_server_state(_STATE_STARTING)
            threading.Thread(target=self._do_start, daemon=True).start()
            _msgbox(self.ctx, "MCP Server",
                    f"Switched to {scheme.upper()}.\nServer restarting...")
        else:
            _msgbox(self.ctx, "MCP Server",
                    f"Switched to {scheme.upper()}.\n"
                    f"Will use {scheme}:// on next start.")

    def _do_toggle_tunnel(self):
        """Toggle tunnel on/off."""
        if _tunnel_state["process"] is not None:
            self._do_stop_tunnel()
        else:
            threading.Thread(target=self._do_start_tunnel, daemon=True).start()

    def _do_start_tunnel(self):
        """Start tunnel — dispatches to provider-specific method."""
        provider = _config.get("tunnel_provider", "tailscale")
        if provider == "bore":
            self._do_start_tunnel_bore()
        elif provider == "cloudflared":
            self._do_start_tunnel_cloudflared()
        elif provider == "ngrok":
            self._do_start_tunnel_ngrok()
        elif provider == "tailscale":
            self._do_start_tunnel_tailscale()
        else:
            _msgbox(self.ctx, "MCP Server",
                    f"Unknown tunnel provider: {provider}")

    def _tunnel_ensure_server(self):
        """Auto-start MCP server if not running. Returns port and scheme."""
        if not _mcp_state["started"]:
            _set_server_state(_STATE_STARTING)
            _start_mcp_server()
        port = _config.get("port", 8765)
        scheme = "https" if _config.get("enable_ssl", True) else "http"
        return port, scheme

    def _tunnel_check_binary(self, name, version_args, install_url):
        """Check if a tunnel binary is installed. Returns True if OK."""
        try:
            _subprocess.run(
                version_args,
                capture_output=True, timeout=5,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
            return True
        except FileNotFoundError:
            _msgbox(self.ctx, "MCP Server",
                    f"{name} not found.\n\nInstall from:\n{install_url}")
            return False
        except Exception as e:
            _msgbox(self.ctx, "MCP Server", f"{name} check failed: {e}")
            return False

    def _tunnel_run_and_parse(self, cmd, label, url_regex):
        """Run a tunnel subprocess, parse URL from stdout, show msgbox.

        Common logic for bore and cloudflared (line-by-line regex parse).
        """
        import re
        try:
            proc = _subprocess.Popen(
                cmd,
                stdout=_subprocess.PIPE,
                stderr=_subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
            _tunnel_state["process"] = proc
            _notify_all_listeners()

            url_found = False
            for line in proc.stdout:
                line = line.strip()
                logger.info(f"{label}: {line}")
                m = re.search(url_regex, line)
                if m and not url_found:
                    public_url = m.group(1)
                    _tunnel_state["public_url"] = public_url
                    url_found = True
                    _notify_all_listeners()
                    # Build display URL — bore gives host:port, others give full URL
                    if public_url.startswith("http"):
                        base_url = public_url
                    else:
                        scheme = "https" if _config.get("enable_ssl", True) else "http"
                        base_url = f"{scheme}://{public_url}"
                    _msgbox_with_copy(self.ctx, "MCP Server",
                            f"Tunnel active!\n\n"
                            f"ChatGPT:      {base_url}/sse\n"
                            f"Claude Code:  {base_url}/mcp",
                            f"{base_url}/sse")

            rc = proc.wait()
            if _tunnel_state["process"] is proc:
                _tunnel_state["process"] = None
                _tunnel_state["public_url"] = None
                _notify_all_listeners()
                if not url_found:
                    _msgbox(self.ctx, "MCP Server",
                            f"{label} exited (code {rc}) without establishing tunnel.")
                else:
                    logger.info(f"{label} tunnel closed (exit code {rc})")
        except Exception as e:
            logger.error(f"{label} tunnel error: {e}")
            _tunnel_state["process"] = None
            _tunnel_state["public_url"] = None
            _notify_all_listeners()
            _msgbox(self.ctx, "MCP Server", f"Tunnel error:\n{e}")

    def _do_start_tunnel_bore(self):
        """Start bore tunnel."""
        if not self._tunnel_check_binary(
                "bore", ["bore", "--version"],
                "https://github.com/ekzhang/bore/releases"):
            return
        port, scheme = self._tunnel_ensure_server()
        server = _config.get("tunnel_server", "bore.pub")
        logger.info(f"Starting bore tunnel: localhost:{port} -> {server}")
        self._tunnel_run_and_parse(
            ["bore", "local", str(port), "--to", server],
            "bore",
            r"listening at ([\w.\-]+:\d+)")

    def _do_start_tunnel_cloudflared(self):
        """Start cloudflared tunnel (quick or named)."""
        if not self._tunnel_check_binary(
                "cloudflared", ["cloudflared", "--version"],
                "https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/"):
            return
        port, scheme = self._tunnel_ensure_server()

        tunnel_name = _config.get("cf_tunnel_name", "").strip()
        public_url = _config.get("cf_public_url", "").strip()

        if tunnel_name:
            # Named tunnel — stable URL via Cloudflare account + domain
            logger.info(f"Starting cloudflared named tunnel '{tunnel_name}': "
                        f"{scheme}://localhost:{port}")
            cmd = ["cloudflared", "tunnel", "--no-autoupdate", "run", tunnel_name]
            if public_url:
                # URL is known from config, set it directly
                _tunnel_state["public_url"] = public_url.rstrip("/")
                _notify_all_listeners()
                _msgbox_with_copy(self.ctx, "MCP Server",
                        f"Named tunnel '{tunnel_name}' starting...\n\n"
                        f"ChatGPT:      {public_url}/sse\n"
                        f"Claude Code:  {public_url}/mcp",
                        f"{public_url}/sse")
            # Run tunnel in background thread (no URL parsing needed
            # for named tunnels with known public URL)
            self._tunnel_run_and_parse(
                cmd, f"cloudflared[{tunnel_name}]",
                # Named tunnels log "Connection ... registered" but not a public URL.
                # If public_url is set, this regex won't match — that's OK.
                # If public_url is NOT set, try to catch any https:// URL in output.
                r"(https://[\w.-]+\.\w{2,})" if not public_url else r"(?!)")
        else:
            # Quick tunnel — random trycloudflare.com URL
            logger.info(f"Starting cloudflared quick tunnel: {scheme}://localhost:{port}")
            self._tunnel_run_and_parse(
                ["cloudflared", "tunnel", "--no-autoupdate",
                 "--url", f"{scheme}://localhost:{port}", "--no-tls-verify"],
                "cloudflared",
                r"(https://[\w-]+\.trycloudflare\.com)")

    def _do_start_tunnel_ngrok(self):
        """Start ngrok tunnel."""
        import json as _json
        if not self._tunnel_check_binary(
                "ngrok", ["ngrok", "version"],
                "https://ngrok.com/download"):
            return
        port, scheme = self._tunnel_ensure_server()
        logger.info(f"Starting ngrok tunnel: {scheme}://localhost:{port}")

        cmd = ["ngrok", "http", f"{scheme}://localhost:{port}",
               "--log", "stdout", "--log-format", "json"]
        # Use authtoken from settings if provided
        authtoken = _config.get("ngrok_authtoken", "").strip()
        if authtoken:
            cmd.extend(["--authtoken", authtoken])

        try:
            proc = _subprocess.Popen(
                cmd,
                stdout=_subprocess.PIPE,
                stderr=_subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
            )
            _tunnel_state["process"] = proc
            _notify_all_listeners()

            url_found = False
            for line in proc.stdout:
                line = line.strip()
                if not line:
                    continue
                logger.info(f"ngrok: {line}")
                try:
                    data = _json.loads(line)
                    # Detect authtoken error
                    err_msg = str(data.get("err", ""))
                    if "ERR_NGROK_105" in err_msg or "authtoken" in err_msg.lower():
                        _msgbox(self.ctx, "MCP Server",
                                "ngrok requires an authtoken.\n\n"
                                "Run: ngrok config add-authtoken <TOKEN>\n"
                                "Get yours at:\nhttps://dashboard.ngrok.com/get-started/your-authtoken")
                        proc.terminate()
                        break
                    # Look for tunnel URL
                    if data.get("msg") == "started tunnel" and "url" in data:
                        public_url = data["url"]
                        _tunnel_state["public_url"] = public_url
                        url_found = True
                        _notify_all_listeners()
                        _msgbox_with_copy(self.ctx, "MCP Server",
                                f"Tunnel active!\n\n"
                                f"ChatGPT:      {public_url}/sse\n"
                                f"Claude Code:  {public_url}/mcp",
                                f"{public_url}/sse")
                except (_json.JSONDecodeError, KeyError):
                    pass

            rc = proc.wait()
            if _tunnel_state["process"] is proc:
                _tunnel_state["process"] = None
                _tunnel_state["public_url"] = None
                _notify_all_listeners()
                if not url_found:
                    _msgbox(self.ctx, "MCP Server",
                            f"ngrok exited (code {rc}) without establishing tunnel.")
                else:
                    logger.info(f"ngrok tunnel closed (exit code {rc})")
        except Exception as e:
            logger.error(f"ngrok tunnel error: {e}")
            _tunnel_state["process"] = None
            _tunnel_state["public_url"] = None
            _notify_all_listeners()
            _msgbox(self.ctx, "MCP Server", f"Tunnel error:\n{e}")

    def _do_start_tunnel_tailscale(self):
        """Start Tailscale Funnel tunnel."""
        if not self._tunnel_check_binary(
                "tailscale", ["tailscale", "version"],
                "https://tailscale.com/download"):
            return
        port, scheme = self._tunnel_ensure_server()
        # Reset any lingering serve/funnel config to avoid
        # "listener already exists for port 443" errors.
        for reset_cmd in [["tailscale", "funnel", "reset"],
                          ["tailscale", "serve", "reset"]]:
            try:
                _subprocess.run(
                    reset_cmd,
                    capture_output=True, text=True, timeout=5,
                    creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception as e:
                logger.debug(f"tailscale reset: {e}")
        # HTTP  backend → tailscale funnel <port>
        # HTTPS backend → tailscale funnel https+insecure://127.0.0.1:<port>
        if scheme == "https":
            target = f"https+insecure://127.0.0.1:{port}"
        else:
            target = str(port)
        logger.info(f"Starting Tailscale Funnel: {scheme}://localhost:{port}")
        self._tunnel_run_and_parse(
            ["tailscale", "funnel", target],
            "tailscale",
            r"(https://[\w.-]+\.ts\.net)")

    def _do_stop_tunnel(self):
        """Stop tunnel (any provider)."""
        proc = _tunnel_state["process"]
        if proc is None:
            return
        provider = _config.get("tunnel_provider", "tailscale")
        try:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except _subprocess.TimeoutExpired:
                proc.kill()
                proc.wait(timeout=3)
        except Exception as e:
            logger.warning(f"Error stopping tunnel: {e}")
        # Tailscale Funnel persists in daemon — reset it on stop
        if provider == "tailscale":
            for reset_cmd in [["tailscale", "funnel", "reset"],
                              ["tailscale", "serve", "reset"]]:
                try:
                    _subprocess.run(
                        reset_cmd,
                        capture_output=True, text=True, timeout=5,
                        creationflags=getattr(_subprocess, "CREATE_NO_WINDOW", 0),
                    )
                except Exception as e:
                    logger.debug(f"tailscale reset on stop: {e}")
            logger.info("Tailscale serve/funnel config reset")
        _tunnel_state["process"] = None
        _tunnel_state["public_url"] = None
        _notify_all_listeners()
        _msgbox(self.ctx, "MCP Server", "Tunnel stopped.")

    def _do_restart(self):
        """Stop, kill zombies, and restart."""
        _stop_mcp_server()
        port = _config.get("port", 8765)
        host = _config.get("host", "localhost")
        _kill_zombies_on_port(host, port)
        _set_server_state(_STATE_STARTING)
        threading.Thread(target=self._do_start, daemon=True).start()

    def _do_status(self):
        port = _config.get("port", 8765)
        host = _config.get("host", "localhost")
        mem_state = _mcp_state["started"]

        # Build instant status from in-memory state
        initial_status = "STARTED" if mem_state else "STOPPED"
        initial_lines = [
            f"MCP Server: {initial_status}",
            f"Version: {EXTENSION_VERSION}",
            f"Port: {port}  Host: {host}",
            f"Autostart: {_config.get('autostart', True)}",
            "",
            "Health check: probing...",
        ]

        # Create a UNO dialog so we can update it live
        try:
            ctx = self.ctx
            smgr = ctx.ServiceManager

            dlg_model = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialogModel", ctx)
            dlg_model.Title = "MCP Server Status"
            dlg_model.Width = 230
            dlg_model.Height = 110

            lbl = dlg_model.createInstance(
                "com.sun.star.awt.UnoControlFixedTextModel")
            lbl.Name = "StatusText"
            lbl.PositionX = 10
            lbl.PositionY = 6
            lbl.Width = 210
            lbl.Height = 72
            lbl.MultiLine = True
            lbl.Label = "\n".join(initial_lines)
            dlg_model.insertByName("StatusText", lbl)

            # "Copy URL" button — visible only when tunnel is active
            has_tunnel = bool(_tunnel_state.get("public_url"))
            copy_btn = dlg_model.createInstance(
                "com.sun.star.awt.UnoControlButtonModel")
            copy_btn.Name = "CopyBtn"
            copy_btn.PositionX = 10
            copy_btn.PositionY = 88
            copy_btn.Width = 65
            copy_btn.Height = 14
            copy_btn.Label = "Copy MCP URL"
            copy_btn.Enabled = has_tunnel
            dlg_model.insertByName("CopyBtn", copy_btn)

            btn = dlg_model.createInstance(
                "com.sun.star.awt.UnoControlButtonModel")
            btn.Name = "OKBtn"
            btn.PositionX = 170
            btn.PositionY = 88
            btn.Width = 50
            btn.Height = 14
            btn.Label = "OK"
            btn.PushButtonType = 1  # OK
            dlg_model.insertByName("OKBtn", btn)

            dlg = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialog", ctx)
            dlg.setModel(dlg_model)
            toolkit = smgr.createInstanceWithContext(
                "com.sun.star.awt.Toolkit", ctx)
            dlg.createPeer(toolkit, None)

            # Wire up Copy button click
            class CopyListener(unohelper.Base,
                               __import__('com.sun.star.awt',
                                          fromlist=['XActionListener'])
                               .XActionListener):
                def __init__(self, dialog, context):
                    self._dlg = dialog
                    self._ctx = context

                def actionPerformed(self, ev):
                    pub = _tunnel_state.get("public_url", "")
                    if not pub:
                        return
                    if pub.startswith("http"):
                        mcp_url = f"{pub}/sse"
                    else:
                        s = "https" if _config.get("enable_ssl", True) else "http"
                        mcp_url = f"{s}://{pub}/sse"
                    # Copy to system clipboard via LO
                    try:
                        clip = self._ctx.ServiceManager.createInstanceWithContext(
                            "com.sun.star.datatransfer.clipboard.SystemClipboard",
                            self._ctx)
                        from com.sun.star.datatransfer import DataFlavor
                        class TextTransferable(unohelper.Base,
                                               __import__('com.sun.star.datatransfer',
                                                          fromlist=['XTransferable'])
                                               .XTransferable):
                            def __init__(self, text):
                                self._text = text
                            def getTransferData(self, flavor):
                                return self._text
                            def getTransferDataFlavors(self):
                                f = DataFlavor()
                                f.MimeType = "text/plain;charset=utf-16"
                                f.HumanPresentableName = "Unicode Text"
                                f.DataType = uno.getTypeByName("string")
                                return (f,)
                            def isDataFlavorSupported(self, flavor):
                                return "text/plain" in flavor.MimeType
                        clip.setContents(TextTransferable(mcp_url), None)
                        # Update button text briefly
                        try:
                            self._dlg.getModel().getByName("CopyBtn").Label = "Copied!"
                        except Exception:
                            pass
                    except Exception as ex:
                        logger.error(f"Clipboard copy failed: {ex}")

                def disposing(self, ev):
                    pass

            dlg.getControl("CopyBtn").addActionListener(
                CopyListener(dlg, ctx))

            # Background probe updates the label while dialog is open
            def probe_and_update():
                try:
                    import time
                    time.sleep(0.05)
                    http_ok = _probe_health(host, port, timeout=1)
                    if mem_state:
                        health = "OK" if http_ok else "FAIL (not responding)"
                    else:
                        if _is_port_bound(host, port, timeout=0.5):
                            health = "ZOMBIE (port bound, server stopped)"
                        else:
                            health = "-- (server stopped)"
                    lines = [
                        f"MCP Server: {initial_status}",
                        f"Version: {EXTENSION_VERSION}",
                        f"Port: {port}  Host: {host}",
                        f"Autostart: {_config.get('autostart', True)}",
                        "",
                        f"Health check: {health}",
                    ]
                    if mem_state and http_ok:
                        scheme = "https" if _config.get("enable_ssl", True) else "http"
                        lines.append(f"URL: {scheme}://{host}:{port}/health")
                    if _tunnel_state["public_url"]:
                        lines.append(f"Tunnel: {_tunnel_state['public_url']}")
                    try:
                        dlg_model.getByName("StatusText").Label = "\n".join(lines)
                        # Enable/disable Copy button based on tunnel state
                        copy_m = dlg_model.getByName("CopyBtn")
                        if _tunnel_state.get("public_url"):
                            copy_m.Enabled = True
                            if copy_m.Label == "Copy MCP URL":
                                pass  # keep label
                        else:
                            copy_m.Enabled = False
                            copy_m.Label = "Copy MCP URL"
                    except Exception:
                        pass  # dialog already closed
                except Exception as e:
                    logger.debug(f"Status probe error: {e}")

            threading.Thread(target=probe_and_update, daemon=True).start()
            dlg.execute()
            dlg.dispose()
        except Exception as e:
            logger.error(f"Status dialog error: {e}")
            # Fallback to simple msgbox
            _msgbox(self.ctx, "MCP Server Status",
                    f"Status: {initial_status}\nVersion: {EXTENSION_VERSION}")

    def _do_about(self):
        lines = [
            f"{EXTENSION_NAME}",
            f"Version: {EXTENSION_VERSION}",
            f"",
            f"MCP (Model Context Protocol) server extension",
            f"for LibreOffice document manipulation via AI.",
            f"",
            f"GitHub: {EXTENSION_URL}",
        ]
        _msgbox(self.ctx, "About MCP Extension", "\n".join(lines))

    # -- XServiceInfo ---------------------------------------------------------
    def getImplementationName(self):
        return IMPLEMENTATION_NAME

    def supportsService(self, name):
        return name in SERVICE_NAMES

    def getSupportedServiceNames(self):
        return SERVICE_NAMES


# ── Options dialog handler (Tools > Options > MCP Server) ────────────────────

from com.sun.star.awt import XContainerWindowEventHandler


class MCPOptionsHandler(unohelper.Base, XContainerWindowEventHandler,
                        XServiceInfo):
    """Event handler for the MCP Server options page in Tools > Options."""

    _IMPL_NAME = "org.mcp.libreoffice.MCPOptionsHandler"
    _SVC_NAMES = ("org.mcp.libreoffice.MCPOptionsHandler",)

    def __init__(self, ctx):
        unohelper.Base.__init__(self)
        self.ctx = ctx
        logger.debug("MCPOptionsHandler.__init__")

    # -- XContainerWindowEventHandler -----------------------------------------

    def callHandlerMethod(self, xWindow, eventObject, methodName):
        logger.debug(f"Options handler: {methodName} / {eventObject}")
        if methodName != "external_event":
            return False

        if eventObject == "initialize":
            self._load_to_dialog(xWindow)
            return True
        elif eventObject == "ok":
            self._save_from_dialog(xWindow)
            return True
        elif eventObject == "back":
            self._load_to_dialog(xWindow)
            return True
        return False

    def getSupportedMethodNames(self):
        return ("external_event",)

    # -- Load / Save ----------------------------------------------------------
    # Two pages share this handler: "Server" (MCPSettings.xdl) and
    # "Tunnel" (MCPTunnel.xdl). We detect which page by checking for
    # a control unique to each.

    def _is_tunnel_page(self, xWindow):
        try:
            return xWindow.getControl("TunnelCheck") is not None
        except Exception:
            return False

    def _load_to_dialog(self, xWindow):
        try:
            cfg = _read_lo_config()
            if cfg is None:
                cfg = dict(_DEFAULT_CONFIG)
            logger.debug(f"Options: loading {cfg}")

            if self._is_tunnel_page(xWindow):
                self._load_tunnel_page(xWindow, cfg)
            else:
                self._load_server_page(xWindow, cfg)
        except Exception as e:
            logger.error(f"Options load error: {e}")
            logger.error(traceback.format_exc())

    def _load_server_page(self, xWindow, cfg):
        xWindow.getControl("AutoStartCheck").setState(
            1 if cfg["autostart"] else 0)
        xWindow.getControl("HostField").setText(cfg["host"])
        xWindow.getControl("PortField").setValue(float(cfg["port"]))
        xWindow.getControl("SSLCheck").setState(
            1 if cfg.get("enable_ssl", True) else 0)
        scheme = "https" if cfg.get("enable_ssl", True) else "http"
        xWindow.getControl("UrlText").setText(
            f"Health check: {scheme}://{cfg['host']}:{cfg['port']}/health")

    _TUNNEL_PROVIDERS = ["tailscale", "cloudflared", "bore", "ngrok"]

    # Provider-specific field groups for visibility toggling
    _PROVIDER_FIELDS = {
        "cloudflared": ["CfTunnelNameLabel", "CfTunnelNameField",
                        "CfPublicUrlLabel", "CfPublicUrlField"],
        "bore": ["TunnelServerLabel", "TunnelServerField"],
        "ngrok": ["NgrokTokenLabel", "NgrokTokenField"],
        "tailscale": [],
    }

    def _show_provider_fields(self, xWindow, provider):
        """Show only the fields relevant to the selected provider."""
        all_fields = set()
        for fields in self._PROVIDER_FIELDS.values():
            all_fields.update(fields)
        active_fields = set(self._PROVIDER_FIELDS.get(provider, []))
        for name in all_fields:
            try:
                ctrl = xWindow.getControl(name)
                ctrl.setVisible(name in active_fields)
            except Exception:
                pass

    def _load_tunnel_page(self, xWindow, cfg):
        xWindow.getControl("TunnelCheck").setState(
            1 if cfg.get("enable_tunnel", False) else 0)
        # Populate provider dropdown
        provider_list = xWindow.getControl("TunnelProviderList")
        provider_list.removeItems(0, provider_list.getItemCount())
        for p in self._TUNNEL_PROVIDERS:
            provider_list.addItem(p, provider_list.getItemCount())
        provider = cfg.get("tunnel_provider", "cloudflared")
        try:
            provider_list.selectItemPos(
                self._TUNNEL_PROVIDERS.index(provider), True)
        except ValueError:
            provider_list.selectItemPos(0, True)
        # Provider-specific fields
        xWindow.getControl("TunnelServerField").setText(
            cfg.get("tunnel_server", "bore.pub"))
        xWindow.getControl("CfTunnelNameField").setText(
            cfg.get("cf_tunnel_name", ""))
        xWindow.getControl("CfPublicUrlField").setText(
            cfg.get("cf_public_url", ""))
        xWindow.getControl("NgrokTokenField").setText(
            cfg.get("ngrok_authtoken", ""))
        # Show/hide fields for selected provider
        self._show_provider_fields(xWindow, provider)
        # Attach listener to toggle fields on provider change
        try:
            XItemListener = __import__(
                'com.sun.star.awt', fromlist=['XItemListener']).XItemListener

            handler = self

            class ProviderChangeListener(unohelper.Base, XItemListener):
                def __init__(self, window):
                    self._win = window

                def itemStateChanged(self, ev):
                    sel = ev.Selected
                    providers = handler._TUNNEL_PROVIDERS
                    if 0 <= sel < len(providers):
                        handler._show_provider_fields(
                            self._win, providers[sel])

                def disposing(self, ev):
                    pass

            provider_list.addItemListener(ProviderChangeListener(xWindow))
        except Exception as e:
            logger.warning(f"Could not attach provider listener: {e}")

    def _save_from_dialog(self, xWindow):
        try:
            if self._is_tunnel_page(xWindow):
                self._save_tunnel_page(xWindow)
            else:
                self._save_server_page(xWindow)
        except Exception as e:
            logger.error(f"Options save error: {e}")
            logger.error(traceback.format_exc())

    def _save_server_page(self, xWindow):
        new_values = {
            "autostart": xWindow.getControl("AutoStartCheck").getState() == 1,
            "port": int(xWindow.getControl("PortField").getValue()),
            "host": xWindow.getControl("HostField").getText(),
            "enable_ssl": xWindow.getControl("SSLCheck").getState() == 1,
        }
        logger.info(f"Options (server): saving {new_values}")

        # Detect if server needs restart (port, host, or SSL changed)
        old_port = _config.get("port", 8765)
        old_host = _config.get("host", "localhost")
        old_ssl = _config.get("enable_ssl", True)
        needs_restart = (
            _mcp_state["started"]
            and (new_values["port"] != old_port
                 or new_values["host"] != old_host
                 or new_values["enable_ssl"] != old_ssl)
        )

        _write_lo_config(new_values)
        _config.update(new_values)

        if needs_restart:
            logger.info(f"Options: port/host changed "
                        f"({old_host}:{old_port} -> "
                        f"{new_values['host']}:{new_values['port']}), "
                        f"restarting server...")
            _stop_mcp_server()
            _kill_zombies_on_port(old_host, old_port)
            threading.Thread(
                target=self._restart_after_config_change,
                daemon=True).start()

    def _save_tunnel_page(self, xWindow):
        provider_list = xWindow.getControl("TunnelProviderList")
        sel = provider_list.getSelectedItemPos()
        providers = self._TUNNEL_PROVIDERS
        new_values = {
            "enable_tunnel": xWindow.getControl("TunnelCheck").getState() == 1,
            "tunnel_provider": providers[sel] if 0 <= sel < len(providers) else "tailscale",
            "tunnel_server": xWindow.getControl("TunnelServerField").getText(),
            "cf_tunnel_name": xWindow.getControl("CfTunnelNameField").getText(),
            "cf_public_url": xWindow.getControl("CfPublicUrlField").getText(),
            "ngrok_authtoken": xWindow.getControl("NgrokTokenField").getText(),
        }
        logger.info(f"Options (tunnel): saving {new_values}")
        _write_lo_config(new_values)
        _config.update(new_values)

    def _restart_after_config_change(self):
        """Restart server on new port after config change."""
        try:
            import time
            time.sleep(1)  # Let old socket close
            _start_mcp_server()
            logger.info("Server restarted on new config")
        except Exception as e:
            logger.error(f"Restart after config change failed: {e}")

    # -- XServiceInfo ---------------------------------------------------------

    def getImplementationName(self):
        return self._IMPL_NAME

    def supportsService(self, name):
        return name in self._SVC_NAMES

    def getSupportedServiceNames(self):
        return self._SVC_NAMES


# ── AutoStart Job (fires on onFirstVisibleTask) ─────────────────────────────

class MCPAutoStartJob(unohelper.Base, XJob, XServiceInfo):
    """LO Job triggered by onFirstVisibleTask — starts MCP server at LO launch."""
    _IMPL_NAME = "org.mcp.libreoffice.MCPAutoStart"
    _SVC_NAMES = ("com.sun.star.task.Job",)

    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        logger.info("AutoStart job triggered (onFirstVisibleTask)")
        try:
            if _config.get("autostart", True):
                threading.Thread(
                    target=_start_mcp_server, daemon=True,
                    name="mcp-job-start").start()
            else:
                logger.info("AutoStart job: autostart disabled in config")
        except Exception as e:
            logger.error(f"AutoStart job failed: {e}")
        return ()

    def getImplementationName(self):
        return self._IMPL_NAME

    def supportsService(self, name):
        return name in self._SVC_NAMES

    def getSupportedServiceNames(self):
        return self._SVC_NAMES


# ── UNO component registration ──────────────────────────────────────────────

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MCPExtension,
    IMPLEMENTATION_NAME,
    SERVICE_NAMES,
)
g_ImplementationHelper.addImplementation(
    MCPOptionsHandler,
    MCPOptionsHandler._IMPL_NAME,
    MCPOptionsHandler._SVC_NAMES,
)
g_ImplementationHelper.addImplementation(
    MCPAutoStartJob,
    MCPAutoStartJob._IMPL_NAME,
    MCPAutoStartJob._SVC_NAMES,
)

logger.info("g_ImplementationHelper configured (3 services)")

# ── Module-level auto-start (fallback) ───────────────────────────────────────
# OnStartApp fires before extensions load, so the Jobs.xcu-based auto-start
# often misses the event.  This fallback ensures the server starts when
# registration.py is loaded by the runtime (second load).
# The first load (UNO registration in unopkg) creates a daemon thread that
# dies silently when the process exits — harmless.

def _module_autostart():
    """Delayed auto-start: wait for LO to be ready, then start if needed."""
    import time
    time.sleep(3)
    if _config.get("autostart", True) and not _mcp_state["started"]:
        logger.info("Module-level auto-start (fallback)")
        try:
            _start_mcp_server()
        except Exception as e:
            logger.error("Module-level auto-start failed: %s", e)

threading.Thread(target=_module_autostart, daemon=True,
                 name="mcp-module-autostart").start()
