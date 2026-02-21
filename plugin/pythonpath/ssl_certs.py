"""Auto-generate and load self-signed TLS certificates."""

import os
import ssl
import logging
import subprocess

logger = logging.getLogger("mcp-extension")


def get_cert_dir():
    """Return the directory where TLS certificates are stored.

    Windows: %APPDATA%/mcp-certs/
    Linux:   ~/.config/libreoffice/mcp-certs/
    """
    if os.name == "nt":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.path.join(os.path.expanduser("~"), ".config", "libreoffice")
    return os.path.join(base, "mcp-certs")


def ensure_certs():
    """Generate cert + key if not present. Returns (cert_path, key_path)."""
    cert_dir = get_cert_dir()
    os.makedirs(cert_dir, exist_ok=True)
    cert_path = os.path.join(cert_dir, "server.pem")
    key_path = os.path.join(cert_dir, "server-key.pem")
    if os.path.exists(cert_path) and os.path.exists(key_path):
        logger.info("TLS certificates found at %s", cert_dir)
        return cert_path, key_path
    _generate_self_signed(cert_path, key_path)
    return cert_path, key_path


def _generate_self_signed(cert_path, key_path):
    """Generate a self-signed certificate using openssl CLI."""
    cmd = [
        "openssl", "req", "-x509", "-newkey", "rsa:2048",
        "-keyout", key_path, "-out", cert_path,
        "-days", "3650", "-nodes",
        "-subj", "/CN=localhost",
        "-addext", "subjectAltName=DNS:localhost,IP:127.0.0.1",
    ]
    try:
        kwargs = {"capture_output": True, "check": True, "timeout": 30}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(
                subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.run(cmd, **kwargs)
        logger.info("Generated self-signed certificate at %s", cert_path)
    except FileNotFoundError:
        raise RuntimeError(
            "openssl not found. Install OpenSSL and ensure it is on PATH."
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(
            f"openssl certificate generation failed: "
            f"{e.stderr.decode('utf-8', errors='replace')}"
        )


def create_ssl_context(cert_path, key_path):
    """Create an SSLContext for the HTTPS server."""
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(cert_path, key_path)
    return ctx
