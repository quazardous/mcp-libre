# Troubleshooting Guide

## Extension Not Loading

**Symptom**: No "MCP Server" menu in LibreOffice after install.

**Solutions**:
1. Check the extension is registered: `Tools > Extension Manager` — look for "LibreOffice MCP Server Extension"
2. Reinstall: `.\scripts\install-plugin.ps1 -Force` (Windows) or `./scripts/install-plugin.sh` (Linux)
3. Check the log file: `~/mcp-extension.log`

## Server Not Responding

**Symptom**: `curl http://localhost:8765/health` fails or times out.

**Solutions**:
1. Start the server from the LibreOffice menu: `MCP Server > Start/Stop`
2. Check if the port is already in use: `netstat -an | grep 8765` (or `findstr` on Windows)
3. Check `~/mcp-extension.log` for errors
4. Restart LibreOffice completely

## "LibreOffice is busy" Error

**Symptom**: Tool calls return `"LibreOffice is busy processing another tool call"`.

**Cause**: The extension processes one tool call at a time (backpressure). A long-running operation is blocking.

**Solutions**:
1. Wait a moment and retry — the message includes `"retryable": true`
2. If stuck, restart the server via `MCP Server > Restart`

## LibreOffice Executable Not Found

**Symptom**: Install scripts can't find LibreOffice.

**Solutions**:
- **Linux**: `sudo apt install libreoffice` (or your distro's package manager)
- **Windows**: Install from [libreoffice.org](https://www.libreoffice.org/download/)
- Verify: `libreoffice --version`

## Java Warnings

**Symptom**: `Warning: failed to launch javaldx`

**Solution**: Install Java (optional, only needed for some features):
```bash
# Ubuntu/Debian
sudo apt install default-jre

# Verify
java -version
```

## Permission Denied

**Symptom**: Cannot read or write document files.

**Solutions**:
1. Check file permissions: `ls -la /path/to/document.odt`
2. Check directory permissions: `ls -la /path/to/directory/`
3. Ensure LibreOffice has write access to the target directory

## Debug Mode

For detailed logging, launch LibreOffice with SAL_LOG:

```powershell
# Windows
.\scripts\launch-lo-debug.ps1 -Full
```

Or check the extension log:
- **Plugin log**: `~/mcp-extension.log`

## Common Fixes Summary

| Error | Quick Fix |
|-------|-----------|
| No MCP Server menu | Reinstall extension, restart LO |
| Server not responding | Start from menu, check port 8765 |
| "LibreOffice is busy" | Wait and retry, or restart server |
| LO not found | Install LibreOffice, check PATH |
| Permission denied | Fix file/directory permissions |
| Java warnings | Install Java (optional) |
