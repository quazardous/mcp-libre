#Requires -Version 5.1
<#
.SYNOPSIS
    Build and install the LibreOffice MCP plugin extension (.oxt).

.DESCRIPTION
    Packages the plugin/ directory into an .oxt file and installs it
    into LibreOffice via unopkg. Designed for development iteration.
    Handles lock files, ghost installs, and non-interactive mode.

.EXAMPLE
    .\install-plugin.ps1                   # Build + install (interactive)
    .\install-plugin.ps1 -Force            # Build + install (no prompts, kills LO)
    .\install-plugin.ps1 -BuildOnly        # Only create the .oxt
    .\install-plugin.ps1 -Uninstall        # Remove the extension
    .\install-plugin.ps1 -Uninstall -Force # Remove without prompts
    .\install-plugin.ps1 -Cache            # Hot-deploy to LO cache (dev iteration)
#>

[CmdletBinding()]
param(
    [switch]$BuildOnly,
    [switch]$Uninstall,
    [switch]$NoRestart,
    [switch]$Force,
    [switch]$Cache
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$global:LASTEXITCODE = 0

# ── Paths ────────────────────────────────────────────────────────────────────

$ProjectRoot = Split-Path $PSScriptRoot
$PluginDir   = Join-Path $ProjectRoot "plugin"
$BuildDir    = Join-Path $ProjectRoot "build"
$StagingDir  = Join-Path $BuildDir "staging"
$OxtFile     = Join-Path $BuildDir "libreoffice-mcp-extension.oxt"

$ExtensionId = "org.mcp.libreoffice.extension"

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Step  { param([string]$T) Write-Host "[*] $T" -ForegroundColor White }
function Write-OK    { param([string]$T) Write-Host "[OK] $T" -ForegroundColor Green }
function Write-Warn  { param([string]$T) Write-Host "[!!] $T" -ForegroundColor Yellow }
function Write-Err   { param([string]$T) Write-Host "[X] $T" -ForegroundColor Red }
function Write-Info  { param([string]$T) Write-Host "    $T" -ForegroundColor Gray }

function Write-Utf8NoBom {
    <# Write text to a file as UTF-8 without BOM (PowerShell 5 adds BOM). #>
    param([string]$Path, [string]$Content)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8)
}

function Confirm-OrForce {
    <# Returns $true if -Force or user says yes. #>
    param([string]$Prompt)
    if ($Force) { return $true }
    $response = Read-Host "$Prompt (Y/n)"
    return ($response -eq '' -or $response -match '^[Yy]')
}

# ── LibreOffice detection ────────────────────────────────────────────────────

function Find-LibreOffice {
    $searchPaths = @(
        "${env:ProgramFiles}\LibreOffice\program",
        "${env:ProgramFiles(x86)}\LibreOffice\program",
        "${env:LOCALAPPDATA}\Programs\LibreOffice\program",
        "C:\Program Files\LibreOffice\program"
    )
    foreach ($p in $searchPaths) {
        if (Test-Path (Join-Path $p "soffice.exe")) { return $p }
    }
    $existing = Get-Command soffice -ErrorAction SilentlyContinue
    if ($existing) { return (Split-Path $existing.Source) }
    return $null
}

function Find-Unopkg {
    param([string]$LODir)
    if ($LODir) {
        $candidate = Join-Path $LODir "unopkg.exe"
        if (Test-Path $candidate) { return $candidate }
    }
    $existing = Get-Command unopkg -ErrorAction SilentlyContinue
    if ($existing) { return $existing.Source }
    return $null
}

# ── Process management ───────────────────────────────────────────────────────

function Test-LibreOfficeRunning {
    $procs = Get-Process -Name "soffice*" -ErrorAction SilentlyContinue
    return ($null -ne $procs -and $procs.Count -gt 0)
}

function Stop-LibreOffice {
    <# Forcefully stop all LibreOffice processes, with retries. #>
    Write-Step "Closing LibreOffice..."

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        Get-Process -Name "soffice*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        if (-not (Test-LibreOfficeRunning)) {
            Write-OK "LibreOffice closed"
            return $true
        }
        Write-Info "Attempt $attempt/3 - processes still running, retrying..."
        Start-Sleep -Seconds 2
    }

    if (Test-LibreOfficeRunning) {
        Write-Err "Could not close LibreOffice after 3 attempts"
        Write-Info "Close it manually (Task Manager > soffice.bin) and re-run this script."
        return $false
    }
    Write-OK "LibreOffice closed"
    return $true
}

function Ensure-LibreOfficeStopped {
    <# Make sure LO is not running. Returns $true if safe to proceed. #>
    if (-not (Test-LibreOfficeRunning)) { return $true }

    Write-Warn "LibreOffice is running. It must be closed for unopkg."
    if (-not (Confirm-OrForce "Close LibreOffice now?")) {
        Write-Err "Cannot proceed while LibreOffice is running."
        return $false
    }
    return (Stop-LibreOffice)
}

# ── Lock file / cache management ────────────────────────────────────────────

function Get-UnopkgCacheDir {
    # Standard location: %APPDATA%\LibreOffice\4\user\uno_packages
    $candidate = Join-Path $env:APPDATA "LibreOffice\4\user\uno_packages"
    if (Test-Path $candidate) { return $candidate }
    # Fallback: search
    $found = Get-ChildItem "$env:APPDATA\LibreOffice" -Recurse -Directory -Filter "uno_packages" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { return $found.FullName }
    return $null
}

function Remove-UnopkgLock {
    <# Remove stale lock files left by crashed unopkg processes. #>
    $cacheDir = Get-UnopkgCacheDir
    if (-not $cacheDir) { return }

    # Look for .lock files in the cache tree
    $lockFiles = Get-ChildItem $cacheDir -Recurse -Filter "*.lock" -ErrorAction SilentlyContinue
    foreach ($lock in $lockFiles) {
        try {
            Remove-Item $lock.FullName -Force
            Write-Info "Removed stale lock: $($lock.FullName)"
        } catch {
            Write-Warn "Could not remove lock: $($lock.FullName)"
        }
    }

    # Also check for the unopkg bridge lock in the user profile
    $userDir = Split-Path $cacheDir
    $bridgeLock = Join-Path $userDir ".lock"
    # Note: we do NOT delete the user .lock — that's LibreOffice's own lock.
}

function Clear-GhostExtension {
    <#
    .SYNOPSIS
        Clean up a ghost extension that shows "already added" but not in unopkg list.
        This happens when a previous install failed mid-way (e.g. BOM error, lock file).
    #>
    param([string]$Unopkg)

    $cacheDir = Get-UnopkgCacheDir
    if (-not $cacheDir) { return }

    $packagesDir = Join-Path $cacheDir "cache\uno_packages"
    if (-not (Test-Path $packagesDir)) { return }

    # Look for directories containing our extension
    $ghostDirs = Get-ChildItem $packagesDir -Directory -ErrorAction SilentlyContinue |
        Where-Object {
            (Test-Path (Join-Path $_.FullName "libreoffice-mcp-extension.oxt")) -or
            (Test-Path (Join-Path $_.FullName "description.xml"))
        }

    foreach ($dir in $ghostDirs) {
        Write-Info "Cleaning ghost install: $($dir.Name)"
        try {
            Remove-Item $dir.FullName -Recurse -Force
        } catch {
            Write-Warn "Could not clean $($dir.FullName): $_"
        }
    }

    # Also clean registry backenddb entries that reference our extension
    $bundleDb = Join-Path $cacheDir "cache\registry\com.sun.star.comp.deployment.bundle.PackageRegistryBackend\backenddb.xml"
    if (Test-Path $bundleDb) {
        try {
            $dbContent = Get-Content $bundleDb -Raw -Encoding UTF8
            if ($dbContent -match $ExtensionId -or $dbContent -match "libreoffice-mcp-extension") {
                # Reset the backenddb to empty state
                $cleanDb = @'
<?xml version="1.0"?>
<plugin-backend-db xmlns="http://openoffice.org/extensionmanager/backend-db/2010"/>
'@
                Write-Utf8NoBom $bundleDb $cleanDb
                Write-Info "Reset bundle registry (had ghost references)"
            }
        } catch {
            Write-Warn "Could not clean bundle registry: $_"
        }
    }
}

# ── unopkg wrapper with retries ──────────────────────────────────────────────

function Invoke-Unopkg {
    <#
    .SYNOPSIS
        Run unopkg with retries and lock-file recovery.
        Returns @{ Success = $bool; Output = $string; ExitCode = $int }
    #>
    param(
        [string]$Exe,
        [string[]]$Args,
        [int]$MaxRetries = 2,
        [int]$DelaySec = 3
    )

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        $output = & $Exe @Args 2>&1
        $exit = $LASTEXITCODE
        $outStr = ($output | Out-String).Trim()

        # Success
        if ($exit -eq 0) {
            return @{ Success = $true; Output = $outStr; ExitCode = 0 }
        }

        # Lock file error → clean and retry
        if ($outStr -match "fichier verrou|lock file|is locked|en cours d.ex") {
            if ($attempt -le $MaxRetries) {
                Write-Warn "Lock file detected (attempt $attempt/$($MaxRetries + 1)). Cleaning up..."
                Remove-UnopkgLock
                Start-Sleep -Seconds $DelaySec
                continue
            }
        }

        # "Already added" → ghost install, clean and retry
        if ($outStr -match "d.j. ajout|already added|already installed") {
            if ($attempt -le $MaxRetries) {
                Write-Warn "Ghost extension detected (attempt $attempt/$($MaxRetries + 1)). Cleaning cache..."
                Clear-GhostExtension $Exe
                # Also try a remove first
                $null = & $Exe remove $ExtensionId 2>&1
                Start-Sleep -Seconds $DelaySec
                continue
            }
        }

        # "Activation error" on `add` → extension IS installed, just registration
        # failed (LO Python can't resolve all imports in unopkg context). Treat as OK.
        if ($outStr -match "activation|l.activation") {
            Write-Warn "Activation warning (harmless - extension loads at LO startup)"
            return @{ Success = $true; Output = $outStr; ExitCode = $exit }
        }

        # Other error on last attempt → give up
        if ($attempt -gt $MaxRetries) {
            return @{ Success = $false; Output = $outStr; ExitCode = $exit }
        }

        Write-Info "Retrying in ${DelaySec}s (attempt $attempt/$($MaxRetries + 1))..."
        Start-Sleep -Seconds $DelaySec
    }

    return @{ Success = $false; Output = "Max retries exceeded"; ExitCode = -1 }
}

# ── Build ────────────────────────────────────────────────────────────────────

function Fix-PythonImports {
    <#
    .SYNOPSIS
        Convert relative imports to absolute ONLY for .py files at the root of
        the given directory (not in sub-packages). LO adds pythonpath/ to sys.path
        directly, so root-level "from .mcp_server import X" must become
        "from mcp_server import X". Sub-packages (services/, tools/) keep their
        relative imports intact since Python resolves them normally.
        All files are re-written as UTF-8 without BOM (LO Python chokes on BOM).
    #>
    param([string]$Directory)

    # Root-level .py files: strip BOM + fix relative imports
    Get-ChildItem $Directory -Filter "*.py" -File | ForEach-Object {
        $content = Get-Content $_.FullName -Raw -Encoding UTF8
        if (-not $content) { return }
        $content = $content -replace '^\xEF\xBB\xBF', ''
        $content = $content.TrimStart([char]0xFEFF)
        $original = $content

        # from .module import X  -->  from module import X
        $content = $content -replace 'from \.([\w]+) import', 'from $1 import'
        # import .module  -->  import module
        $content = $content -replace 'import \.([\w]+)', 'import $1'

        Write-Utf8NoBom $_.FullName $content
        if ($content -ne $original) {
            Write-Info "Fixed imports in $($_.Name)"
        }
    }

    # Sub-package .py files: strip BOM only (keep relative imports)
    Get-ChildItem $Directory -Filter "*.py" -Recurse -File |
        Where-Object { $_.DirectoryName -ne (Resolve-Path $Directory).Path } |
        ForEach-Object {
            $content = Get-Content $_.FullName -Raw -Encoding UTF8
            if (-not $content) { return }
            $original = $content
            $content = $content -replace '^\xEF\xBB\xBF', ''
            $content = $content.TrimStart([char]0xFEFF)
            Write-Utf8NoBom $_.FullName $content
            if ($content -ne $original) {
                Write-Info "Fixed BOM in $($_.Name)"
            }
        }
}

function Build-Oxt {
    Write-Host ""
    Write-Host "=== Building .oxt ===" -ForegroundColor Cyan
    Write-Host ""

    # Validate source files exist
    $requiredFiles = @(
        "pythonpath\registration.py",
        "pythonpath\mcp_server.py",
        "pythonpath\ai_interface.py",
        "pythonpath\main_thread_executor.py",
        "pythonpath\version.py",
        "pythonpath\services\__init__.py",
        "pythonpath\services\base.py",
        "pythonpath\tools\__init__.py",
        "pythonpath\tools\base.py",
        "Addons.xcu",
        "ProtocolHandler.xcu",
        "MCPServerConfig.xcs",
        "MCPServerConfig.xcu",
        "OptionsDialog.xcu",
        "dialogs\MCPSettings.xdl",
        "dialogs\MCPTunnel.xdl",
        "icons\stopped_16.png",
        "icons\running_16.png",
        "icons\starting_16.png"
    )
    $missing = $requiredFiles | Where-Object { -not (Test-Path (Join-Path $PluginDir $_)) }
    if ($missing) {
        Write-Err "Missing source files in plugin/:"
        $missing | ForEach-Object { Write-Info "  - $_" }
        return $false
    }

    # Clean previous build
    if (Test-Path $StagingDir) { Remove-Item $StagingDir -Recurse -Force }
    if (Test-Path $OxtFile)    { Remove-Item $OxtFile -Force }
    New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

    # ── Copy plugin sources ──────────────────────────────────────────────
    Write-Step "Copying plugin files to staging..."

    # pythonpath/ — copy entire tree (LO adds this dir to sys.path)
    $stagePy = Join-Path $StagingDir "pythonpath"
    Copy-Item (Join-Path $PluginDir "pythonpath") $stagePy -Recurse
    # Remove __pycache__ dirs and registration.py (goes to root)
    Get-ChildItem $stagePy -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force }
    $stageReg = Join-Path $stagePy "registration.py"
    if (Test-Path $stageReg) { Remove-Item $stageReg -Force }

    # registration.py → extension root (UNO component entry point)
    Copy-Item (Join-Path $PluginDir "pythonpath\registration.py") $StagingDir

    # Fix relative imports → absolute in all staged .py files (recursive)
    Write-Step "Fixing Python imports for LibreOffice extension structure..."
    Fix-PythonImports $stagePy
    Fix-PythonImports $StagingDir

    # META-INF/
    $stageMeta = Join-Path $StagingDir "META-INF"
    New-Item -ItemType Directory -Path $stageMeta -Force | Out-Null

    $manifestXml = @'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.uno-component;type=Python" manifest:full-path="registration.py"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-schema" manifest:full-path="MCPServerConfig.xcs"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="MCPServerConfig.xcu"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="Addons.xcu"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="ProtocolHandler.xcu"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="OptionsDialog.xcu"/>
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="Jobs.xcu"/>
</manifest:manifest>
'@
    Write-Utf8NoBom (Join-Path $stageMeta "manifest.xml") $manifestXml
    Write-Info "manifest.xml generated"

    # description.xml — copy from plugin/, patch version from version.py (single source of truth)
    Copy-Item (Join-Path $PluginDir "description.xml") $StagingDir
    $versionPy = Get-Content (Join-Path $PluginDir "pythonpath\version.py") -Raw
    if ($versionPy -match 'EXTENSION_VERSION\s*=\s*"([^"]+)"') {
        $version = $Matches[1]
        $descPath = Join-Path $StagingDir "description.xml"
        $descContent = Get-Content $descPath -Raw -Encoding UTF8
        $descContent = $descContent -replace '<version value="[^"]*"/>', "<version value=`"$version`"/>"
        Write-Utf8NoBom $descPath $descContent
        Write-Info "description.xml patched with version $version"
    } else {
        Write-Warn "Could not read version from version.py"
    }

    # XCU/XCS config files
    Copy-Item (Join-Path $PluginDir "Addons.xcu") $StagingDir
    Copy-Item (Join-Path $PluginDir "ProtocolHandler.xcu") $StagingDir
    Copy-Item (Join-Path $PluginDir "MCPServerConfig.xcs") $StagingDir
    Copy-Item (Join-Path $PluginDir "MCPServerConfig.xcu") $StagingDir
    Copy-Item (Join-Path $PluginDir "OptionsDialog.xcu") $StagingDir
    Copy-Item (Join-Path $PluginDir "Jobs.xcu") $StagingDir

    # Dialogs (options page layout)
    $stageDialogs = Join-Path $StagingDir "dialogs"
    New-Item -ItemType Directory -Path $stageDialogs -Force | Out-Null
    Copy-Item (Join-Path $PluginDir "dialogs\MCPSettings.xdl") $stageDialogs
    Copy-Item (Join-Path $PluginDir "dialogs\MCPTunnel.xdl") $stageDialogs

    # Icons (dynamic menu state icons)
    $stageIcons = Join-Path $StagingDir "icons"
    New-Item -ItemType Directory -Path $stageIcons -Force | Out-Null
    Copy-Item (Join-Path $PluginDir "icons\*.png") $stageIcons
    Write-Info "Copied $((@(Get-ChildItem $stageIcons -Filter '*.png')).Count) icon files"

    # Text files
    foreach ($f in @("description-en.txt", "release-notes-en.txt")) {
        $src = Join-Path $PluginDir $f
        if (Test-Path $src) { Copy-Item $src $StagingDir }
    }

    # LICENSE from project root
    $licenseFile = Join-Path $ProjectRoot "LICENSE"
    if (Test-Path $licenseFile) {
        Copy-Item $licenseFile $StagingDir
        Write-Info "Included LICENSE"
    }

    # ── Create .oxt (ZIP) ────────────────────────────────────────────────
    Write-Step "Creating .oxt package..."
    New-Item -ItemType Directory -Path $BuildDir -Force | Out-Null
    # Compress-Archive uses backslashes in entry names — Java/LO rejects those.
    # Use .NET ZipFile API which uses forward slashes.
    if (Test-Path $OxtFile) { Remove-Item $OxtFile -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($OxtFile, 'Create')
    try {
        $basePath = (Resolve-Path $StagingDir).Path.TrimEnd('\') + '\'
        Get-ChildItem -Path $StagingDir -Recurse -File | ForEach-Object {
            $entryName = $_.FullName.Substring($basePath.Length).Replace('\', '/')
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip, $_.FullName, $entryName, 'Optimal') | Out-Null
        }
    } finally {
        $zip.Dispose()
    }

    if (Test-Path $OxtFile) {
        $size = (Get-Item $OxtFile).Length
        Write-OK "Built: $OxtFile ($size bytes)"
    } else {
        Write-Err "Failed to create .oxt file"
        return $false
    }

    # Clean staging
    Remove-Item $StagingDir -Recurse -Force
    return $true
}

# ── Install / Uninstall ─────────────────────────────────────────────────────

function Install-Extension {
    param([string]$Unopkg)

    Write-Host ""
    Write-Host "=== Installing Extension ===" -ForegroundColor Cyan
    Write-Host ""

    # ── 1. Ensure LibreOffice is stopped ─────────────────────────────────
    if (-not (Ensure-LibreOfficeStopped)) { return $false }

    # ── 2. Clean lock files preemptively ─────────────────────────────────
    Remove-UnopkgLock

    # ── 3. Remove previous version ───────────────────────────────────────
    Write-Step "Removing previous version (if any)..."
    $removeResult = Invoke-Unopkg $Unopkg @("remove", $ExtensionId)
    if ($removeResult.Success) {
        Write-Info "Previous version removed"
    } else {
        $out = $removeResult.Output
        if ($out -match "pas d.ploy|not deployed|no such|aucune extension") {
            Write-Info "No previous version found (OK)"
        } else {
            Write-Warn "Remove returned: $out"
            Write-Info "Attempting cache cleanup..."
            Clear-GhostExtension $Unopkg
        }
    }

    # Small delay between remove and add to avoid lock contention
    Start-Sleep -Seconds 2

    # ── 4. Install new version ───────────────────────────────────────────
    Write-Step "Installing $OxtFile ..."
    $installResult = Invoke-Unopkg $Unopkg @("add", $OxtFile) -MaxRetries 3 -DelaySec 3
    if (-not $installResult.Success) {
        Write-Err "unopkg add failed after retries"
        Write-Info "Output: $($installResult.Output)"
        Write-Host ""
        Write-Info "Troubleshooting:"
        Write-Info "  1. Make sure LibreOffice is fully closed (check Task Manager)"
        Write-Info "  2. Try: .\install-plugin.ps1 -Uninstall -Force"
        Write-Info "  3. Then: .\install-plugin.ps1 -Force"
        Write-Info "  4. If still failing, delete the cache manually:"

        $cacheDir = Get-UnopkgCacheDir
        if ($cacheDir) {
            Write-Info "     Remove-Item '$cacheDir\cache' -Recurse -Force"
        }
        return $false
    }

    Write-OK "Extension installed successfully!"

    # ── 5. Verify by checking the cache directly ────────────────────────
    Write-Step "Verifying installation..."
    $cacheDir = Get-UnopkgCacheDir
    $verified = $false
    if ($cacheDir) {
        $cacheVersionFiles = Get-ChildItem $cacheDir -Recurse -Filter "version.py" -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "libreoffice-mcp" }
        foreach ($vf in $cacheVersionFiles) {
            $vContent = Get-Content $vf.FullName -Raw -ErrorAction SilentlyContinue
            if ($vContent -match 'EXTENSION_VERSION\s*=\s*"([^"]+)"') {
                $cachedVersion = $Matches[1]
                Write-OK "Extension verified in cache: v$cachedVersion"
                $verified = $true
                break
            }
        }
    }
    if (-not $verified) {
        Write-Warn "Could not verify cached version (LO will load it on start)"
    }

    return $true
}

function Uninstall-Extension {
    param([string]$Unopkg)

    Write-Host ""
    Write-Host "=== Uninstalling Extension ===" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Ensure-LibreOfficeStopped)) { return $false }

    Remove-UnopkgLock

    Write-Step "Removing extension $ExtensionId ..."
    $result = Invoke-Unopkg $Unopkg @("remove", $ExtensionId) -MaxRetries 2
    if ($result.Success) {
        Write-OK "Extension removed"
    } else {
        $out = $result.Output
        if ($out -match "pas d.ploy|not deployed|no such|aucune extension") {
            Write-Info "Extension was not installed"
        } else {
            Write-Warn "unopkg remove: $out"
            Write-Info "Cleaning cache as fallback..."
            Clear-GhostExtension $Unopkg
            Write-OK "Cache cleaned"
        }
    }
    return $true
}

function Start-LibreOfficeApp {
    param([string]$LODir)
    $soffice = Join-Path $LODir "soffice.exe"
    Write-Step "Starting LibreOffice..."
    Start-Process $soffice
    Write-OK "LibreOffice started"
    Write-Info "Check menu bar for 'MCP Server' entry"
    Write-Info "Click 'Start MCP Server', then test with:"
    Write-Info "  curl -k https://localhost:8765/health"
}

# ── Cache install (hot-deploy for dev iteration) ─────────────────────────────

function Install-ToCache {
    <#
    .SYNOPSIS
        Copy source files directly into the LO extension cache, converting
        relative imports and stripping BOM. Much faster than unopkg for
        development iteration — just restart LO afterwards.

        Requires that the extension was already installed via unopkg at least
        once (so the cache directory structure exists).
    #>

    Write-Host ""
    Write-Host "=== Cache Install (hot-deploy) ===" -ForegroundColor Cyan
    Write-Host ""

    # ── Find the cached extension directory ──────────────────────────────
    $cacheDir = Get-UnopkgCacheDir
    if (-not $cacheDir) {
        Write-Err "Could not find uno_packages cache directory"
        exit 1
    }

    $packagesDir = Join-Path $cacheDir "cache\uno_packages"
    if (-not (Test-Path $packagesDir)) {
        Write-Err "Cache packages dir not found: $packagesDir"
        Write-Info "Run a normal install first: .\install-plugin.ps1 -Force"
        exit 1
    }

    # Find the *.tmp_ directory containing our extension
    $extDir = $null
    Get-ChildItem $packagesDir -Directory -Filter "*.tmp_" -ErrorAction SilentlyContinue | ForEach-Object {
        $oxtDir = Join-Path $_.FullName "libreoffice-mcp-extension.oxt"
        if (Test-Path $oxtDir) {
            $extDir = $oxtDir
        }
    }
    if (-not $extDir) {
        Write-Err "Extension not found in cache. Run a normal install first."
        exit 1
    }
    Write-OK "Cache dir: $extDir"

    # ── Copy and fix files ───────────────────────────────────────────────

    # Track what we deploy
    $deployed = @()

    # --- pythonpath/ tree (root-level .py: fix imports; sub-packages: BOM only) ---
    $srcPy  = Join-Path $PluginDir "pythonpath"
    $dstPy  = Join-Path $extDir    "pythonpath"

    if (-not (Test-Path $dstPy)) {
        New-Item -ItemType Directory -Path $dstPy -Force | Out-Null
    }

    # Root-level .py files: fix relative imports + strip BOM
    Get-ChildItem $srcPy -Filter "*.py" -File | ForEach-Object {
        $content = Get-Content $_.FullName -Raw -Encoding UTF8
        if (-not $content) { return }
        $content = $content -replace '^\xEF\xBB\xBF', ''
        $content = $content.TrimStart([char]0xFEFF)
        # from .module import X  -->  from module import X
        $content = $content -replace 'from \.([\w]+) import', 'from $1 import'
        # import .module  -->  import module
        $content = $content -replace 'import \.([\w]+)', 'import $1'

        $dst = Join-Path $dstPy $_.Name
        Write-Utf8NoBom $dst $content
        $deployed += "pythonpath/$($_.Name)"
    }

    # Sub-package .py files: strip BOM only, keep relative imports
    Get-ChildItem $srcPy -Filter "*.py" -Recurse -File |
        Where-Object { $_.DirectoryName -ne (Resolve-Path $srcPy).Path } |
        ForEach-Object {
            $relPath = $_.FullName.Substring((Resolve-Path $srcPy).Path.Length + 1)
            $dst = Join-Path $dstPy $relPath
            $dstDir = Split-Path $dst
            if (-not (Test-Path $dstDir)) {
                New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            }
            $content = Get-Content $_.FullName -Raw -Encoding UTF8
            if (-not $content) { return }
            $content = $content -replace '^\xEF\xBB\xBF', ''
            $content = $content.TrimStart([char]0xFEFF)
            Write-Utf8NoBom $dst $content
            $deployed += "pythonpath/$relPath"
        }

    # --- registration.py → extension root (UNO entry point) ---
    $regSrc = Join-Path $PluginDir "pythonpath\registration.py"
    $regDst = Join-Path $extDir "registration.py"
    $content = Get-Content $regSrc -Raw -Encoding UTF8
    $content = $content -replace '^\xEF\xBB\xBF', ''
    $content = $content.TrimStart([char]0xFEFF)
    $content = $content -replace 'from \.([\w]+) import', 'from $1 import'
    $content = $content -replace 'import \.([\w]+)', 'import $1'
    Write-Utf8NoBom $regDst $content
    $deployed += "registration.py"

    # --- Config / XCU / XCS files ---
    $configFiles = @(
        "MCPServerConfig.xcs",
        "MCPServerConfig.xcu",
        "Addons.xcu",
        "ProtocolHandler.xcu",
        "OptionsDialog.xcu",
        "Jobs.xcu",
        "description.xml"
    )
    foreach ($f in $configFiles) {
        $src = Join-Path $PluginDir $f
        if (Test-Path $src) {
            Copy-Item $src (Join-Path $extDir $f) -Force
            $deployed += $f
        }
    }

    # --- Dialogs ---
    $dstDialogs = Join-Path $extDir "dialogs"
    if (-not (Test-Path $dstDialogs)) {
        New-Item -ItemType Directory -Path $dstDialogs -Force | Out-Null
    }
    Get-ChildItem (Join-Path $PluginDir "dialogs") -Filter "*.xdl" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName (Join-Path $dstDialogs $_.Name) -Force
        $deployed += "dialogs/$($_.Name)"
    }

    # --- Icons ---
    $dstIcons = Join-Path $extDir "icons"
    if (-not (Test-Path $dstIcons)) {
        New-Item -ItemType Directory -Path $dstIcons -Force | Out-Null
    }
    Get-ChildItem (Join-Path $PluginDir "icons") -Filter "*.png" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName (Join-Path $dstIcons $_.Name) -Force
        $deployed += "icons/$($_.Name)"
    }

    # --- AGENT.md ---
    $agentMd = Join-Path $ProjectRoot "AGENT.md"
    if (Test-Path $agentMd) {
        Copy-Item $agentMd (Join-Path $extDir "AGENT.md") -Force
        $deployed += "AGENT.md"
    }

    # ── Report ───────────────────────────────────────────────────────────
    Write-Host ""
    Write-OK "Deployed $($deployed.Count) files to cache"
    $deployed | ForEach-Object { Write-Info "  $_" }
    Write-Host ""
    Write-Info "Restart LibreOffice to pick up changes."
    Write-Host ""
}

# ── Main ─────────────────────────────────────────────────────────────────────

function Main {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  LibreOffice MCP Plugin Installer" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Find LibreOffice
    $loDir = Find-LibreOffice
    if (-not $loDir) {
        Write-Err "LibreOffice not found. Install it first (.\install.ps1)"
        exit 1
    }
    Write-OK "LibreOffice: $loDir"

    # Cache mode (hot-deploy without unopkg)
    if ($Cache) {
        Install-ToCache
        return
    }

    # Find unopkg
    $unopkg = Find-Unopkg $loDir
    if (-not $unopkg) {
        Write-Err "unopkg.exe not found in LibreOffice directory"
        exit 1
    }
    Write-OK "unopkg: $unopkg"

    # Uninstall mode
    if ($Uninstall) {
        Uninstall-Extension $unopkg
        return
    }

    # Build
    $buildOk = Build-Oxt
    if (-not $buildOk) { exit 1 }

    if ($BuildOnly) {
        Write-Host ""
        Write-OK "Build complete. Install manually with:"
        Write-Info "  & `"$unopkg`" add `"$OxtFile`""
        return
    }

    # Install
    $installOk = Install-Extension $unopkg
    if (-not $installOk) { exit 1 }

    # Restart LibreOffice
    if (-not $NoRestart) {
        Write-Host ""
        if (Confirm-OrForce "Start LibreOffice now?") {
            Start-LibreOfficeApp $loDir
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Done!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Next steps:" -ForegroundColor White
    Write-Host "  1. Open a document in LibreOffice" -ForegroundColor Gray
    Write-Host "  2. MCP Server > Start MCP Server  (in the menu bar)" -ForegroundColor Gray
    Write-Host "  3. Test: curl -k https://localhost:8765/health" -ForegroundColor Gray
    Write-Host ""
}

Main
