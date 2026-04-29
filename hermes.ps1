#!/usr/bin/env pwsh
#Requires -Version 5.1
<#
.SYNOPSIS
    Hermes Agent Windows native launcher.
.DESCRIPTION
    Launches hermes-agent on Windows without WSL/Docker.
    Automatically sets UTF-8 encoding and routes to the project venv.
.EXAMPLE
    .\hermes.ps1                    # Interactive chat
    .\hermes.ps1 doctor             # Run diagnostics
    .\hermes.ps1 setup              # Setup wizard
    .\hermes.ps1 -q "hello"         # Single query
    .\hermes.ps1 --list-tools       # List tools
    .\hermes.ps1 --tui              # Try TUI mode
#>

$ErrorActionPreference = "Stop"

# --- UTF-8 enforcement for Windows console ----------------------------------
# Rich / prompt_toolkit output contains Braille patterns and box-drawing chars
# that fail under the default GBK (936) code page.
$env:PYTHONIOENCODING = "utf-8"

# Legacy: force Git Bash (not WSL bash) on Windows.
# As of Phase 1 Windows-native refactor, Hermes now uses PowerShell
# directly via WindowsLocalEnvironment.  The Git Bash fallback is
# still available via HERMES_USE_GIT_BASH=1 if needed.
# $env:HERMES_GIT_BASH_PATH = "C:\Program Files\Git\bin\bash.exe"

# Also switch the active console code page to UTF-8 (65001) if possible.
# This makes PowerShell-native output safe as well.
$_originalCp = [Console]::OutputEncoding.CodePage
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
} catch {
    # Non-interactive hosts (e.g. some CI) may not allow this; ignore.
}

# --- Locate project root and venv ------------------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$venvPath  = Join-Path $scriptDir ".venv"
$pythonExe = Join-Path $venvPath "Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERROR] Virtual environment not found: $venvPath" -ForegroundColor Red
    Write-Host "        Run the following to create it:" -ForegroundColor Yellow
    Write-Host "        uv venv .venv --python 3.13" -ForegroundColor Cyan
    Write-Host "        uv pip install --python .venv\Scripts\python.exe -e `".[all]`"" -ForegroundColor Cyan
    exit 1
}

# --- Route subcommands ------------------------------------------------------
# If no arguments are given, default to interactive chat (same as `hermes`).
$passArgs = $args
if ($passArgs.Count -eq 0) {
    $passArgs = @("chat")
}

# Detect whether the first argument is a known hermes subcommand or a CLI flag.
# Known subcommands are handled by `hermes_cli.main`; everything else (like `-q`)
# is forwarded to `cli.py` which uses Python Fire.
$mainSubcommands = @(
    "chat", "gateway", "cron", "setup", "status", "doctor", "model", "tools",
    "config", "skills", "sessions", "logs", "auth", "honcho", "claw",
    "version", "update", "uninstall", "acp", "backup", "dump", "plugins"
)

$first = $passArgs[0]
if ($first -in $mainSubcommands) {
    $module = "hermes_cli.main"
} else {
    # Fallback to cli.py for Fire-style invocations (`-q`, `--list-tools`, etc.)
    $module = "cli"
}

# --- Launch -----------------------------------------------------------------
$pythonArgs = @("-m", $module) + $passArgs

# For TUI mode, route through the dedicated TUI entry if requested.
if ($first -eq "--tui") {
    $pythonArgs = @("-m", "tui_gateway.entry")
}

Write-Host "[Hermes] launching: $pythonExe $([string]::Join(' ', $pythonArgs))" -ForegroundColor DarkGray
& $pythonExe @pythonArgs

$exitCode = $LASTEXITCODE

# --- Restore console code page ----------------------------------------------
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding($_originalCp)
} catch { }

exit $exitCode
