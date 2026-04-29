@echo off
setlocal EnableDelayedExpansion

:: Hermes Agent Windows native launcher (cmd.exe version)
:: Automatically sets UTF-8 encoding and routes to the project venv.
::
:: Usage:
::   hermes.cmd                Interactive chat
::   hermes.cmd doctor         Run diagnostics
::   hermes.cmd setup          Setup wizard
::   hermes.cmd -q "hello"     Single query
::   hermes.cmd --list-tools   List tools

:: UTF-8 enforcement for Rich / prompt_toolkit output
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul 2>&1

:: Legacy: force Git Bash (not WSL bash) on Windows.
:: As of Phase 1 Windows-native refactor, Hermes now uses PowerShell
:: directly via WindowsLocalEnvironment.  The Git Bash fallback is
:: still available via HERMES_USE_GIT_BASH=1 if needed.
:: set "HERMES_GIT_BASH_PATH=C:\Program Files\Git\bin\bash.exe"

set "SCRIPT_DIR=%~dp0"
set "VENV_PATH=%SCRIPT_DIR%.venv"
set "PYTHON_EXE=%VENV_PATH%\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment not found: %VENV_PATH%
    echo         Run the following to create it:
    echo         uv venv .venv --python 3.13
    echo         uv pip install --python .venv\Scripts\python.exe -e ".[all]"
    exit /b 1
)

:: Determine entry point based on first argument
set "FIRST_ARG=%~1"
set "MODULE=hermes_cli.main"

if "%FIRST_ARG%"=="" set "MODULE=hermes_cli.main" & goto :launch
if "%FIRST_ARG:~0,1%"=="-" set "MODULE=cli" & goto :launch
if "%FIRST_ARG%"=="--tui" set "MODULE=tui_gateway.entry" & goto :launch

:: Known hermes_cli.main subcommands
set "SUBCOMMANDS=chat gateway cron setup status doctor model tools config skills sessions logs auth honcho claw version update uninstall acp backup dump plugins"
echo %SUBCOMMANDS% | findstr /I /C:" %FIRST_ARG% " >nul || set "MODULE=cli"

:launch
"%PYTHON_EXE%" -m %MODULE% %*
exit /b %ERRORLEVEL%
