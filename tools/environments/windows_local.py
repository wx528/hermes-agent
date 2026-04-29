"""Windows-native execution environment using PowerShell.

Phase 2: direct PowerShell spawn with session snapshot (env var persistence).
Env vars set via `$env:FOO = "bar"` or `set FOO=bar` persist across commands.
"""

import logging
import os
import platform
import shutil
import subprocess
import tempfile
from typing import Optional

from tools.environments.base import BaseEnvironment, _pipe_stdin
from tools.environments.local import _sanitize_subprocess_env

logger = logging.getLogger(__name__)

_IS_WINDOWS = platform.system() == "Windows"


def _find_powershell() -> str:
    """Find the best PowerShell executable on Windows.

    Priority:
    1. PowerShell 7 (pwsh.exe) — modern, faster, better pipeline support
    2. PowerShell 5.1 (powershell.exe) — built-in on all Windows versions
    """
    # 1. Check PATH for pwsh (PowerShell 7)
    pwsh = shutil.which("pwsh")
    if pwsh:
        return pwsh

    # 2. Known PowerShell 7 locations
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    for base in (program_files, os.environ.get("ProgramW6432", "")):
        if not base:
            continue
        candidate = os.path.join(base, "PowerShell", "7", "pwsh.exe")
        if os.path.isfile(candidate):
            return candidate

    # 3. Fallback to Windows PowerShell 5.1
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    candidate = os.path.join(
        system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"
    )
    if os.path.isfile(candidate):
        return candidate

    # 4. Last resort — hope it's on PATH
    ps = shutil.which("powershell")
    if ps:
        return ps

    raise RuntimeError(
        "PowerShell not found on Windows. "
        "Please ensure PowerShell is installed and available in PATH."
    )


def _quote_ps_literal(value: str) -> str:
    """Quote a string for PowerShell single-quoted literals.

    In PowerShell single-quoted strings, the only escape is doubling the
    single quote: ' → ''.
    """
    return "'" + value.replace("'", "''") + "'"


# Separator used in the env snapshot file.
# Chosen to be extremely unlikely in env var names or values.
_ENV_SEP = "\x00"


class WindowsLocalEnvironment(BaseEnvironment):
    """Run commands directly on Windows using PowerShell (no bash).

    Session snapshot preserves environment variables across commands.
    CWD persists via file-based read after each command.
    """

    def __init__(self, cwd: str = "", timeout: int = 60, env: dict = None):
        if cwd:
            cwd = os.path.expanduser(cwd)
        super().__init__(cwd=cwd or os.getcwd(), timeout=timeout, env=env)
        self.init_session()

    def get_temp_dir(self) -> str:
        """Return a writable temp dir for local execution on Windows."""
        candidate = (
            self.env.get("TEMP")
            or self.env.get("TMP")
            or os.environ.get("TEMP")
            or os.environ.get("TMP")
            or tempfile.gettempdir()
        )
        # Keep forward slashes for consistency with BaseEnvironment path handling.
        return candidate.replace("\\", "/").rstrip("/") or "/"

    # ------------------------------------------------------------------
    # Session snapshot
    # ------------------------------------------------------------------

    def init_session(self):
        """Capture initial environment into a snapshot file.

        The snapshot is a UTF-8 text file with NUL-separated key-value pairs,
        one per line:  NAME\x00VALUE
        This format handles env values containing newlines, equals signs, etc.
        """
        # Build a PowerShell script that dumps all env vars to the snapshot.
        # We use NUL as separator because it cannot appear in env var names
        # or values on Windows.
        quoted_snap = _quote_ps_literal(self._snapshot_path)
        bootstrap = (
            "$ProgressPreference = 'SilentlyContinue'\n"
            # Dump all env vars as KEY\x00VALUE lines
            "Get-ChildItem Env: | ForEach-Object {\n"
            f"  \"$($_.Name){_ENV_SEP}$($_.Value)\"\n"
            "} | Out-File -FilePath " + quoted_snap + " -Encoding utf8\n"
        )
        try:
            proc = self._run_bash(bootstrap, login=True, timeout=self._snapshot_timeout)
            self._wait_for_process(proc, timeout=self._snapshot_timeout)
            # Verify the snapshot file is readable
            self._load_snapshot()
            self._snapshot_ready = True
            logger.info(
                "PowerShell session snapshot created (session=%s, cwd=%s)",
                self._session_id, self.cwd,
            )
        except Exception as exc:
            logger.warning(
                "init_session failed (session=%s): %s — "
                "env vars will not persist across commands",
                self._session_id, exc,
            )
            self._snapshot_ready = False

    def _load_snapshot(self) -> dict:
        """Load the env snapshot file into a dict."""
        env = {}
        try:
            with open(self._snapshot_path, encoding="utf-8-sig") as f:
                for line in f:
                    line = line.rstrip("\n\r")
                    if not line:
                        continue
                    parts = line.split(_ENV_SEP, 1)
                    if len(parts) == 2:
                        env[parts[0]] = parts[1]
        except (OSError, FileNotFoundError):
            pass
        return env

    def _save_snapshot(self, env: dict) -> None:
        """Write env dict back to the snapshot file (BOM-free UTF-8)."""
        try:
            with open(self._snapshot_path, "w", encoding="utf-8") as f:
                for key, value in env.items():
                    f.write(f"{key}{_ENV_SEP}{value}\n")
        except OSError as exc:
            logger.debug("Failed to save env snapshot: %s", exc)

    # ------------------------------------------------------------------
    # Command wrapping
    # ------------------------------------------------------------------

    def _wrap_command(self, command: str, cwd: str) -> str:
        """Build a PowerShell script that restores env, cd's, runs command,
        re-dumps env vars, saves CWD, and emits CWD marker."""
        quoted_cwd = _quote_ps_literal(cwd)
        quoted_cwd_file = _quote_ps_literal(self._cwd_file)
        quoted_snap = _quote_ps_literal(self._snapshot_path)
        marker = self._cwd_marker

        # --- Restore env from snapshot ---
        restore_env = ""
        if self._snapshot_ready:
            # Read the snapshot file line by line, split on NUL, set $env:
            # Using -Raw + split is faster than Get-Content for large files.
            restore_env = (
                f"if (Test-Path {quoted_snap}) {{\n"
                f"  $__snap = [IO.File]::ReadAllText({quoted_snap}, [System.Text.UTF8Encoding]::new($false))\n"
                f"  $__snap -split \"`n\" | ForEach-Object {{\n"
                f"    $__parts = $_ -split [char]0, 2\n"
                f"    if ($__parts.Length -eq 2 -and $__parts[0]) {{\n"
                f"      Set-Item -Path \"Env:$($__parts[0])\" -Value $__parts[1] -ErrorAction SilentlyContinue\n"
                f"    }}\n"
                f"  }}\n"
                f"}}\n"
            )

        # --- Dump env back to snapshot after command ---
        dump_env = ""
        if self._snapshot_ready:
            # Collect env lines into a variable, then write all at once.
            # Can't pipe directly into [IO.File]::WriteAllLines (PS parser error).
            dump_env = (
                "$__envLines = Get-ChildItem Env: | ForEach-Object {\n"
                f"  \"$($_.Name){_ENV_SEP}$($_.Value)\"\n"
                "}\n"
                f"[IO.File]::WriteAllLines({quoted_snap}, $__envLines, [System.Text.UTF8Encoding]::new($false))\n"
            )

        ps_script = (
            f"$ProgressPreference = 'SilentlyContinue'\n"
            # Override Get-Location / pwd to emit path string to stdout
            f"Remove-Item Alias:pwd -ErrorAction SilentlyContinue\n"
            f"function pwd {{ Write-Output (Microsoft.PowerShell.Management\\Get-Location).Path }}\n"
            f"function Get-Location {{ Write-Output (Microsoft.PowerShell.Management\\Get-Location).Path }}\n"
            # Restore env vars from snapshot
            f"{restore_env}"
            # cd to working directory
            f"Set-Location -LiteralPath {quoted_cwd}\n"
            # Run user command via here-string
            f"$__cmd = @'\n"
            f"{command}\n"
            f"'@\n"
            f"Invoke-Expression $__cmd\n"
            f"$__hermes_ec = $LASTEXITCODE\n"
            f"if ($__hermes_ec -eq $null) {{ $__hermes_ec = 0 }}\n"
            # Save CWD
            f"[IO.File]::WriteAllText({quoted_cwd_file}, (Microsoft.PowerShell.Management\\Get-Location).Path, [System.Text.UTF8Encoding]::new($false))\n"
            # Re-dump env vars to snapshot
            f"{dump_env}"
            # CWD marker
            f'Write-Output "`n{marker}$((Microsoft.PowerShell.Management\\Get-Location).Path){marker}"\n'
            f"exit $__hermes_ec\n"
        )
        return ps_script

    # ------------------------------------------------------------------
    # Process spawning
    # ------------------------------------------------------------------

    def _run_bash(
        self,
        cmd_string: str,
        *,
        login: bool = False,
        timeout: int = 120,
        stdin_data: Optional[str] = None,
    ) -> subprocess.Popen:
        """Spawn PowerShell to run *cmd_string* (a PowerShell script).

        Writes the script to a temporary .ps1 file so we don't have to fight
        PowerShell's command-line quoting rules, then invokes it with
        ``-NoProfile -NonInteractive -ExecutionPolicy Bypass -File <path>``.
        """
        pwsh = _find_powershell()
        run_env = _sanitize_subprocess_env(os.environ, self.env)

        temp_dir = self.get_temp_dir()
        ps1_path = os.path.join(temp_dir, f"hermes-ps-{self._session_id}.ps1")
        with open(ps1_path, "w", encoding="utf-8") as f:
            f.write(cmd_string)

        args = [
            pwsh,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ps1_path,
        ]

        proc = subprocess.Popen(
            args,
            text=True,
            env=run_env,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE if stdin_data is not None else subprocess.DEVNULL,
            cwd=self.cwd.strip("\ufeff"),  # Guard against BOM leaking into cwd
        )

        if stdin_data is not None:
            _pipe_stdin(proc, stdin_data)

        return proc

    # ------------------------------------------------------------------
    # Process lifecycle
    # ------------------------------------------------------------------

    def _kill_process(self, proc):
        """Kill the process and its children on Windows."""
        try:
            proc.terminate()
        except Exception:
            pass
        try:
            subprocess.run(
                ["taskkill", "/T", "/F", "/PID", str(proc.pid)],
                capture_output=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass

    # ------------------------------------------------------------------
    # CWD tracking
    # ------------------------------------------------------------------

    def _update_cwd(self, result: dict):
        """Read CWD from temp file (local-only, no round-trip needed).

        Uses ``utf-8-sig`` to transparently strip any BOM that PowerShell 5.1's
        ``Out-File -Encoding utf8`` may have left in the file.
        """
        try:
            with open(self._cwd_file, encoding="utf-8-sig") as f:
                cwd_path = f.read().strip()
            if cwd_path:
                self.cwd = cwd_path
        except (OSError, FileNotFoundError):
            pass

        # Strip the CWD marker from output using a precise line-based approach
        # that preserves normal command output (e.g. Get-Location output).
        self._strip_cwd_marker(result)

    def _strip_cwd_marker(self, result: dict):
        """Remove the CWD marker line from output without touching command output.

        The marker line is always the LAST line of output and has the format:
            __HERMES_CWD_{id}__{path}__HERMES_CWD_{id}__

        We only remove that exact line, preserving everything above it — even
        if the command output (like Get-Location) prints the same path on the
        preceding line.
        """
        output = result.get("output", "")
        marker = self._cwd_marker
        last = output.rfind(marker)
        if last == -1:
            return

        # Find the closing marker
        first = output.rfind(marker, max(0, last - 4096), last)
        if first == -1 or first == last:
            return

        # Find the start of the marker line (the \n we injected before it)
        line_start = output.rfind("\n", 0, first)
        # Find the end of the marker line
        line_end = output.find("\n", last + len(marker))
        if line_end != -1:
            line_end += 1  # include the trailing newline
        else:
            line_end = len(output)

        # Only strip if the line contains NOTHING but the marker.
        marker_line = output[line_start + 1 : line_end].strip() if line_start != -1 else output[:line_end].strip()

        # Verify this line is a pure marker (starts and ends with the marker tag)
        if marker_line.startswith(marker) and marker_line.endswith(marker):
            if line_start != -1:
                result["output"] = output[:line_start] + output[line_end:]
            else:
                result["output"] = output[line_end:]

    # ------------------------------------------------------------------
    # Cleanup
    # ------------------------------------------------------------------

    def cleanup(self):
        """Clean up temp files."""
        for f in (self._snapshot_path, self._cwd_file):
            try:
                if f and os.path.exists(f):
                    os.unlink(f)
            except OSError:
                pass
        # Also clean up the temporary .ps1 script.
        temp_dir = self.get_temp_dir()
        ps1_path = os.path.join(temp_dir, f"hermes-ps-{self._session_id}.ps1")
        try:
            if os.path.exists(ps1_path):
                os.unlink(ps1_path)
        except OSError:
            pass
