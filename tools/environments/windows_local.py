"""Windows-native execution environment using PowerShell.

Phase 1 MVP: direct PowerShell spawn without bash intermediate.
No session snapshot (env vars do not persist across commands yet).
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


class WindowsLocalEnvironment(BaseEnvironment):
    """Run commands directly on Windows using PowerShell (no bash)."""

    def __init__(self, cwd: str = "", timeout: int = 60, env: dict = None):
        if cwd:
            cwd = os.path.expanduser(cwd)
        super().__init__(cwd=cwd or os.getcwd(), timeout=timeout, env=env)
        # Phase 1: skip init_session — env snapshot not yet implemented for PowerShell.
        # Commands will run with the current process environment only.
        self._snapshot_ready = False

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

    def _wrap_command(self, command: str, cwd: str) -> str:
        """Build a PowerShell script that cd's, runs the command, and emits CWD."""
        quoted_cwd = _quote_ps_literal(cwd)
        quoted_cwd_file = _quote_ps_literal(self._cwd_file)
        marker = self._cwd_marker

        # Use a here-string so the user's command needs almost no escaping.
        # The only thing that breaks a here-string is '@' at the start of a
        # line immediately followed by a single quote. That's vanishingly rare.
        #
        # $ProgressPreference suppresses CLIXML noise from PowerShell 5.1's
        # first-run module-initialization progress bars.
        # Write CWD via [IO.File]::WriteAllText() — avoids the UTF-8 BOM that
        # PowerShell 5.1's Out-File -Encoding utf8 injects.  The BOM (\ufeff)
        # poisons self.cwd and breaks Set-Location on the next call.
        quoted_cwd_file_bare = self._cwd_file.replace("\\", "/")
        ps_script = (
            f"$ProgressPreference = 'SilentlyContinue'\n"
            f"Set-Location -LiteralPath {quoted_cwd}\n"
            f"$__cmd = @'\n"
            f"{command}\n"
            f"'@\n"
            f"Invoke-Expression $__cmd\n"
            f"$__hermes_ec = $LASTEXITCODE\n"
            f"if ($__hermes_ec -eq $null) {{ $__hermes_ec = 0 }}\n"
            f"[IO.File]::WriteAllText({quoted_cwd_file}, (Get-Location).Path, [System.Text.UTF8Encoding]::new($false))\n"
            f'Write-Output "`n{marker}$((Get-Location).Path){marker}"\n'
            f"exit $__hermes_ec\n"
        )
        return ps_script

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

        # Still strip the marker from output so it's not visible to the user.
        self._extract_cwd_from_output(result)

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
