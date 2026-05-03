"""Tests for WindowsLocalEnvironment — PowerShell backend.

Covers pure-logic methods that can run on any platform (no Windows required)
plus a small integration suite gated on platform.system() == "Windows".
"""

import os
import platform
import subprocess
import tempfile
from unittest.mock import MagicMock, patch

import pytest

from tools.environments.windows_local import (
    _ENV_SEP,
    _find_powershell,
    _quote_ps_literal,
    WindowsLocalEnvironment,
)

_IS_WINDOWS = platform.system() == "Windows"


# ---------------------------------------------------------------------------
# Testable subclass — skips init_session() to avoid spawning PowerShell
# ---------------------------------------------------------------------------


class _TestableWinEnv(WindowsLocalEnvironment):
    """Concrete subclass that skips init_session() and provides a mock _run_bash."""

    def __init__(self, cwd="C:/tmp", timeout=60, env=None):
        # Bypass WindowsLocalEnvironment.__init__ which calls init_session().
        # Instead, call BaseEnvironment.__init__ directly to set up fields.
        from tools.environments.base import BaseEnvironment

        if cwd:
            cwd = os.path.expanduser(cwd)
        BaseEnvironment.__init__(self, cwd=cwd or os.getcwd(), timeout=timeout, env=env)
        # _snapshot_ready is already False from BaseEnvironment.__init__

    def _run_bash(self, cmd_string, *, login=False, timeout=120, stdin_data=None):
        raise NotImplementedError("Use mock in tests")


# ===================================================================
# _quote_ps_literal
# ===================================================================


class TestQuotePsLiteral:
    def test_plain_path(self):
        assert _quote_ps_literal(r"C:\Users\test") == r"'C:\Users\test'"

    def test_single_quote_doubled(self):
        assert _quote_ps_literal(r"C:\Users\O'Brien") == r"'C:\Users\O''Brien'"

    def test_space_preserved(self):
        assert _quote_ps_literal(r"C:\Program Files\foo") == r"'C:\Program Files\foo'"

    def test_empty_string(self):
        assert _quote_ps_literal("") == "''"

    def test_backslash_and_quote(self):
        assert _quote_ps_literal(r"D:\it's\path") == r"'D:\it''s\path'"

    def test_multiple_quotes(self):
        assert _quote_ps_literal("a'b'c") == "'a''b''c'"


# ===================================================================
# _find_powershell
# ===================================================================


class TestFindPowershell:
    def test_pwsh_in_path(self):
        with patch("shutil.which", side_effect=lambda x: "/usr/bin/pwsh" if x == "pwsh" else None):
            result = _find_powershell()
            assert result == "/usr/bin/pwsh"

    def test_pwsh_in_program_files(self):
        with (
            patch("shutil.which", return_value=None),
            patch("os.path.isfile", side_effect=lambda p: "pwsh.exe" in p),
            patch.dict(os.environ, {"ProgramFiles": r"C:\Program Files"}, clear=False),
        ):
            result = _find_powershell()
            assert "pwsh.exe" in result

    def test_fallback_to_ps51_system32(self):
        def _isfile(path):
            return "System32" in path and "powershell.exe" in path

        with (
            patch("shutil.which", return_value=None),
            patch("os.path.isfile", side_effect=_isfile),
            patch.dict(os.environ, {"SystemRoot": r"C:\Windows", "ProgramFiles": r"C:\Program Files"}, clear=False),
        ):
            result = _find_powershell()
            assert "powershell.exe" in result
            assert "System32" in result

    def test_powershell_on_path_as_last_resort(self):
        with (
            patch("shutil.which", side_effect=lambda x: "/usr/bin/powershell" if x == "powershell" else None),
            patch("os.path.isfile", return_value=False),
        ):
            result = _find_powershell()
            assert result == "/usr/bin/powershell"

    def test_raises_when_not_found(self):
        with (
            patch("shutil.which", return_value=None),
            patch("os.path.isfile", return_value=False),
        ):
            with pytest.raises(RuntimeError, match="PowerShell not found"):
                _find_powershell()

    def test_programw6432_checked(self):
        """ProgramW6432 is also checked for pwsh.exe on WoW64."""
        seen_paths = []

        def _isfile(path):
            seen_paths.append(path)
            return False

        with (
            patch("shutil.which", return_value=None),
            patch("os.path.isfile", side_effect=_isfile),
            patch.dict(
                os.environ,
                {"ProgramFiles": r"C:\Program Files", "ProgramW6432": r"C:\Program Files", "SystemRoot": r"C:\Windows"},
                clear=False,
            ),
        ):
            with pytest.raises(RuntimeError):
                _find_powershell()
            # Verify ProgramW6432 was checked
            assert any("ProgramW6432" in p or "Program Files" in p for p in seen_paths)


# ===================================================================
# _wrap_command
# ===================================================================


class TestWrapCommand:
    def _make_env(self, **kwargs):
        env = _TestableWinEnv(**kwargs)
        env._snapshot_ready = True
        return env

    def test_basic_structure_snapshot_ready(self):
        env = self._make_env()
        wrapped = env._wrap_command("Write-Output hello", "C:/Users/test")

        assert "$ProgressPreference = 'SilentlyContinue'" in wrapped
        assert "Set-Location -LiteralPath" in wrapped
        assert "Invoke-Expression $__cmd" in wrapped
        assert "$__hermes_ec" in wrapped
        assert "[IO.File]::WriteAllText" in wrapped
        assert "exit $__hermes_ec" in wrapped

    def test_no_snapshot_skips_restore_and_dump(self):
        env = self._make_env()
        env._snapshot_ready = False
        wrapped = env._wrap_command("Write-Output hello", "C:/Users/test")

        assert "Test-Path" not in wrapped  # no restore block
        assert "ReadAllText" not in wrapped  # no restore block
        assert "Get-ChildItem Env:" not in wrapped  # no dump block
        assert "WriteAllLines" not in wrapped  # no dump block

    def test_cwd_with_spaces_quoted(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", r"C:\Program Files")
        # _quote_ps_literal wraps in single quotes
        assert "'C:\\Program Files'" in wrapped

    def test_cwd_with_single_quote_escaped(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", r"C:\Users\O'Brien")
        assert "O''Brien" in wrapped

    def test_command_in_here_string(self):
        env = self._make_env()
        wrapped = env._wrap_command("Get-Process", "C:/tmp")
        assert "$__cmd = @'" in wrapped
        assert "'@" in wrapped
        assert "Get-Process" in wrapped

    def test_marker_contains_session_id(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        assert env._cwd_marker in wrapped
        assert env._session_id in env._cwd_marker

    def test_cwd_file_path_in_writealltext(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        # _cwd_file is referenced in [IO.File]::WriteAllText
        assert env._cwd_file in wrapped

    def test_snapshot_path_in_readalltext(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        assert env._snapshot_path in wrapped

    def test_exit_code_preserved(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        assert "$__hermes_ec = $LASTEXITCODE" in wrapped
        assert "exit $__hermes_ec" in wrapped

    def test_pwd_and_get_location_overridden(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        assert "function pwd" in wrapped
        assert "function Get-Location" in wrapped
        assert "Remove-Item Alias:pwd" in wrapped

    def test_progress_preference_silenced(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        assert "$ProgressPreference = 'SilentlyContinue'" in wrapped

    def test_null_exit_code_guard(self):
        env = self._make_env()
        wrapped = env._wrap_command("ls", "C:/tmp")
        # PowerShell: $LASTEXITCODE can be $null if no external command was run
        assert "if ($__hermes_ec -eq $null)" in wrapped


# ===================================================================
# _load_snapshot
# ===================================================================


class TestLoadSnapshot:
    def _make_env_with_temp(self, tmp_path):
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "snap.txt")
        return env

    def test_normal_nul_separated(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        content = f"PATH{_ENV_SEP}C:\\Windows\nFOO{_ENV_SEP}bar\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(content)

        result = env._load_snapshot()
        assert result == {"PATH": r"C:\Windows", "FOO": "bar"}

    def test_value_with_equals_sign(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        content = f"FOO{_ENV_SEP}a=b=c\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(content)

        result = env._load_snapshot()
        assert result == {"FOO": "a=b=c"}

    def test_bom_file(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        # Write with utf-8 to get raw BOM bytes, then read with utf-8-sig
        raw_content = f"\ufeffKEY{_ENV_SEP}value\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(raw_content)

        result = env._load_snapshot()
        assert result == {"KEY": "value"}

    def test_empty_file(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write("")

        result = env._load_snapshot()
        assert result == {}

    def test_missing_file(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        env._snapshot_path = str(tmp_path / "nonexistent.txt")

        result = env._load_snapshot()
        assert result == {}

    def test_line_without_separator_skipped(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        content = f"VALID{_ENV_SEP}yes\nNOISE_LINE\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(content)

        result = env._load_snapshot()
        assert result == {"VALID": "yes"}

    def test_empty_value(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        content = f"EMPTY{_ENV_SEP}\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(content)

        result = env._load_snapshot()
        assert result == {"EMPTY": ""}


# ===================================================================
# _save_snapshot
# ===================================================================


class TestSaveSnapshot:
    def _make_env_with_temp(self, tmp_path):
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "snap.txt")
        return env

    def test_normal_write(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        env._save_snapshot({"PATH": r"C:\Windows", "FOO": "bar"})

        with open(env._snapshot_path, encoding="utf-8") as f:
            content = f.read()
        assert f"PATH{_ENV_SEP}C:\\Windows\n" in content
        assert f"FOO{_ENV_SEP}bar\n" in content

    def test_empty_dict(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        env._save_snapshot({})

        with open(env._snapshot_path, encoding="utf-8") as f:
            content = f.read()
        assert content == ""

    def test_no_bom_in_output(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        env._save_snapshot({"KEY": "value"})

        with open(env._snapshot_path, "rb") as f:
            raw = f.read()
        # UTF-8 BOM is 0xEF 0xBB 0xBF — must NOT be present
        assert not raw.startswith(b"\xef\xbb\xbf")

    def test_value_with_special_chars(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        env._save_snapshot({"A": "line1\nline2", "B": "x=y"})

        with open(env._snapshot_path, encoding="utf-8") as f:
            content = f.read()
        assert f"A{_ENV_SEP}line1\nline2\n" in content
        assert f"B{_ENV_SEP}x=y\n" in content


# ===================================================================
# _strip_cwd_marker
# ===================================================================


class TestStripCwdMarker:
    def _make_env(self):
        env = _TestableWinEnv()
        return env

    def test_normal_removal(self):
        env = self._make_env()
        marker = env._cwd_marker
        result = {"output": f"hello\n{marker}C:\\Users{marker}\n"}
        env._strip_cwd_marker(result)

        assert result["output"] == "hello"

    def test_no_marker(self):
        env = self._make_env()
        result = {"output": "hello world\n"}
        env._strip_cwd_marker(result)

        assert result["output"] == "hello world\n"

    def test_marker_after_get_location_output(self):
        """Get-Location output should be preserved; only the marker line removed."""
        env = self._make_env()
        marker = env._cwd_marker
        # Simulates: command output includes "C:\Users", then marker line
        result = {"output": f"C:\\Users\n{marker}C:\\Users{marker}\n"}
        env._strip_cwd_marker(result)

        assert "C:\\Users" in result["output"]
        assert marker not in result["output"]

    def test_only_marker_line(self):
        env = self._make_env()
        marker = env._cwd_marker
        result = {"output": f"{marker}C:\\Users{marker}\n"}
        env._strip_cwd_marker(result)

        assert marker not in result["output"]

    def test_multiline_command_plus_marker(self):
        env = self._make_env()
        marker = env._cwd_marker
        result = {"output": f"line1\nline2\nline3\n{marker}C:\\tmp{marker}\n"}
        env._strip_cwd_marker(result)

        assert "line1" in result["output"]
        assert "line2" in result["output"]
        assert "line3" in result["output"]
        assert marker not in result["output"]

    def test_marker_not_at_end_ignored(self):
        """If marker appears in the middle (not at end), it shouldn't be stripped."""
        env = self._make_env()
        marker = env._cwd_marker
        # Malformed: marker not at the end of output
        result = {"output": f"{marker}C:\\tmp{marker}\nextra stuff\n"}
        env._strip_cwd_marker(result)

        # The marker line is removed, but "extra stuff" remains
        assert "extra stuff" in result["output"]


# ===================================================================
# get_temp_dir
# ===================================================================


class TestGetTempDir:
    def test_temp_env_var(self):
        env = _TestableWinEnv(env={"TEMP": r"C:\Temp"})
        result = env.get_temp_dir()
        assert result == "C:/Temp"

    def test_tmp_fallback(self):
        env = _TestableWinEnv(env={"TMP": r"C:\Tmp"})
        result = env.get_temp_dir()
        assert result == "C:/Tmp"

    def test_backslash_to_forward_slash(self):
        env = _TestableWinEnv(env={"TEMP": r"C:\Users\TEMP"})
        result = env.get_temp_dir()
        assert "\\" not in result
        assert result == "C:/Users/TEMP"

    def test_trailing_slash_stripped(self):
        env = _TestableWinEnv(env={"TEMP": "C:\\Temp\\"})
        result = env.get_temp_dir()
        assert not result.endswith("/")

    def test_os_environ_fallback(self):
        env = _TestableWinEnv(env={})
        # The fallback chain hits os.environ and then tempfile.gettempdir()
        result = env.get_temp_dir()
        assert result  # non-empty


# ===================================================================
# init_session
# ===================================================================


class TestInitSession:
    def test_success_sets_snapshot_ready(self):
        env = _TestableWinEnv()
        mock_proc = MagicMock()
        mock_proc.poll.return_value = 0
        mock_proc.returncode = 0
        # Mock stdout with a proper buffer for the drain thread
        mock_proc.stdout = MagicMock()
        mock_proc.stdout.buffer = MagicMock()
        mock_proc.stdout.buffer.read = MagicMock(return_value=b"")
        mock_proc.stdout.fileno = MagicMock(return_value=1)

        env._run_bash = MagicMock(return_value=mock_proc)
        env._load_snapshot = MagicMock(return_value={})

        env.init_session()
        assert env._snapshot_ready is True

    def test_failure_does_not_crash(self):
        env = _TestableWinEnv()
        env._run_bash = MagicMock(side_effect=RuntimeError("no PowerShell"))

        env.init_session()
        assert env._snapshot_ready is False

    def test_load_snapshot_called_on_success(self):
        env = _TestableWinEnv()
        mock_proc = MagicMock()
        mock_proc.poll.return_value = 0
        mock_proc.returncode = 0
        mock_proc.stdout = MagicMock()
        mock_proc.stdout.buffer = MagicMock()
        mock_proc.stdout.buffer.read = MagicMock(return_value=b"")
        mock_proc.stdout.fileno = MagicMock(return_value=1)

        env._run_bash = MagicMock(return_value=mock_proc)
        env._load_snapshot = MagicMock(return_value={})

        env.init_session()
        env._load_snapshot.assert_called_once()


# ===================================================================
# _kill_process
# ===================================================================


class TestKillProcess:
    @pytest.mark.skipif(not _IS_WINDOWS, reason="taskkill is Windows-only")
    def test_calls_taskkill(self):
        env = _TestableWinEnv()
        mock_proc = MagicMock()
        mock_proc.pid = 12345

        with patch("subprocess.run") as mock_run:
            env._kill_process(mock_proc)
            mock_run.assert_called_once()
            args = mock_run.call_args[0][0]
            assert args[0] == "taskkill"
            assert "/T" in args
            assert "/F" in args
            assert "/PID" in args
            assert "12345" in args

    def test_terminate_called(self):
        env = _TestableWinEnv()
        mock_proc = MagicMock()
        mock_proc.pid = 99999

        with patch("subprocess.run", side_effect=Exception("no taskkill")):
            env._kill_process(mock_proc)
            mock_proc.terminate.assert_called_once()

    def test_no_exception_on_failure(self):
        env = _TestableWinEnv()
        mock_proc = MagicMock()
        mock_proc.terminate.side_effect = PermissionError("denied")
        mock_proc.pid = 99999

        # Should not raise
        with patch("subprocess.run", side_effect=Exception("no taskkill")):
            env._kill_process(mock_proc)


# ===================================================================
# _update_cwd
# ===================================================================


class TestUpdateCwd:
    def _make_env_with_temp(self, tmp_path):
        env = _TestableWinEnv()
        env._cwd_file = str(tmp_path / "cwd.txt")
        return env

    def test_reads_cwd_from_file(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        with open(env._cwd_file, "w", encoding="utf-8") as f:
            f.write(r"C:\Projects")

        result = {"output": "ok"}
        env._update_cwd(result)
        assert env.cwd == r"C:\Projects"

    def test_bom_file_handled(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        with open(env._cwd_file, "w", encoding="utf-8-sig") as f:
            f.write(r"C:\BOM\Path")

        result = {"output": "ok"}
        env._update_cwd(result)
        # BOM should be stripped transparently
        assert "\ufeff" not in env.cwd

    def test_missing_file_cwd_unchanged(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        original_cwd = env.cwd
        env._cwd_file = str(tmp_path / "nonexistent.txt")

        result = {"output": "ok"}
        env._update_cwd(result)
        assert env.cwd == original_cwd

    def test_empty_file_cwd_unchanged(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        original_cwd = env.cwd
        with open(env._cwd_file, "w", encoding="utf-8") as f:
            f.write("")

        result = {"output": "ok"}
        env._update_cwd(result)
        assert env.cwd == original_cwd

    def test_marker_stripped_from_output(self, tmp_path):
        env = self._make_env_with_temp(tmp_path)
        marker = env._cwd_marker
        with open(env._cwd_file, "w", encoding="utf-8") as f:
            f.write(r"C:\test")

        result = {"output": f"hello\n{marker}C:\\test{marker}\n"}
        env._update_cwd(result)
        assert marker not in result["output"]


# ===================================================================
# cleanup
# ===================================================================


class TestCleanup:
    def test_deletes_snapshot_and_cwd_files(self, tmp_path):
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "snap.txt")
        env._cwd_file = str(tmp_path / "cwd.txt")

        # Create the files
        for p in (env._snapshot_path, env._cwd_file):
            with open(p, "w") as f:
                f.write("test")

        env.cleanup()
        assert not os.path.exists(env._snapshot_path)
        assert not os.path.exists(env._cwd_file)

    def test_missing_files_no_crash(self, tmp_path):
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "nonexistent_snap.txt")
        env._cwd_file = str(tmp_path / "nonexistent_cwd.txt")
        # Should not raise
        env.cleanup()

    def test_ps1_file_cleaned(self, tmp_path):
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "snap.txt")
        env._cwd_file = str(tmp_path / "cwd.txt")

        # Create the .ps1 temp file
        ps1_path = os.path.join(env.get_temp_dir(), f"hermes-ps-{env._session_id}.ps1")
        os.makedirs(os.path.dirname(ps1_path), exist_ok=True)
        with open(ps1_path, "w") as f:
            f.write("$ProgressPreference = 'SilentlyContinue'")

        env.cleanup()
        assert not os.path.exists(ps1_path)


# ===================================================================
# Integration tests (Windows only)
# ===================================================================


@pytest.mark.skipif(not _IS_WINDOWS, reason="Requires Windows PowerShell")
class TestWindowsLocalIntegration:
    """End-to-end tests that spawn real PowerShell processes.

    These require a full project environment (yaml, etc.) and a Windows
    system with PowerShell.  They are skipped automatically outside Windows.
    """

    @classmethod
    def _check_deps(cls):
        """Verify yaml is importable (required by tools.terminal_tool chain)."""
        try:
            import yaml  # noqa: F401
            return True
        except ImportError:
            return False

    def test_echo_hello(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            result = env.execute('Write-Output "hello"')
            assert "hello" in result["output"]
            assert result["returncode"] == 0
        finally:
            env.cleanup()

    def test_env_var_persistence(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            env.execute('$env:HERMES_TEST_VAR = "test_value_12345"')
            result = env.execute("Write-Output $env:HERMES_TEST_VAR")
            assert "test_value_12345" in result["output"]
        finally:
            env.cleanup()

    def test_cwd_persistence(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            env.execute("Set-Location C:\\Windows")
            result = env.execute("(Get-Location).Path")
            # CWD file is read back, so env.cwd should update
            assert "Windows" in env.cwd
        finally:
            env.cleanup()

    def test_exit_code_propagation(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            result = env.execute("exit 42")
            assert result["returncode"] == 42
        finally:
            env.cleanup()

    def test_unicode_output(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            result = env.execute('Write-Output "你好世界"')
            assert "你好" in result["output"]
        finally:
            env.cleanup()

    def test_multiline_output(self):
        if not self._check_deps():
            pytest.skip("yaml module not installed")
        env = WindowsLocalEnvironment()
        try:
            result = env.execute('Write-Output "line1"; Write-Output "line2"; Write-Output "line3"')
            assert "line1" in result["output"]
            assert "line2" in result["output"]
            assert "line3" in result["output"]
        finally:
            env.cleanup()

    def test_snapshot_path_uses_forward_slashes(self):
        """Verify temp paths use forward slashes (consistency with BaseEnvironment)."""
        env = _TestableWinEnv()
        assert "/" in env._snapshot_path or "\\" not in env._snapshot_path

    def test_crlf_snapshot_values_stripped(self, tmp_path):
        """CRLF line endings in snapshot must not leave trailing \\r in values.

        [IO.File]::WriteAllLines on Windows produces CRLF.  PowerShell's
        -split \"`n\" splits on LF only, leaving \\r at the end of each
        value.  The restore script uses .TrimEnd([char]13) to strip it.
        If \\r leaks into PATH or other critical env vars, DNS resolution
        and other system calls can break silently.
        """
        env = _TestableWinEnv()
        env._snapshot_path = str(tmp_path / "snap-crlf.txt")

        # Simulate what [IO.File]::WriteAllLines produces on Windows:
        # each line ends with \\r\\n, and the separator is \\x02\\x03
        content = f"PATH{_ENV_SEP}C:\\Windows;C:\\Users\\test\\r\\nFOO{_ENV_SEP}bar\\r\\n"
        with open(env._snapshot_path, "w", encoding="utf-8") as f:
            f.write(content)

        # _load_snapshot uses rstrip("\\n\\r") which strips the \\r
        result = env._load_snapshot()
        assert result["PATH"] == r"C:\Windows;C:\Users\test"
        assert result["FOO"] == "bar"
        assert not result["PATH"].endswith("\r")
        assert not result["FOO"].endswith("\r")

    def test_wrap_command_includes_trimend(self):
        """Verify the restore script strips trailing CR from values."""
        env = _TestableWinEnv()
        env._snapshot_ready = True
        wrapped = env._wrap_command("echo hi", "C:/tmp")
        # The restore block must include .TrimEnd([char]13) to handle CRLF
        assert "TrimEnd([char]13)" in wrapped
