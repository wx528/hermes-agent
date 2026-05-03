# Hermes Windows Native Terminal Refactor Plan

Branch: `feature/windows-native-terminal`

## Goal

Make Hermes on Windows shed the Git Bash middleware and execute commands directly
through PowerShell natively, following OpenClaw's Windows strategy: acknowledge
that Windows is not POSIX, and use the platform's native shell.

## Progress

### Phase 0: Windows Foundation Patches ✅

| Change | File | Description |
|--------|------|-------------|
| `select.select` fix | `base.py` | Use `stdout.buffer.read()` instead of `select.select` on Windows |
| `shlex.quote` paths | `base.py` | Quote snapshot/cwd file paths (Windows user dirs contain spaces) |
| `get_temp_dir()` | `local.py` | Return `TEMP` path on Windows instead of `/tmp` |

Commit: `50ff3f82` (merged into Phase 1 commit)

### Phase 1: WindowsLocalEnvironment ✅

| Change | File | Description |
|--------|------|-------------|
| New WindowsLocalEnvironment | `windows_local.py` | Spawn PowerShell directly, bypassing bash |
| PowerShell lookup priority | `windows_local.py` | pwsh7 → powershell5.1 |
| BOM-safe CWD write | `windows_local.py` | `[IO.File]::WriteAllText` + `utf-8-sig` read |
| Process termination | `windows_local.py` | `taskkill /T /F /PID` |
| CLIXML noise suppression | `windows_local.py` | `$ProgressPreference = 'SilentlyContinue'` |
| Routing change | `terminal_tool.py` | Windows → `_WindowsLocalEnvironment` |
| Fallback switch | `terminal_tool.py` | `HERMES_USE_GIT_BASH=1` reverts to legacy mode |
| hermes.ps1 | `hermes.ps1` | Comment out `HERMES_GIT_BASH_PATH` |
| hermes.cmd | `hermes.cmd` | cmd.exe launcher |

Commits: `50ff3f82`, `82e1ac32`, `367417a0`

Key fixes:
- BOM contaminating cwd → `NotADirectoryError` (`Out-File -Encoding utf8` writes BOM on PS5.1)
- `Get-Location`/`pwd` produces no output in `-NonInteractive -File` mode → Override with `Write-Output`

### Phase 2: PowerShell Env Snapshot ✅

| Change | File | Description |
|--------|------|-------------|
| `init_session()` | `windows_local.py` | Capture environment variables to NUL-delimited snapshot file on first launch |
| `_wrap_command()` restore env | `windows_local.py` | Restore `$env:` from snapshot before each command |
| `_wrap_command()` re-export env | `windows_local.py` | Re-export to snapshot after each command (captures new variables) |
| BOM-free write | `windows_local.py` | `[IO.File]::WriteAllText/WriteAllLines` + `UTF8Encoding($false)` |

Commit: `8446226b`

Test verification:
- `$env:MY_VAR = "hello hermes"` → persists across commands ✅
- `Write-Output $env:MY_VAR` → `hello hermes` ✅
- `Get-Location` → `D:\Agents\hermes-agent` ✅
- `Set-Location C:\Users` + `Get-Location` → `C:\Users` ✅

### Phase 3: De-bash BaseEnvironment (TODO)

Goal: Refactor `BaseEnvironment` from "bash-specific" to "generic execution framework".

**3.1 Rename / Abstract**

```
_run_bash(...)      →  _spawn(self, script, ...) -> ProcessHandle
_wrap_command(...)  →  _build_script(self, command, cwd) -> str
init_session()      →  Keep as-is; subclasses handle concrete implementation
```

**3.2 Introduce ShellBackend Strategy**

```python
class ShellBackend(ABC):
    def find_shell(self) -> str: ...
    def capture_env(self, cwd) -> dict: ...
    def build_script(self, command, cwd, env_snapshot) -> str: ...
    def parse_env_output(self, output) -> dict: ...

class BashBackend(ShellBackend): ...       # Unix
class PowerShellBackend(ShellBackend): ... # Windows

class LocalEnvironment(BaseEnvironment):
    def __init__(self):
        self._shell = PowerShellBackend() if _IS_WINDOWS else BashBackend()
```

**3.3 Test Requirements**

- Unix regression: all bash-path behavior unchanged
- Windows: all PowerShell-path behavior unchanged
- Remote backends (Docker/SSH/Modal) unaffected

### Phase 4: Windows Encoding & .cmd Shim (TODO)

**4.1 Encoding Auto-Detection**

```python
# tools/environments/windows_encoding.py
def resolve_windows_console_encoding() -> str | None:
    """Detect Windows console codepage (936=GBK, 65001=UTF-8)."""
    result = subprocess.run(
        ["cmd.exe", "/d", "/s", "/c", "chcp"],
        capture_output=True, text=True, timeout=5,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    # Parse "Active code page: 936"
    ...
```

Inspired by OpenClaw's `windows-encoding.ts`:
- Query `chcp` at runtime for the codepage
- Use `TextDecoder` for streaming decode
- Handle GBK characters split across multiple chunks

**4.2 .cmd Command Resolution**

Similar to OpenClaw's `resolveNpmArgvForWindows`:

```python
def resolve_windows_argv(argv: list[str]) -> list[str]:
    """Resolve command for Windows spawn.
    - npm/npx → [node.exe, npm-cli.js, ...]
    - .cmd/.bat → [cmd.exe, "/c", ...]
    - .exe → direct spawn
    """
```

Addresses Node.js CVE-2024-27980: directly spawning `.cmd` files can cause EINVAL.

### Phase 5: Cleanup (TODO)

| Task | Description |
|------|-------------|
| Remove `HERMES_GIT_BASH_PATH` | No longer needed on Windows |
| Remove `_find_bash` Windows branch | No longer searching for bash |
| Documentation update | Change "Windows not supported" in README to supported |
| Test coverage | Windows spawn, encoding, exit code, .cmd shim |

## Fallback Plan

Set `HERMES_USE_GIT_BASH=1` at any time to revert to the legacy Git Bash mode:

```powershell
$env:HERMES_USE_GIT_BASH=1
.\hermes.ps1
```

## Architecture Reference

```
Current (Phase 1-2 complete):
  terminal_tool → WindowsLocalEnvironment → PowerShell (spawn-per-call)
                                             ↑ env snapshot restore/re-export

Target (Phase 3-4):
  terminal_tool → LocalEnvironment → ShellBackend (Bash/PowerShell)
                                     ↑ unified interface, platform-adaptive
```

## Known Issues

1. `Get-Location` output consumed by marker-stripping logic → Fixed (override with Write-Output)
2. PowerShell 5.1 BOM contaminating cwd → Fixed (BOM-free write + utf-8-sig read)
3. `python` command may trigger Windows Store stub → Deferred to Phase 4
4. GBK-encoded command output may garble → Deferred to Phase 4 encoding detection
5. Per-command env dump includes all variables; may be slow with large environments → Can optimize with incremental diff
