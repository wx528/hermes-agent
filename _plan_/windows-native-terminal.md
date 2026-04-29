# Hermes Windows Native Terminal Refactor Plan

Branch: `feature/windows-native-terminal`

## 目标

让 Hermes 在 Windows 上摆脱 Git Bash 中间层，直接使用 PowerShell 原生执行命令，
借鉴 OpenClaw 的 Windows 策略：承认 Windows 不是 POSIX，用平台原生 shell。

## 进度

### Phase 0: Windows 基础补丁 ✅

| 改动 | 文件 | 说明 |
|------|------|------|
| `select.select` 修复 | `base.py` | Windows 用 `stdout.buffer.read()` 替代 `select.select` |
| `shlex.quote` 路径 | `base.py` | snapshot/cwd 文件路径加引号（Windows 用户目录有空格） |
| `get_temp_dir()` | `local.py` | Windows 返回 `TEMP` 路径而非 `/tmp` |

Commit: `50ff3f82` (合并在 Phase 1 提交中)

### Phase 1: WindowsLocalEnvironment ✅

| 改动 | 文件 | 说明 |
|------|------|------|
| 新建 WindowsLocalEnvironment | `windows_local.py` | 直接 spawn PowerShell，不经过 bash |
| PowerShell 查找优先级 | `windows_local.py` | pwsh7 → powershell5.1 |
| BOM 安全 CWD 写入 | `windows_local.py` | `[IO.File]::WriteAllText` + `utf-8-sig` 读取 |
| 进程终止 | `windows_local.py` | `taskkill /T /F /PID` |
| CLIXML 噪音抑制 | `windows_local.py` | `$ProgressPreference = 'SilentlyContinue'` |
| 路由改造 | `terminal_tool.py` | Windows → `_WindowsLocalEnvironment` |
| 回退开关 | `terminal_tool.py` | `HERMES_USE_GIT_BASH=1` 回退旧模式 |
| hermes.ps1 | `hermes.ps1` | 注释掉 `HERMES_GIT_BASH_PATH` |
| hermes.cmd | `hermes.cmd` | cmd.exe 版启动器 |

Commits: `50ff3f82`, `82e1ac32`, `367417a0`

关键修复：
- BOM 污染 cwd → `NotADirectoryError`（`Out-File -Encoding utf8` 在 PS5.1 写 BOM）
- `Get-Location`/`pwd` 在 `-NonInteractive -File` 模式不输出 → Override 为 `Write-Output`

### Phase 2: PowerShell Env Snapshot ✅

| 改动 | 文件 | 说明 |
|------|------|------|
| `init_session()` | `windows_local.py` | 首次启动捕获环境变量到 NUL 分隔的快照文件 |
| `_wrap_command()` 恢复环境 | `windows_local.py` | 每次命令前从快照恢复 `$env:` |
| `_wrap_command()` 重导出环境 | `windows_local.py` | 每次命令后重新导出到快照（捕获新变量） |
| BOM-free 写入 | `windows_local.py` | `[IO.File]::WriteAllText/WriteAllLines` + `UTF8Encoding($false)` |

Commit: `8446226b`

测试验证：
- `$env:MY_VAR = "hello hermes"` → 跨命令保持 ✅
- `Write-Output $env:MY_VAR` → `hello hermes` ✅
- `Get-Location` → `D:\Agents\hermes-agent` ✅
- `Set-Location C:\Users` + `Get-Location` → `C:\Users` ✅

### Phase 3: BaseEnvironment 去 Bash 化 (待做)

目标：把 `BaseEnvironment` 从 "bash 专属" 重构为 "通用执行框架"。

**3.1 重命名/抽象化**

```
_run_bash(...)      →  _spawn(self, script, ...) -> ProcessHandle
_wrap_command(...)  →  _build_script(self, command, cwd) -> str
init_session()      →  保持，子类负责具体实现
```

**3.2 引入 ShellBackend 策略**

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

**3.3 测试要求**

- Unix 回归测试：所有 bash 路径行为不变
- Windows 测试：所有 PowerShell 路径行为不变
- 远程后端（Docker/SSH/Modal）不受影响

### Phase 4: Windows 编码与 .cmd Shim (待做)

**4.1 编码自动检测**

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

借鉴 OpenClaw 的 `windows-encoding.ts`：
- 运行时查 `chcp` 获取代码页
- 用 `TextDecoder` 做流式解码
- 处理 GBK 字符被拆分成多个 chunk 的情况

**4.2 .cmd 命令解析**

类似 OpenClaw 的 `resolveNpmArgvForWindows`：

```python
def resolve_windows_argv(argv: list[str]) -> list[str]:
    """Resolve command for Windows spawn.
    - npm/npx → [node.exe, npm-cli.js, ...]
    - .cmd/.bat → [cmd.exe, "/c", ...]
    - .exe → direct spawn
    """
```

解决 Node.js CVE-2024-27980：直接 spawn `.cmd` 文件可能导致 EINVAL。

### Phase 5: 清理 (待做)

| 任务 | 说明 |
|------|------|
| 删除 `HERMES_GIT_BASH_PATH` | Windows 不再需要 Git Bash |
| 删除 `_find_bash` Windows 分支 | 不再找 bash |
| 文档更新 | README 中 "Windows not supported" 改为支持说明 |
| 测试覆盖 | Windows spawn、编码、exit code、.cmd shim |

## 回退方案

任何时候设 `HERMES_USE_GIT_BASH=1` 即可回退到旧的 Git Bash 模式：

```powershell
$env:HERMES_USE_GIT_BASH=1
.\hermes.ps1
```

## 参考架构

```
当前 (Phase 1-2 完成):
  terminal_tool → WindowsLocalEnvironment → PowerShell (spawn-per-call)
                                             ↑ env snapshot 恢复/重导出

目标 (Phase 3-4):
  terminal_tool → LocalEnvironment → ShellBackend (Bash/PowerShell)
                                     ↑ 统一接口，平台自适应
```

## 已知问题

1. `Get-Location` 输出被 marker 剥离逻辑吃掉 → 已修复（override 为 Write-Output）
2. PowerShell 5.1 BOM 污染 cwd → 已修复（BOM-free 写入 + utf-8-sig 读取）
3. `python` 命令可能触发 Windows Store stub → 待 Phase 4 处理
4. GBK 编码的命令输出可能乱码 → 待 Phase 4 编码检测
5. 每次命令 env dump 包含所有环境变量，大环境下可能较慢 → 可优化为增量 diff
