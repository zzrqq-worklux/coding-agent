# my-excel MCP 启动失败修复 + MCP venv 隔离 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修复安装器部署的 `my-excel` MCP 启动即崩溃的问题，并让 Excel / Outlook 各自使用独立 venv，避免二者依赖的 `mcp` / `fastmcp` 版本互相破坏。

**Architecture:** Excel MCP 固定 `mcp>=1.24,<2`（它使用 mcp 1.x 的 `mcp.server.fastmcp` API）；Outlook MCP 固定 `fastmcp>=2.0,<5`。安装器为每个 MCP 创建 `mcp-servers\venv-<key>`，配置中的 `command` 指向对应 venv 的 python；安装后用真实模块 `*.server` 做导入探针，失败则强制重装依赖，仍失败则不写配置。

**Tech Stack:** Windows PowerShell 5.1、Pester 3.4、Python 3.12、MCP Python SDK 1.x、FastMCP 2.x/4.x、ps2exe。

**Spec:** 本次故障的实证：`pip install` 按未限制的 `mcp` 装了 mcp 2.3.0，`excel_mcp/server.py` 的 `from mcp.server.fastmcp import FastMCP` 直接 `ModuleNotFoundError`；安装器只执行 `import excel_mcp`（空 `__init__.py`）所以误判为成功。

## Global Constraints

- 只支持 Windows 10/11；安装器运行在 Windows PowerShell 5.1。
- 安装目标仍为 `~/.config/opencode`，配置格式保持 V1/V2 兼容。
- 不删除旧版共享 `mcp-servers\venv`（可能仍被未重装的组件引用）。
- 不修改 OpenCode 自身的配置 schema。
- 版本号统一升到 `4.2.2`（`install.ps1`、`installer-core.ps1`、`build.ps1`）。
- 未经用户要求不提交 git。

---

### Task 1: 为 venv 路径与依赖 pin 写失败测试

**Files:**
- Modify: `tests/installer-core.Tests.ps1`（追加两个 Describe）
- Test: 同上

**Interfaces:**
- Produces: 对 `Get-McpVenvDir -McpRoot <string> -Key <string>` 的契约：返回 `<McpRoot>\venv-<Key>`，空 Key 抛错。
- Produces: `excel-mcp/pyproject.toml` 必须含 `mcp>=1.24,<2` 且不含未使用的 `fastmcp`；`outlook-mcp/pyproject.toml` 必须含 `fastmcp>=2.0,<5`。

- [ ] **Step 1: 写失败测试**

在 `tests/installer-core.Tests.ps1` 末尾追加：

```powershell
Describe 'Get-McpVenvDir' {
    It 'creates one venv directory per MCP key' {
        $root = 'C:\Users\me\.config\opencode\mcp-servers'
        Get-McpVenvDir -McpRoot $root -Key 'excel' | Should Be (Join-Path $root 'venv-excel')
        Get-McpVenvDir -McpRoot $root -Key 'outlook' | Should Be (Join-Path $root 'venv-outlook')
    }
    It 'throws when the key is empty' {
        { Get-McpVenvDir -McpRoot 'C:\mcp' -Key '' } | Should Throw
    }
}

Describe 'MCP dependency pins' {
    It 'pins excel MCP to the mcp 1.x API and drops unused fastmcp' {
        $path = Join-Path $here '..\src\assets\excel-mcp\pyproject.toml'
        $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $text | Should Match 'mcp>=1\.24,<2'
        $text | Should Not Match '"fastmcp"'
    }
    It 'pins outlook MCP below the next fastmcp major' {
        $path = Join-Path $here '..\src\assets\outlook-mcp\pyproject.toml'
        $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $text | Should Match 'fastmcp>=2\.0,<5'
    }
}
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```powershell
powershell -ExecutionPolicy Bypass -Command "Invoke-Pester -Script tests\installer-core.Tests.ps1 -PassThru"
```
Expected: `Get-McpVenvDir` 未定义；两个 pin 断言失败（当前文件是裸 `mcp` / `fastmcp`）。

---

### Task 2: 实现 venv 路径函数与依赖 pin

**Files:**
- Modify: `src/lib/installer-core.ps1`（新增函数）
- Modify: `src/assets/excel-mcp/pyproject.toml`
- Modify: `src/assets/outlook-mcp/pyproject.toml`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- Produces: `Get-McpVenvDir -McpRoot <string> -Key <string>`，供 `install.ps1` 的 `Get-McpPython` 使用。

- [ ] **Step 1: 新增 `Get-McpVenvDir`**

在 `installer-core.ps1` 的路径/打包工具函数区域（`Expand-ZipPackage` 附近）加入：

```powershell
function Get-McpVenvDir {
    param([string]$McpRoot, [string]$Key)

    if ([string]::IsNullOrWhiteSpace($McpRoot)) { throw 'McpRoot 不能为空' }
    if ([string]::IsNullOrWhiteSpace($Key)) { throw 'Key 不能为空' }
    return (Join-Path $McpRoot ("venv-" + $Key))
}
```

- [ ] **Step 2: 固定 Excel 依赖**

`src/assets/excel-mcp/pyproject.toml` 的 `dependencies` 改为：

```toml
dependencies = [
    "mcp>=1.24,<2",
    "openpyxl",
    "typer",
]
```

- [ ] **Step 3: 固定 Outlook 依赖**

`src/assets/outlook-mcp/pyproject.toml` 的 `dependencies` 改为：

```toml
dependencies = [
    "pywin32>=306",
    "fastmcp>=2.0,<5",
]
```

- [ ] **Step 4: 运行测试确认通过**

Run:
```powershell
powershell -ExecutionPolicy Bypass -Command "Invoke-Pester -Script tests\installer-core.Tests.ps1 -PassThru"
```
Expected: 新增 4 个测试通过，原 170 个仍通过。

---

### Task 3: 安装器改为每个 MCP 独立 venv + 真实启动探针

**Files:**
- Modify: `src/install.ps1:96-97`（去掉单一 `$script:VenvDir`）
- Modify: `src/install.ps1:1036-1130`（`Get-McpPython` / `Install-McpPackage` / 两个 Install-*）
- Modify: `src/install.ps1:1203-1229`（按组件取各自 python）
- Modify: `src/install.ps1:1401`（报告显示 MCP 根目录）
- Modify: `src/install.ps1:121-138`（catalog 增加 `ProbeModule`）

**Interfaces:**
- Consumes: `Get-McpVenvDir -McpRoot -Key`（Task 2）。
- Produces: `Get-McpPython -Mcp <hashtable> -HasPython <bool>`；每个 MCP 的 venv 路径为 `$script:McpDir\venv-<Key>`。
- Produces: `Install-McpPackage` 在探针失败后返回 `$false`，调用方不再写该 MCP 的配置。

- [ ] **Step 1: 目录变量**

把：

```powershell
$script:VenvDir = Join-Path $script:McpDir 'venv'
```

改为（保留 `$script:McpDir`，删除 `$script:VenvDir`）：

```powershell
# 每个 MCP 使用独立 venv：venv-excel / venv-outlook
```

- [ ] **Step 2: catalog 增加探针模块**

```powershell
$script:OutlookMcp = @{
    ...
    ProbeModule = 'outlook_com_mcp.server'
}
$script:ExcelMcp = @{
    ...
    ProbeModule = 'excel_mcp.server'
}
```

- [ ] **Step 3: `Get-McpPython` 按 key 建 venv**

```powershell
function Get-McpPython {
    param([hashtable]$Mcp, [bool]$HasPython)

    if (-not $HasPython) { return $null }
    $venvDir = Get-McpVenvDir -McpRoot $script:McpDir -Key $Mcp.Key
    $venvPython = Join-Path $venvDir 'Scripts\python.exe'
    if (Test-Path -LiteralPath $venvPython) {
        try {
            $probe = Invoke-CommandWithTimeout -FilePath $venvPython -Arguments @('--version') -TimeoutSec 15 -HeartbeatSec 0
            if (-not $probe.TimedOut -and $probe.ExitCode -eq 0) { return $venvPython }
        } catch { $null = $_ }
        Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($script:DryRun) {
        Write-Status 'plan' "创建 $($Mcp.Display) 独立环境：$venvDir"
        return $venvPython
    }
    try {
        $pythonCommand = Get-Command python -ErrorAction Stop
        $pythonExe = if ($pythonCommand.Source) { [string]$pythonCommand.Source } else { 'python' }
        if (-not (Invoke-PipStep -Label "创建 $($Mcp.Display) 独立 Python 环境" -PythonExe $pythonExe -Arguments @('-m', 'venv', $venvDir) -TimeoutSec 180)) {
            throw 'python -m venv 失败'
        }
        return $venvPython
    } catch {
        Write-Status 'warn' "$($Mcp.Display) 独立环境创建失败，回退全局 Python：$($_.Exception.Message)"
        return 'python'
    }
}
```

- [ ] **Step 4: `Install-McpPackage` 用真实 server 模块探针，失败重装依赖并终止**

安装成功后的检查段改为：

```powershell
    # 同版本重跑时 pip 可能跳过安装；强制刷新本地包代码（不重装依赖）
    if (-not (Invoke-PipStep -Label "刷新 $($Mcp.Display) 代码" -PythonExe $PythonExe -Arguments @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', '--no-deps', $packageRoot) -Quiet)) {
        Write-Status 'warn' "$($Mcp.Display) 代码刷新失败，当前可能是旧版本"
    }

    $probeModule = if ($Mcp.ProbeModule) { [string]$Mcp.ProbeModule } else { [string]$Mcp.Module }
    $probeArgs = @('-c', "import $probeModule")
    if (-not (Invoke-PipStep -Label "启动检查 $($Mcp.Display)" -PythonExe $PythonExe -Arguments $probeArgs -TimeoutSec 60 -Quiet)) {
        Write-Status 'warn' "$($Mcp.Display) 启动检查失败，尝试强制重装依赖..."
        $repairArgs = @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', $packageRoot, '-i', $script:PipMirror)
        if (-not (Invoke-PipStep -Label "修复 $($Mcp.Display) 依赖" -PythonExe $PythonExe -Arguments $repairArgs)) {
            Write-Status 'warn' '镜像修复失败，尝试默认源...'
            $repairArgs = @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', $packageRoot)
            if (-not (Invoke-PipStep -Label "修复 $($Mcp.Display) 依赖（默认源）" -PythonExe $PythonExe -Arguments $repairArgs)) {
                Write-Status 'error' "$($Mcp.Display) 依赖修复失败"
                return $false
            }
        }
        if (-not (Invoke-PipStep -Label "复检 $($Mcp.Display)" -PythonExe $PythonExe -Arguments $probeArgs -TimeoutSec 60 -Quiet)) {
            Write-Status 'error' "$($Mcp.Display) 启动检查仍失败，已跳过配置（避免 OpenCode 报 MCP server failed）"
            return $false
        }
    }
    Write-Status 'ok' "$($Mcp.Display) 启动检查通过"
    return $true
```

- [ ] **Step 5: 每个组件用自己的 python**

`Install-Components` 中把单个 `$pythonExe = Get-McpPython -HasPython $hasPython` 改为：

```powershell
    if ($selected -contains 'outlook') {
        $outlookPython = Get-McpPython -Mcp $script:OutlookMcp -HasPython $hasPython
        $outlook = Install-OutlookMcp -PythonExe $outlookPython
        ...
    }
    if ($selected -contains 'excel') {
        $excelPython = Get-McpPython -Mcp $script:ExcelMcp -HasPython $hasPython
        $excel = Install-ExcelMcp -PythonExe $excelPython
        ...
    }
```

- [ ] **Step 6: 报告路径**

```powershell
    Write-KeyValue 'MCP 环境' $script:McpDir
```

- [ ] **Step 7: 语法自检**

Run:
```powershell
powershell -NoProfile -Command "$null = [System.Management.Automation.Language.Parser]::ParseFile('src\install.ps1', [ref]$null, [ref]$errors); if ($errors) { $errors; exit 1 }"
```
Expected: 无输出、退出码 0。

---

### Task 4: 版本号、文档与构建

**Files:**
- Modify: `src/install.ps1:21`
- Modify: `src/lib/installer-core.ps1:3`
- Modify: `build/build.ps1:30`
- Modify: `README.md`（标题、v4.2.2 章节、MCP 路径说明、测试数量）

- [ ] **Step 1: 版本号**

```powershell
$INSTALLER_VERSION = '4.2.2'          # install.ps1
$script:InstallerCoreVersion = '4.2.2' # installer-core.ps1
$version = '4.2.2.0'                   # build.ps1
```

- [ ] **Step 2: README**

- 标题改 `v4.2.2`。
- 新增“v4.2.2 MCP 修复”章节：mcp 2.x 破坏 `mcp.server.fastmcp`、Excel 固定 `mcp<2`、每个 MCP 独立 venv、真实启动探针、失败不写配置。
- 把 `~/.config/opencode/mcp-servers/venv` 改为 `~/.config/opencode/mcp-servers/venv-excel` / `venv-outlook`。
- 测试数量按实际结果更新。

- [ ] **Step 3: 跑全部单测**

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -SkipTests -NoSmokeTest -NoPackage
```
不要用这个验证测试；改用：
```powershell
powershell -ExecutionPolicy Bypass -Command "Invoke-Pester -Script tests\installer-core.Tests.ps1 -PassThru"
```
Expected: 0 failed。

---

### Task 5: 构建两个 exe 并做端到端验证

**Files:**
- Output: `build/dist/OpenCode安装器-便携版.exe`、`build/dist/OpenCode安装器-厚重版.exe`
- Output: `build/release/opencode安装器-便携版.zip`、`build/release/opencode安装器-厚重版.zip`

- [ ] **Step 1: 构建**

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```
Expected: 单测全过 → 资源校验 → 两个 exe → DryRun 冒烟 → SelfTest → release zip 全部成功。

- [ ] **Step 2: 端到端验证 MCP 包与握手**

在临时目录创建两个独立 venv，按安装器同样的命令安装并探测：

```powershell
# Excel venv：必须能完成 initialize 握手
python -m venv $env:TEMP\verify-excel
& $env:TEMP\verify-excel\Scripts\python.exe -m pip install src\assets\excel-mcp -i https://pypi.tuna.tsinghua.edu.cn/simple
& $env:TEMP\verify-excel\Scripts\python.exe -c "import excel_mcp.server; print('excel ok')"

# Outlook venv：必须能导入 server
python -m venv $env:TEMP\verify-outlook
& $env:TEMP\verify-outlook\Scripts\python.exe -m pip install src\assets\outlook-mcp -i https://pypi.tuna.tsinghua.edu.cn/simple
& $env:TEMP\verify-outlook\Scripts\python.exe -c "import outlook_com_mcp.server; print('outlook ok')"
```

Excel 还要用 JSON-RPC `initialize` 请求确认返回 `serverInfo.name = my-excel`。

- [ ] **Step 3: 确认配置不会写入坏 MCP**

用临时目录运行构建出的 exe：

```powershell
& "build\dist\OpenCode安装器-便携版.exe" -DryRun -TestRoot $env:TEMP\opencode-installer-dryrun -NoPause
```
Expected: 日志中出现 `创建 Outlook ... 独立环境：...\venv-outlook` 与 `创建 Excel ... 独立环境：...\venv-excel`，退出码 0。

---

## Self-Review

- 依赖 pin 覆盖根因（mcp 2.x 删除 `mcp.server.fastmcp`）。
- venv 隔离覆盖两个 MCP 的 `mcp`/`fastmcp` 冲突。
- 真实 `*.server` 探针覆盖安装器误判。
- 失败不写配置覆盖“坏配置导致 OpenCode 报 MCP server failed”。
- 旧共享 venv 保留，避免破坏未重装组件的配置。
