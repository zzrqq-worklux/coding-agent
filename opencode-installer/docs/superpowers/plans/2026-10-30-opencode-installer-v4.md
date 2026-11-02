# OpenCode 安装器 v4 Implementation Plan

**Status:** Completed 2026-10-30. Verification: 123 Pester tests pass; both editions built; embedded-resource checks pass; both exes pass `-DryRun` smoke tests.

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把现有 OpenCode 安装器升级到 v4.0.0，适配 OpenCode V2、DeepSeek 动态模型发现，并产出「便携版」（仅 MCP）与「厚重版」（PPT 双技能 + MCP）两个 exe。

**Architecture:** 单一 PowerShell 5.1 代码库；`install.ps1` 负责流程，`lib/installer-core.ps1` 放纯函数，`lib/ui.ps1` 放 TUI。构建时把 `edition.txt` 和对应资源嵌入 exe，运行时读取 edition 决定组件菜单与必需资源。

**Tech Stack:** Windows PowerShell 5.1、PS2EXE、Pester 3.4、Python 3.12、Node.js LTS、OpenCode V2、DeepSeek OpenAI-compatible API。

**Spec:** 本文件（设计已在 2026-10-30 会话中确认）。

## Global Constraints

- 只支持 Windows 10/11、PowerShell 5.1。
- 不派 subagent；所有实现与验证在主会话完成。
- 每个新函数先写失败测试，再写最小实现。
- 不提交 git，除非用户明确要求。
- 便携版只装 Outlook/Excel MCP；厚重版额外装 PPT-Master 6.7.3 与归藏网页 PPT。
- 新配置写 OpenCode V2 原生格式；已有 V1 配置按区段保留。
- DeepSeek 模型名不写死；安装时调用 `GET {baseURL}/models`。
- 使用本机资源：PPT-Master 来自 `C:\Users\31177347\AppData\Roaming\Kobold\Storage\PPT-Master\ppt-master\skills\ppt-master`；Outlook/Excel MCP 来自本机更新后的源码；不重新下载。
- Payload 目录：`%TEMP%\opencode-installer-v4`。

---

## File Structure

- `src/install.ps1` — 参数、edition 解析、主流程、安装动作
- `src/lib/installer-core.ps1` — 纯函数（edition、配置、模型、安装计划、zip、版本）
- `src/lib/ui.ps1` — banner、预检表、步骤、报告、确认/密钥输入
- `src/assets/ppt-master-skill/` — 本机 6.7.3（排除 `.env`）
- `src/assets/guizang-ppt-skill/` — 现有快照
- `src/assets/outlook-mcp/` — 本机新版源码
- `src/assets/excel-mcp/` — 本机新版源码
- `src/edition.txt` — 开发默认 `full`
- `tests/installer-core.Tests.ps1` — Pester 单测
- `build/build.ps1` — 双版本构建
- `README.md` — v4 文档

---

### Task 1: Edition model

**Files:**
- Modify: `src/lib/installer-core.ps1`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- `Get-EditionFromFile -Paths <string[]>` → `'portable' | 'full'`（默认 `full`）
- `Get-EditionInfo -Edition <string>` → `@{ Key; Name; Description; ExeName }`
- `Get-EnabledComponents -Edition <string>` → `string[]`
- `Get-RequiredPayloads -Edition <string>` → `string[]`

- [ ] 写失败测试：便携版返回 `outlook, excel`；厚重版返回 `ppt-master, guizang, outlook, excel`；未知 edition 回退 `full`；`Get-RequiredPayloads` 便携版不包含 PPT/归藏 zip。
- [ ] 运行 `Invoke-Pester`，确认因函数不存在而失败。
- [ ] 实现四个函数。
- [ ] 运行 `Invoke-Pester`，确认通过。

### Task 2: Config schema detection and patches

**Files:**
- Modify: `src/lib/installer-core.ps1`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- `Get-ConfigSectionStyle -ConfigJson <string> -Section <'provider'|'mcp'|'permissions'>` → `'V1' | 'V2'`
- `New-ProviderPatch -Style -ProviderId -Name -BaseUrl -ApiKeyRef -Models -MainModel -TitleModel` → JSON string
- `New-McpServerPatch -Style -McpName -Command -Environment -McpTimeout` → JSON string
- `Add-SkillPermissionToJson -ConfigJson -Style` → JSON string
- `Resolve-ConfigValueReference -Value -ConfigDir` → 明文值或 `$null`
- `Get-DefaultComponentSelection -Edition` → `@{ Key; Selected }[]`
- `Test-ComponentEnabled -Edition -Key` → bool

- [ ] 写失败测试：
  - provider：`provider` → V1，`providers` → V2，空 → V2
  - mcp：`mcp.servers` → V2，`mcp.outlook` → V1，空 → V2
  - permissions：`permission` → V1，`permissions` → V2，空 → V2
  - V2 provider patch 含 `providers`、`package`、`settings.baseURL`、`agents.title.model`
  - V1 provider patch 含 `provider`、`npm`、`options`、`small_model`
  - V2 MCP patch 写入 `mcp.servers.outlook`，且保留已有 `mcp.timeout.execution`
  - V1 MCP patch 写入 `mcp.outlook`
  - V2 权限追加到已有 `permissions` 数组且幂等；V1 权限保留其他 `permission` 键
  - `{file:...}`、`{env:...}`、字面量解析
  - 便携版默认只选 Outlook/Excel
- [ ] 运行测试确认失败。
- [ ] 实现函数。
- [ ] 运行测试确认通过。

### Task 3: Model discovery

**Files:**
- Modify: `src/lib/installer-core.ps1`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- `Get-ApiModelsUrl -BaseUrl` → `https://.../models`
- `ConvertFrom-ModelsApiResponse -Json` → `string[]`
- `Select-DefaultModel -Models` → string
- `ConvertTo-ModelDisplayName -ModelId -ProviderName` → string
- `Get-ApiKeyReference -Path` → `{file:...}`

- [ ] 写失败测试：
  - `data:[{id:...}]`、`models:[{id:...}]`、纯数组、空/坏 JSON
  - 去重、过滤空 ID
  - 优先 `deepseek-flash`，其次含 `flash`，否则第一个
  - `deepseek-v4-pro` → `DeepSeek V4 Pro`
  - 路径反斜杠转正斜杠并加 `{file:...}`
- [ ] 运行测试确认失败。
- [ ] 实现函数。
- [ ] 运行测试确认通过。

### Task 4: OpenCode install plan

**Files:**
- Modify: `src/lib/installer-core.ps1`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- `Get-MajorVersion -Version` → int 或 `$null`
- `Get-OpenCodeInstallMethod -Path -NpmPrefix` → `'npm' | 'curl' | 'unknown'`
- `Get-OpenCodeInstallPlan -State <hashtable>` → `@{ Action; Reason; Commands }`
  - State: `@{ Version; Path; NpmPrefix; HasV1; HasV2 }`
  - Actions: `Install`, `MigrateV1`, `UpdateNpm`, `UpdateCurl`, `None`

- [ ] 写失败测试：未安装 → Install；V1 npm → MigrateV1（先 uninstall `opencode-ai` 再 install `@opencode/cli`）；V2 npm → UpdateNpm；V2 curl → UpdateCurl；最新 V2 → None；路径判断。
- [ ] 运行测试确认失败。
- [ ] 实现函数。
- [ ] 运行测试确认通过。

### Task 5: Package helpers

**Files:**
- Modify: `src/lib/installer-core.ps1`
- Test: `tests/installer-core.Tests.ps1`

**Interfaces:**
- `Expand-ZipPackage -ZipPath -Destination` → 解压到目录
- `Get-PackageVersionFromPyProject -Path` → 版本字符串或 `$null`

- [ ] 写失败测试：解压 zip 后文件存在；解析 `pyproject.toml` 的 `[project] version`；缺失/坏文件返回 `$null`。
- [ ] 运行测试确认失败。
- [ ] 实现函数。
- [ ] 运行测试确认通过。

### Task 6: UI and flow

**Files:**
- Create: `src/lib/ui.ps1`
- Rewrite: `src/install.ps1`
- Test: `tests/installer-core.Tests.ps1`（纯逻辑部分）

- [ ] 实现 `ui.ps1`：banner、预检表、计划、步骤、报告、`Confirm-YesNo`、`Read-SecretText`、`Wait-Key`。
- [ ] 重写 `install.ps1`：
  - 读取 edition、计算必需资源
  - 环境预检
  - Node/Python/OpenCode V2 安装
  - DeepSeek/自定义 API 配置（动态模型）
  - 组件菜单按 edition 过滤
  - 技能/MCP 部署（MCP 独立 venv）
  - 一次性写配置（备份 + 合并 + 校验）
  - 完成报告
- [ ] `Invoke-Pester` 全绿。
- [ ] 源码模式 `powershell -File src\install.ps1 -DryRun -TestRoot <temp> -NoPause` 退出码 0。

### Task 7: Assets and build

**Files:**
- Add: `src/assets/ppt-master-skill/`、`outlook-mcp/`、`excel-mcp/`
- Remove: `src/assets/ppt-master-skill-v6.6.0.zip`、`outlook_mcp_server.py`
- Modify: `build/build.ps1`
- Modify: `README.md`

- [ ] 复制本机资源，排除 `.env`、`.git`、`__pycache__`、`dist`、`*.egg-info`、`*.bak`。
- [ ] 构建时按 edition 打包资源。
- [ ] 编译两个 exe，分别做资源校验与 `-DryRun` 冒烟。
- [ ] 更新 README（双版本、动态模型、V2 schema、构建命令、许可）。

### Task 8: Verification

- [ ] `Invoke-Pester -Script tests\installer-core.Tests.ps1` 全绿。
- [ ] `build\build.ps1` 两个版本均构建成功。
- [ ] 两个 exe 的 `-DryRun -TestRoot <temp> -NoPause` 退出码 0。
- [ ] 临时 `XDG_CONFIG_HOME` + `opencode models` 验证 V2 配置被识别。
- [ ] PSScriptAnalyzer（若可用）无 error。
- [ ] 汇总验证输出，不声称未验证的完成。
