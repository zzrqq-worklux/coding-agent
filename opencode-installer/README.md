# OpenCode 一键安装器 v4.2.2

面向国内零基础用户的 OpenCode 安装器（Windows 10/11，简体中文，管理员/普通用户双模式）。

## 两个版本

| 版本 | 内置组件 | 产物 |
|---|---|---|
| **便携版** | Outlook 邮件助手 + Excel 表格助手（MCP） | `OpenCode安装器-便携版.exe`（约 0.1 MB） |
| **厚重版** | PPT-Master 6.7.3 + 归藏网页 PPT + Outlook + Excel | `OpenCode安装器-厚重版.exe`（约 27 MB） |

两个版本都包含完整核心流程：Node.js → Python（可选）→ npm/pip 国内镜像 → OpenCode V2 → DeepSeek 动态模型 → 配置写入。

## v4 重点变化

- **OpenCode V2**：安装 `@opencode/cli`；检测到 V1 `opencode-ai` 时先卸载再迁移到 V2。
- **DeepSeek 动态模型**：安装时调用 `GET https://api.deepseek.com/v1/models`，不再写死模型名；接口不可用时可手动输入模型 ID。
- **V2 原生配置**：新配置写 `providers` / `settings` / `mcp.servers` / `permissions` / `agents.title.model`。
- **兼容已有 V1 配置**：按区段检测 `provider` vs `providers`、`mcp.<name>` vs `mcp.servers`、`permission` vs `permissions`，保留原格式，避免破坏旧配置。
- **MCP 独立环境**：Outlook / Excel 分别安装到 `~/.config/opencode/mcp-servers/venv-outlook` / `venv-excel`，配置指向各自的 venv python，避免两个 MCP 的 `mcp` / `fastmcp` 版本互相破坏，也不污染全局 Python。
- **Excel MCP 复用本机名称**：配置项使用 `my-excel`，不会和已有的 Excel MCP 重复。
- **TUI 重做**：环境预检、执行计划确认、分步进度、最终报告、转录日志。
- **内置组件更新**：PPT-Master 升到 6.7.3（本机 Kobold 副本打包，排除 `.env`）；Outlook / Excel MCP 使用本机更新后的源码。

## v4.2.2 my-excel MCP 启动修复

- **根因**：Excel MCP 的依赖写成不限制版本的 `mcp`，新装会拿到 mcp 2.x；mcp 2 已删除 `mcp.server.fastmcp`，`my-excel` 启动即 `ModuleNotFoundError`，OpenCode 显示 `MCP server failed: my-excel`。
- **依赖固定**：Excel MCP 改为 `mcp>=1.24,<2`，并移除未使用的 `fastmcp` 依赖；Outlook MCP 固定 `fastmcp>=2.0,<5`。
- **每个 MCP 独立 venv**：`mcp-servers/venv-excel` 与 `mcp-servers/venv-outlook` 分开，Excel 使用 mcp 1.x、Outlook 使用 FastMCP 4.x 也不会互相降级/破坏。
- **真实启动探针**：安装后导入 `excel_mcp.server` / `outlook_com_mcp.server`，不再只导入空的包 `__init__`；探针失败会强制重装依赖，仍失败则跳过该 MCP 的配置更新（已有配置保持不变）。
- **升级不再复用旧源码**：只要内嵌 zip 存在就重新解压到 `%TEMP%`，避免老用户重跑时继续使用上一版未修复的 `excel_mcp` 代码。
- **旧用户修复**：重跑安装器即可；确认两个新 venv 正常后，可手动删除旧的共享目录 `~/.config/opencode/mcp-servers/venv`。
- 单测增加到 177 个（Windows PowerShell 5.1 下全部通过）。

## v4.2.1 安装计划与输入修复

- **不再重复安装已是最新的 OpenCode**：安装计划会查询 npm 最新版；已是最新版时直接跳过，不再执行 `opencode upgrade` / npm 安装。查询失败时保留现有版本，避免网络异常导致重装卡住。
- **npm 更新改为直接执行**：npm 安装的 OpenCode 不再先跑 `opencode upgrade --method npm`，避免自升级进程和回退 npm 互相等待；npm 安装带超时、进度提示和失败日志。
- **检测到 OpenCode 正在运行**：Windows 无法替换正在运行的 exe，安装器会提示先关闭 OpenCode，并跳过本次 OpenCode 更新，继续安装其他组件。
- **中文输入法 / 全角输入兼容**：AI 服务选择、模型选择、组件选择、Y/N 确认统一做去空白、去控制字符、全角转半角；输入 1 不会再误进“其他 OpenAI 兼容 API”。
- **组件菜单反馈**：按 `a` / `n` 后明确提示“再按一次回车开始”，避免全选后看起来没反应。
- 单测增加到 170 个（Windows PowerShell 5.1 下全部通过）。

## v4.2.0 快捷方式、DeepSeek 费用与 Claude 风格教程

- **桌面快捷方式**：安装完成自动创建 `OpenCode.lnk`，指向 `%LOCALAPPDATA%\OpenCode\OpenCode.cmd`。
- **拖拽启动**：双击从桌面启动；把任意文件夹/文件拖到快捷方式上，就从该目录启动 opencode。
- **DeepSeek 费用说明**：教程补充官方人民币价格、高峰/空闲时段、注册和支付宝/微信充值流程。
- **Claude 风格教程**：暖米白底色、陶土橙强调色、衬线标题、暖灰卡片；重新导出两份 PDF。
- 单测增加到 153 个，构建继续自动执行 `-SelfTest` 和发布打包。

## v4.1.0 PowerShell 7 与网络兜底

- **可选安装 PowerShell 7**：默认询问，已有 `pwsh` 则跳过；检测到 Git Bash 时直接跳过。
- **纯用户目录安装**：下载官方 ZIP 解压到 `%LOCALAPPDATA%\Programs\PowerShell\7`，只写用户 PATH，不调用 winget/MSI，不触发 UAC。
- **下载兜底**：本地 ZIP → GitHub → 清华 TUNA 镜像 → `OC_POWERSHELL_URL` 自定义镜像；全部失败则跳过，不影响 OpenCode。
- **SHA256 校验**：固定 PowerShell 7.6.9，所有来源下载后校验 `CE7484804C0F0AD4DDA10C958E71ACDAF7EEBFB22505C1B480F89E0E9F1DA94D`，不匹配自动换源。
- **网络提示**：开始前提示建议使用手机热点，避免公司网络拦截 npm/GitHub/镜像。
- **启动方式**：完成页检测到 `pwsh` 时优先用 PowerShell 7 启动 opencode，否则回退 `cmd`。
- 流程从 7 步调整为 8 步：Node → Python → PowerShell 7 → OpenCode → AI → 组件 → 配置 → 完成。

## v4.0.1 健壮性修复

- **npm 命令固定走 npm.cmd**：避免 PowerShell 把 `npm` 解析成 `npm.ps1`。`npm.ps1` 会用 `Invoke-Expression` 重新执行原始命令文本，导致 `$script:NpmMirror` 等变量在 npm.ps1 作用域里变成未定义。
- **启动自检**：`-SelfTest` 检查 PowerShell 版本、核心组件版本、核心函数、npm 命令类型、Node/Python/OpenCode 和代理配置；构建时自动执行。
- **核心版本校验**：`install.ps1` 与 `installer-core.ps1` 版本不一致时直接给出明确错误，避免加载 `%TEMP%` 里的旧核心。
- **优先本地核心**：从源码运行时优先加载 `src/lib`，不再被 `%TEMP%\opencode-installer-v4` 的旧文件覆盖。
- **下载诊断**：Node/Python 下载前检查镜像连通性，支持 `HTTPS_PROXY` / `HTTP_PROXY` / 系统代理，失败时输出具体 URL 和异常。
- **自动打包**：构建完成后同步生成 `build/release/opencode安装器-*.zip`，避免发布旧 exe。

## 目录结构

```
opencode-installer/
├─ src/
│  ├─ install.ps1                 # 主流程（环境、安装、配置、组件）
│  ├─ edition.txt                 # 源码运行默认 full；构建时被 edition-*.txt 覆盖
│  ├─ lib/
│  │  ├─ installer-core.ps1       # 纯逻辑（版本、schema、patch、模型、安装计划）
│  │  └─ ui.ps1                   # TUI 渲染
│  └─ assets/
│     ├─ ppt-master-skill-v6.7.3.zip   # 本机 6.7.3，已排除 .env
│     ├─ guizang-ppt-skill/            # 归藏技能源目录（构建时打包）
│     ├─ outlook-mcp/                  # 本机新版 Outlook MCP 源码
│     └─ excel-mcp/                    # 本机新版 Excel MCP 源码
├─ tests/installer-core.Tests.ps1 # Pester 单测（177 个）
├─ build/
│  ├─ build.ps1                   # 单测 → 打包 → 编译 → 资源校验 → 冒烟 →（可选）签名
│  ├─ vendor/ps2exe.ps1           # PS2EXE 编译器（离线构建用，MS-LPL）
│  ├─ vendor/PS2EXE-LICENSE.txt
│  └─ dist/                       # 构建产物（不入库）
├─ docs/superpowers/plans/        # 实施计划
└─ README.md
```

## 构建

前置条件：Windows PowerShell 5.1、.NET Framework（系统自带）、Pester（Windows 自带 3.4.0）。

```powershell
# 构建两个版本（默认）
powershell -ExecutionPolicy Bypass -File build\build.ps1

# 只构建便携版
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Edition portable

# 只构建厚重版
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Edition full

# 构建 + 自签名
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign

# 用正式证书签名
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign -PfxPath D:\cert\my.pfx -PfxPassword xxxx -TimestampUrl http://timestamp.digicert.com
```

产物：

- `build/dist/OpenCode安装器-便携版.exe`
- `build/dist/OpenCode安装器-厚重版.exe`
- `build/release/opencode安装器-便携版.zip`
- `build/release/opencode安装器-厚重版.zip`

构建流程自带验证：**177 个单测** → **资源打包** → **exe 编译** → **内嵌资源校验** → **`-DryRun` 冒烟测试** → **`-SelfTest` 自检** → **发布目录/zip 打包**，任一失败即中止构建。`-NoPackage` 可跳过最后的发布打包。

## 安装器行为

1. **环境预检**：系统、管理员/普通用户、Node、Python、OpenCode 版本与安装方式、现有配置格式、安装目标。
2. **Node.js LTS**：winget 优先，失败回退 npmmirror MSI。
3. **Python 3.12**：winget 优先，失败回退国内镜像安装包（PPT 依赖和 MCP 需要）。
4. **OpenCode V2**：
   - 未安装 → `npm install -g @opencode/cli@latest`
   - V1 → 卸载 `opencode-ai`，安装 `@opencode/cli`
   - V2 npm 安装 → 查询最新版；需要更新时直接 `npm install -g @opencode/cli@latest`（带超时/进度）
   - V2 curl 安装 → `opencode upgrade`
5. **AI 模型**：
   - DeepSeek：输入 Key（掩码）→ 调 `/models` 动态获取模型 → 选择默认模型 → Key 存到 `~/.config/opencode/.secrets/deepseek-api-key`，配置用 `{file:...}` 引用。
   - 自定义 OpenAI 兼容 API：同样先拉 `/models`；失败可手动输入模型 ID。
   - 模型列表完全来自接口；接口不可用时才要求手动输入，不内置写死的模型名。
6. **组件菜单**：按 edition 过滤；便携版只有 Outlook / Excel，厚重版四个组件全显示。
7. **技能部署**：`~/.config/opencode/skills`，保留用户 `.env`，写版本清单。
8. **MCP 部署**：每个 MCP 独立 venv（`mcp-servers/venv-outlook` / `venv-excel`），`pip install` 本地包（清华源，失败回退默认源），配置指向对应 venv python；安装后导入 `*.server` 做启动探针；Excel 使用 `my-excel` 名称。
9. **配置写入**：先备份，再深度合并，最后校验 JSON；只写一次，避免半成品配置。
10. **完成报告**：每个组件成功/失败、配置路径、日志路径、下一步。

## 配置与密钥

- 新配置写 OpenCode V2 原生格式；已有 V1 配置按区段保留。
- V2 权限是 `permissions` 数组，安装器只**追加**技能规则，不覆盖已有规则。
- V1 配置下写 `permission.skill.*`，保持与旧格式一致。
- DeepSeek Key 单独存放在 `~/.config/opencode/.secrets/deepseek-api-key`（隐藏文件），配置文件只保存 `{file:...}` 引用；不要分享 `.secrets` 目录。
- 如果已有配置不是合法 JSON，安装器会询问是否备份并重建。
- 配置写入前会自动备份为 `opencode.json.backup-yyyyMMddHHmmss`。

## 调试参数

```powershell
# 自检：检查核心版本、npm 命令、Node/Python/OpenCode、代理（不修改系统）
.\OpenCode安装器-厚重版.exe -SelfTest -NoPause

# 干跑：只显示计划，不修改系统（可用于任何机器上验证）
.\OpenCode安装器-厚重版.exe -DryRun -NoPause

# 沙箱：把配置/技能/MCP 目录重定向到指定目录（配合 -DryRun）
.\OpenCode安装器-厚重版.exe -DryRun -TestRoot C:\temp\sandbox -NoPause

# 源码运行（需要 -PayloadRoot 指向 build\dist 才能找到打包资源）
powershell -ExecutionPolicy Bypass -File src\install.ps1 -DryRun -TestRoot C:\temp\sandbox -PayloadRoot build\dist -NoPause

# 单测
Invoke-Pester -Script tests\installer-core.Tests.ps1
```

## 签名与 SmartScreen（务必阅读）

- **自签名不会消除陌生外网用户看到的 SmartScreen 提示**——对未导入证书的机器，效果与未签名相同（微软 2024 年后不再为 EV/OV 提供即时信誉豁免）。自签名的实际用途：① 企业内由 IT 导入 `.cer` 形成信任；② 配合 AppLocker/WDAC 按发布者放行；③ 把签名流水线先跑通，买到正式证书后一条命令即可切换。
- 正式解法：购买 OV 代码签名证书后用上面的 `-PfxPath` 方式签名，并持续用同一张证书发布，SmartScreen 信誉会随下载量累积。
- 企业内信任自签名证书（管理员在目标机器执行）：
  ```powershell
  Import-Certificate -FilePath OpenCode安装器-签名证书.cer -CertStoreLocation Cert:\LocalMachine\TrustedPublisher
  Import-Certificate -FilePath OpenCode安装器-签名证书.cer -CertStoreLocation Cert:\LocalMachine\Root
  ```
- **发布注意**：上传到站点/网盘时请从构建机直接复制文件。若先把 exe 用浏览器下载一次再上传，会带上「网络来源」标记（MOTW），所有用户都会收到额外拦截提示。

## 第三方组件与许可

- **PPT-Master 6.7.3**：MIT，随包携带其 `LICENSE` 与署名文件；打包时已排除本机 `.env`。
- **归藏网页 PPT**：**AGPL-3.0**，整包原样分发、未做任何修改，随包携带 `LICENSE`。
- **Outlook MCP（outlook-com-mcp）**：MIT，随包携带 `LICENSE`。
- **Excel MCP（excel-mcp-server-selfhosted）**：本机私有自托管分支，未公开发布；随包携带其 `pyproject.toml` 与源码。
- **PS2EXE（构建工具）**：Microsoft Limited Public License，许可见 `build/vendor/PS2EXE-LICENSE.txt`。
- **Python 依赖**（python-pptx / PyMuPDF / fastmcp / pywin32 / openpyxl 等）：**不随包分发**，由用户机器 `pip install` 自行安装（其中 PyMuPDF 为 AGPL，注意合规）。
