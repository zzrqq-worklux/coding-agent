# OpenCode 一键安装器 v3

面向国内零基础用户的 OpenCode 安装器（Windows 10/11，简体中文，管理员/普通用户双模式）。

**内置组件（全部嵌入 exe，单文件外发）：**

| 组件 | 说明 | 版本 | 许可 |
|---|---|---|---|
| [PPT-Master](https://github.com/hugohe3/ppt-master) | 生成原生可编辑 PPTX（PowerPoint 里可继续改） | 6.6.0 | MIT |
| [归藏网页 PPT](https://github.com/op7418/guizang-ppt-skill) | 生成单文件 HTML 幻灯片（浏览器直接放映） | 2025-09 快照 | AGPL-3.0 |
| Outlook 邮件助手（MCP） | 列邮箱 / 搜邮件 / 读邮件 / 发邮件（默认存草稿），需桌面版 Outlook | 本仓库自带 | 私有 |

安装器同时完成：Node.js → Python（可选）→ npm/pip 国内镜像 → OpenCode 安装/更新 → AI 模型配置（DeepSeek / 自定义 OpenAI 兼容）→ 技能与 MCP 部署。

## 目录结构

```
opencode-installer/
├─ src/
│  ├─ install.ps1                 # 主程序（UI 与流程编排）
│  ├─ lib/installer-core.ps1      # 核心逻辑（版本判断/配置合并/解压部署，均有单测）
│  └─ assets/
│     ├─ ppt-master-skill-v6.6.0.zip   # 官方技能包（59.1 MB）
│     ├─ guizang-ppt-skill/            # 归藏技能源目录（构建时打包为 zip）
│     └─ outlook_mcp_server.py         # Outlook MCP 单文件服务
├─ tests/installer-core.Tests.ps1 # Pester 单测（37 个）
├─ build/
│  ├─ build.ps1                   # 构建：单测 → 打包 → 编译 → 资源校验 → 冒烟 →（可选）签名
│  ├─ vendor/ps2exe.ps1           # PS2EXE 编译器（离线构建用，MS-LPL）
│  ├─ vendor/PS2EXE-LICENSE.txt   # PS2EXE 许可全文
│  └─ dist/                       # 构建产物（不入库）
└─ README.md
```

## 构建

前置条件：Windows PowerShell 5.1、.NET Framework（系统自带）、Pester（Windows 自带 3.4.0）。

```powershell
# 普通构建（不签名）
powershell -ExecutionPolicy Bypass -File build\build.ps1

# 构建 + 自签名（生成自签名代码签名证书并签名，导出 .cer）
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign

# 用已有的正式证书签名（买证书后用这个）
powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign -PfxPath D:\cert\my.pfx -PfxPassword xxxx -TimestampUrl http://timestamp.digicert.com
```

产物：`build/dist/OpenCode安装器.exe`（约 64.6 MB）。

构建流程自带三道验证：**37 个单测** → **exe 内嵌资源校验（5 项）** → **冒烟测试（真实运行 exe 的 `-DryRun` 全流程）**，任一失败即中止构建。

## 安装器行为

1. Node.js / Python：优先 winget，失败回退国内镜像安装包
2. npm（npmmirror）/ pip（清华）镜像
3. OpenCode：未装则 `npm install -g opencode-ai`；已装则询问后 `opencode upgrade`（失败回退 npm）
4. AI 模型：DeepSeek（先做 `api.deepseek.com:443` 连通性检查，再验证 Key，401/403/429/网络错误分类提示）或自定义 OpenAI 兼容 API（同样做 host:port 连通性检查）
5. 组件多选菜单（默认全选）：PPT-Master（默认装，随后 `pip install -r requirements.txt`）、归藏网页 PPT、Outlook 邮件助手
6. 更新策略：重跑新版安装器 → 按技能版本号自动覆盖更新（保留用户 `.env`）；OpenCode 用 `opencode upgrade`

配置写入一律**合并式**（保留用户已有的 `provider` / `mcp` 等其他配置，写前自动备份）。

## 调试用参数

```powershell
# 干跑：只显示计划，不修改系统（可用于任何机器上验证）
.\OpenCode安装器.exe -DryRun -NoPause

# 沙箱：把配置/技能目录重定向到指定目录（配合 -DryRun）
.\OpenCode安装器.exe -DryRun -TestRoot C:\temp\sandbox -NoPause

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

- PPT-Master：MIT，随包携带其 `LICENSE` 与署名文件
- 归藏网页 PPT：**AGPL-3.0**，整包原样分发、未做任何修改，随包携带 `LICENSE`
- PS2EXE（构建工具）：Microsoft Limited Public License，许可见 `build/vendor/PS2EXE-LICENSE.txt`
- Python 依赖（python-pptx / PyMuPDF 等）：**不随包分发**，由用户机器 `pip install` 自行安装（其中 PyMuPDF 为 AGPL，注意合规）
