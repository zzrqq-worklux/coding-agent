# outlook-com-mcp

**MCP server that drives your local desktop Outlook via Windows COM — no Azure AD, no Graph API, no cloud auth. Works with on-prem Exchange.**

> 通过本机 Outlook COM 接口的 MCP 服务器：不需要 Azure AD / Graph API / 云端认证，适配私有 Exchange 与桌面版 Outlook。

## 为什么又一个 Outlook MCP？

PyPI 上已有的 `outlook-mcp` 走 Microsoft Graph API，需要企业 Azure AD 注册应用与云权限。
本项目走**本机 COM 接口**，定位完全不同：

| | Outlook MCP (Graph) | **outlook-com-mcp** |
|---|---|---|
| 认证 | Azure AD 应用注册 + 云端权限 | **无需**——复用本机 Outlook 登录态 |
| 网络 | 需访问 Microsoft Graph 云 | **纯本机**（内网/离线可用） |
| 环境 | M365 / Outlook.com | **任意桌面版 Outlook（含企业私有 Exchange）** |
| 真实回复 | — | ✅「真回复」保留会话线程 + 原文引用 |
| 签名 | — | ✅ 自动读取 Outlook 预设签名（如「对外/对内」） |
| 附件 | — | ✅ 可导出邮件附件到本地 |

## 要求

- Windows + 桌面版 Outlook（经典 Outlook，已配置好邮箱账号）
- Python 3.10+

## 安装

```bash
pip install outlook-com-mcp
```

或使用 [uv](https://docs.astral.sh/uv/)：

```bash
uvx outlook-com-mcp
```

## 配置

### opencode

```json
{
  "mcp": {
    "outlook": {
      "type": "local",
      "command": ["python", "-m", "outlook_com_mcp"],
      "environment": {
        "OUTLOOK_MCP_DEFAULT_SIGNATURE": "对外"
      }
    }
  }
}
```

### Claude Desktop

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": ["outlook-com-mcp"]
    }
  }
}
```

## 工具一览

| 工具 | 说明 |
|------|------|
| `list_mailboxes` | 列出全部邮箱账户与收件箱子文件夹（含邮件数） |
| `search_emails` | 按关键词/发件人/日期搜索（支持子文件夹、递归） |
| `read_email` | 按 EntryID 读完整正文；可导出附件与内嵌图片 |
| `send_email` | 起草/发送新邮件（默认只存草稿，防误发） |
| `reply_email` | **真回复**：保留会话线程与原文引用，自动追加预设签名，返回 draft_id |
| `send_draft` | 发送草稿箱中已确认的邮件（配合 reply_email 的 draft_id） |

## 环境变量

| 变量 | 说明 | 默认 |
|------|------|------|
| `OUTLOOK_MCP_DEFAULT_SIGNATURE` | `reply_email` 未显式传 `signature` 时使用的签名名 | 空（不追加签名） |

## 签名机制

`reply_email` 读取本机 Outlook 预设签名文件 `%APPDATA%\Microsoft\Signatures\<名称>.txt`
（Outlook「签名」设置里配置的名称）。若文件不存在则不追加。

> 注意：若你的 Outlook 设置了「回复/转发时自动使用签名」，请显式传 `signature=""` 以免重复。

## 安全说明

- 所有操作复用本机 Outlook 会话，凭据不出本机；
- 邮件发送类接口默认 `send=False`（仅存草稿），需要显式确认后才发送。

## License

MIT License — 详见 [LICENSE](LICENSE)
