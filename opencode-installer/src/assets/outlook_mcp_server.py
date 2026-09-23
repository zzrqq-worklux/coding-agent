"""
outlook-mcp — 本机 Outlook COM 接口 MCP 服务器（私域 Exchange 免云端认证）
Dependencies: pywin32, mcp, fastmcp
Run: python outlook_mcp_server.py
"""
import sys
import pythoncom
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

from fastmcp import FastMCP

mcp = FastMCP("outlook-mcp")


def _get_outlook():
    """返回 (app, ns, owned)。owned=True: 我们新建的实例(结束时可 Quit);
    owned=False: 复用用户正在运行的实例(绝不能 Quit, 否则会关掉用户的 Outlook)。"""
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("Outlook.Application")
        return app, app.GetNamespace("MAPI"), False
    except Exception:
        pass
    # 无运行实例时新建(防止干扰用户界面, 用 DispatchEx 隔离)
    app = win32com.client.DispatchEx("Outlook.Application")
    return app, app.GetNamespace("MAPI"), True


def _safe_quit(app, owned):
    if owned and app is not None:
        try:
            app.Quit()
        except Exception:
            pass


def _gencache_outlook():
    """早绑定 Outlook 应用(DispatchEx 创建隔离实例, 避免干扰用户界面外窗)。"""
    import win32com.client
    return win32com.client.gencache.EnsureDispatch("Outlook.Application")


def _unwrap(value):
    if value is None:
        return ""
    try:
        return value.strip()
    except AttributeError:
        return str(value)


def _item_to_dict(item, include_body=False, body_limit=4000):
    ts = item.ReceivedTime
    body = ""
    if include_body:
        try:
            body = _unwrap(item.Body)
            if len(body) > body_limit:
                body = body[:body_limit] + f"\n...[截断,共 {len(item.Body)} 字符]"
        except Exception:
            body = ""
    return {
        "id": str(item.EntryID),
        "subject": _unwrap(item.Subject),
        "sender": _unwrap(item.SenderName) if hasattr(item, "SenderName") else "",
        "sender_email": _unwrap(item.SenderEmailAddress) if hasattr(item, "SenderEmailAddress") else "",
        "received": ts.strftime("%Y-%m-%d %H:%M") if ts else "",
        "folder": "",
        "body": body,
    }


def _folder_by_name(ns, folder_path):
    """按名称/路径找文件夹; 支持 '收件箱/项目相关' 式路径或在任意位置匹配同名子文件夹。"""
    parts = [p.strip() for p in folder_path.split("/") if p.strip()]

    # 快速路径: 收件箱/已发送/日历等默认文件夹 — GetDefaultFolder(6)=收件箱
    default_map = {"收件箱": 6, "inbox": 6, "已发送邮件": 5, "sent items": 5,
                   "发件箱": 4, "outbox": 4, "已删除邮件": 3, "deleted items": 3,
                   "日历": 9, "calendar": 9, "任务": 13, "tasks": 13,
                   "草稿": 16, "drafts": 16}
    if len(parts) == 1 and parts[0].lower() in default_map:
        try:
            d = ns.GetDefaultFolder(default_map[parts[0].lower()])
            if d is not None:
                return d
        except Exception:
            pass

    def walk(folders, depth=0, maxd=4):
        for i in range(1, folders.Count + 1):
            f = folders.Item(i)
            if f.Name.lower() == parts[0].lower():
                if len(parts) == 1:
                    return f
                return _walk_path(f, parts[1:])
            if depth < maxd:
                res = _walk_sub(f, depth + 1)
                if res:
                    return res
        return None

    def _walk_sub(folder, depth):
        for i in range(1, folder.Folders.Count + 1):
            f = folder.Folders.Item(i)
            try:
                if f.Name.lower() == parts[0].lower():
                    if len(parts) == 1:
                        return f
                    return _walk_path(f, parts[1:])
                if depth < 4:
                    res = _walk_sub(f, depth + 1)
                    if res:
                        return res
            except Exception:
                continue
        return None

    def _walk_path(folder, rest):
        for i in range(1, folder.Folders.Count + 1):
            f = folder.Folders.Item(i)
            if f.Name.lower() == rest[0].lower():
                if len(rest) == 1:
                    return f
                return _walk_path(f, rest[1:])
        return None

    return walk(ns.Folders)


def _iter_subfolders(folder, include_self=True, maxdepth=4):
    """yield 所有子文件夹(含自身), 深度限 maxdepth 防循环。"""
    if include_self:
        yield folder
    if maxdepth <= 0:
        return
    for i in range(1, folder.Folders.Count + 1):
        try:
            child = folder.Folders.Item(i)
        except Exception:
            continue
        yield from _iter_subfolders(child, include_self=True, maxdepth=maxdepth - 1)


@mcp.tool()
def list_mailboxes() -> list:
    """列出本机 Outlook 已配置的全部邮箱账户及收件箱子文件夹(含邮件条数)。"""
    try:
        app, ns, owned = _get_outlook()
        results = []
        for i in range(1, ns.Folders.Count + 1):
            acct = ns.Folders.Item(i)
            subs = []
            for f in _iter_subfolders(acct, maxdepth=2):
                try:
                    cnt = f.Items.Count
                except Exception:
                    cnt = 0
                subs.append({"folder": f.Name, "items": cnt})
            results.append({"account": acct.Name, "folders": subs})
        return results
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def search_emails(folder: str = "收件箱", keyword: str = "", sender: str = "",
                  date_from: str = "", date_to: str = "", limit: int = 20,
                  include_body: bool = True, body_limit: int = 4000,
                  recursive: bool = True) -> list:
    """按关键词/发件人/日期范围搜索 Outlook 邮件(默认递归收件箱全部子文件夹)。

    Args:
        folder: 文件夹名(支持'收件箱/项目相关'路径; 默认 收件箱; 常用子文件夹: 周报/会议/JUA/杂项/扫描件/发票; 存档)
        keyword: 主题+正文关键词(模糊, 不区分大小写); 可多词用空格分隔(全部命中才返回)
        sender: 发件人姓名或邮箱(模糊匹配)
        date_from: 起始日期 YYYY-MM-DD
        date_to: 结束日期 YYYY-MM-DD(含当日)
        limit: 最多返回条数
        include_body: 是否含正文(只取前 body_limit 字符)
        body_limit: 正文截断长度
        recursive: 递归搜索 folder 下所有子文件夹(默认 True)
    """
    try:
        app, ns, owned = _get_outlook()
        folder_obj = _folder_by_name(ns, folder)
        if folder_obj is None:
            raise ValueError(f"未找到文件夹: {folder}")

        targets = list(_iter_subfolders(folder_obj, include_self=True, maxdepth=3)) if recursive else [folder_obj]

        def any_word_in(fulltext, words):
            return all(w.lower() in fulltext.lower() for w in words)

        result = []
        for fo in targets:
            if len(result) >= limit:
                break
            items = fo.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            filters = []
            if date_from:
                filters.append(f'"[ReceivedTime] >= "{date_from} 00:00"')
            if date_to:
                filters.append(f'"[ReceivedTime] <= "{date_to} 23:59"')
            filtered = items.Restrict(" AND ".join(filters)) if filters else items
            for item in filtered:
                try:
                    if item.Class != 43:
                        continue
                    subj = _unwrap(item.Subject)
                    body = _unwrap(item.Body) if (keyword or include_body) else ""
                    if keyword and not any_word_in(subj + " " + body, keyword.split()):
                        continue
                    if sender:
                        sname = _unwrap(item.SenderName) if hasattr(item, "SenderName") else ""
                        semail = _unwrap(item.SenderEmailAddress) if hasattr(item, "SenderEmailAddress") else ""
                        if sender.lower() not in (sname + " " + semail).lower():
                            continue
                    d = _item_to_dict(item, include_body, body_limit)
                    d["folder"] = fo.Name if fo.Parent else fo.Name
                    result.append(d)
                    if len(result) >= limit:
                        break
                except Exception:
                    continue
        result.sort(key=lambda x: x.get("received", ""), reverse=True)
        return result[:limit]
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def read_email(item_id: str, body_limit: int = 10000, save_attachments: bool = False,
               save_dir: str = "") -> dict:
    """按 EntryID 读取单封邮件完整正文(可选导出附件)。

    Args:
        item_id: 邮件 EntryID(来自 search_emails 结果)
        body_limit: 正文截断长度
        save_attachments: True 时将附件/内嵌图片另存到本地(不入邮件目录)
        save_dir: 附件保存目录(默认 %TEMP%/opencode/outlook-attachments)
    """
    import os
    try:
        app, ns, owned = _get_outlook()
        try:
            item = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 Item ID(可能邮件已不在本机或已删除)")

        d = _item_to_dict(item, include_body=True, body_limit=body_limit)
        try:
            to = []
            for r in item.Recipients:
                to.append(_unwrap(r.Name))
            d["to"] = ", ".join(to)
        except Exception:
            d["to"] = ""
        try:
            atts = []
            for a in item.Attachments:
                try:
                    atts.append({
                        "name": _unwrap(a.FileName) or "(inline)",
                        "index": a.Index,
                    })
                except Exception:
                    pass
            d["attachments"] = atts
        except Exception:
            d["attachments"] = []

        if save_attachments and d["attachments"]:
            if not save_dir:
                save_dir = os.path.join(os.environ.get("TEMP", r"C:\Users\31177347\AppData\Local\Temp"),
                                        "opencode", "outlook-attachments")
            os.makedirs(save_dir, exist_ok=True)
            saved = []
            for a in item.Attachments:
                try:
                    name = _unwrap(a.FileName) or f"inline_{a.Index}.bin"
                    safe = "".join(c for c in name if c not in '\\/:*?"<>|') or "attachment"
                    path = os.path.join(save_dir, safe)
                    # 防重复同名自动加序号
                    n = 1
                    base, ext = os.path.splitext(path)
                    while os.path.exists(path):
                        path = f"{base}_{n}{ext}"
                        n += 1
                    a.SaveAsFile(path)
                    size = os.path.getsize(path)
                    is_pic = any(name.lower().endswith(ext2) for ext2 in
                                 (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff"))
                    saved.append({"name": name, "path": path, "size": size, "is_pic": is_pic})
                except Exception as e:
                    saved.append({"name": name, "error": str(e)})
            d["saved_attachments"] = saved
            d["attachment_dir"] = save_dir
        return d
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def send_email(to: str, subject: str, body: str, cc: str = "", bcc: str = "",
               attachments: list = None, send: bool = False,
               use_html: bool = False, signature: bool = True) -> dict:
    """起草或发送邮件(经本机 Outlook 账号, 走公司 Exchange, 无需额外认证)。

    Args:
        to: 收件人(多个用英文分号 ; 分隔)
        subject: 主题
        body: 正文
        cc: 抄送(可选; 英文分号分隔)
        bcc: 密送(可选; 英文分号分隔)
        attachments: 本地附件文件路径列表(可选)
        send: True=直接发送; False=仅保存到草稿(默认, 避免误发)
        use_html: True=正文按 HTML 渲染(可含 <br>/<b> 等标签); False=纯文本
        signature: 是否带默认签名(默认 True)
    """
    import os
    try:
        app, ns, owned = _get_outlook()
        item = app.CreateItem(0)  # olMailItem
        item.To = to
        item.CC = cc or ""
        item.BCC = bcc or ""
        item.Subject = subject
        if use_html:
            item.HTMLBody = body
        else:
            item.Body = body
        for path in (attachments or []):
            if os.path.exists(path):
                item.Attachments.Add(path)
        if send:
            item.Send()
            return {"status": "sent", "to": to, "subject": subject}
        item.Save()
        return {"status": "draft_saved", "to": to, "subject": subject,
                "hint": "已保存到草稿文件夹; 如需直接发送请将 send 设为 True"}
    finally:
        _safe_quit(app, owned)


if __name__ == "__main__":
    mcp.run(transport="stdio")
