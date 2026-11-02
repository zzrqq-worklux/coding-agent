"""
outlook-com-mcp — 本机 Outlook COM 接口 MCP 服务器（私域 Exchange 免云端认证）

通过 Windows COM 直接驱动本机桌面版 Outlook，无需 Azure AD / Graph API / 云端认证，
适配企业内网 Exchange 与桌面 Outlook 工作流。

Dependencies: pywin32, fastmcp
Run: python -m outlook_com_mcp
"""
import os
import sys
import pythoncom
pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

from fastmcp import FastMCP

mcp = FastMCP("outlook-com-mcp")


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


def _load_signature(name):
    """读取 Outlook 预设签名文件(%APPDATA%\\Microsoft\\Signatures\\<名称>.txt)。
    常见名称: 对外 / 对内; 不存在返回 None。"""
    import os
    if not name:
        return None
    path = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures", name + ".txt")
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:2] == b"\xff\xfe":
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("gbk", errors="replace")


def _load_signature_html(name):
    """读取 Outlook 预设签名的 HTML 版(%APPDATA%\\Microsoft\\Signatures\\<名称>.htm)。

    返回 (html, logos): html 中图片位置替换为占位符 "%%SIGLOGO%%"(多图依次编号);
    logos = [(占位符, 图片本地路径)] — 供 Word 编辑器通道插入为隐藏内嵌图。
    文件不存在或读取失败返回 (None, [])。
    """
    import re
    if not name:
        return None, []
    base = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures")
    path = os.path.join(base, name + ".htm")
    if not os.path.exists(path):
        return None, []
    with open(path, "rb") as f:
        raw = f.read()
    text = None
    for enc in ("gb2312", "gbk", "utf-8", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if text is None:
        return None, []
    styles = "".join(m.group(0) for m in re.finditer(r"<style[^>]*>.*?</style>", text, re.S | re.I))
    m = re.search(r"<body[^>]*>(.*)</body>", text, re.S | re.I)
    inner = m.group(1) if m else text
    # 去掉 VML 条件块, 保留 <img> 回退形式(与 Outlook 自身插入签名的方式一致)
    inner = re.sub(r"<!--\[if gte vml 1\]>.*?<!\[endif\]-->", "", inner, flags=re.S)
    inner = re.sub(r"<!\[if !vml\]>(.*?)<!\[endif\]>", r"\1", inner, flags=re.S)
    # 图片引用 -> 占位符(稍后由 WordEditor 替换为隐藏内嵌图)
    logos = []
    folder = name + ".files/"

    def _img_to_marker(mo):
        tag = mo.group(0)
        fm = re.search(re.escape(folder) + r"([^\"'>\s]+)", tag)
        if not fm:
            return tag
        img_path = os.path.join(base, name + ".files", fm.group(1))
        if not os.path.exists(img_path):
            return tag
        marker = "%%SIGLOGO" + (str(len(logos)) if logos else "") + "%%"
        logos.append((marker, img_path))
        return marker

    inner = re.sub(r"<img[^>]*" + re.escape(folder) + r"[^>]*>", _img_to_marker, inner, flags=re.I)
    return styles + inner, logos


def _build_signature(name):
    """构建签名块: 返回 (html, logos)。优先 HTML 版签名, 缺失时回退 .txt。"""
    import html as _html
    if not name:
        return None, []
    sig_html, sig_logos = _load_signature_html(name)
    if not sig_html:
        sig_txt = _load_signature(name)
        if sig_txt:
            sig_html = ('<div style="font-family:等线;font-size:10.5pt;">' +
                        _html.escape(sig_txt.strip()).replace("\r\n", "<br>") + '</div>')
    return sig_html, sig_logos


def _embed_signature_logos(item, sig_logos):
    """将签名 logo 占位符经 Word 编辑器通道替换为"隐藏"内嵌图(Outlook 原生行为,
    与用户手动插入签名的效果一致; 直接以附件+CID 方式添加的图会显示为可见附件)。

    返回状态: inserted / partial / not_found / failed; 失败时清除占位符(保留文字签名)。
    """
    inserted = 0
    insp = None
    status = "failed"
    try:
        insp = item.GetInspector
        insp.Display()
        doc = insp.WordEditor
        for marker, logo_path in sig_logos:
            rng = doc.Content
            f = rng.Find
            f.ClearFormatting()
            f.Text = marker
            f.Forward = True
            f.Wrap = 0
            if f.Execute():
                rng.Text = ""
                doc.InlineShapes.AddPicture(logo_path, False, True, rng)
                inserted += 1
        item.Save()
        if inserted == len(sig_logos):
            status = "inserted"
        elif inserted:
            status = "partial"
        else:
            status = "not_found"
    except Exception:
        status = "failed"
    finally:
        try:
            if insp is not None:
                insp.Close(0)
        except Exception:
            pass
    if inserted < len(sig_logos):
        try:
            h2 = item.HTMLBody or ""
            for marker, _ in sig_logos:
                h2 = h2.replace(marker, "")
            item.HTMLBody = h2
        except Exception:
            pass
    return status


def _validate_attachments(attachments):
    """附件路径必须存在; 缺失直接报错, 不静默跳过。"""
    missing = [p for p in (attachments or []) if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError("附件不存在: " + "; ".join(missing))
    return attachments or []


def _item_to_dict(item, include_body=False, body_limit=4000, include_html=False):
    try:
        ts = item.ReceivedTime
        received = ts.strftime("%Y-%m-%d %H:%M") if ts else ""
    except Exception:
        received = ""
    body = ""
    if include_body:
        try:
            body = _unwrap(item.Body)
            if len(body) > body_limit:
                body = body[:body_limit] + f"\n...[截断,共 {len(item.Body)} 字符]"
        except Exception:
            body = ""
    out = {
        "id": str(item.EntryID),
        "subject": _unwrap(item.Subject),
        "sender": _unwrap(item.SenderName) if hasattr(item, "SenderName") else "",
        "sender_email": _unwrap(item.SenderEmailAddress) if hasattr(item, "SenderEmailAddress") else "",
        "received": received,
        "folder": "",
        "body": body,
    }
    if include_html:
        try:
            html_body = _unwrap(item.HTMLBody) or ""
            if len(html_body) > body_limit:
                html_body = html_body[:body_limit] + f"\n...[截断,共 {len(item.HTMLBody)} 字符]"
            out["body_html"] = html_body
        except Exception:
            out["body_html"] = ""
    return out


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
                  recursive: bool = True, offset: int = 0) -> list:
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
        offset: 跳过前 N 条(分页用; 结果按收到时间倒序)
    """
    try:
        app, ns, owned = _get_outlook()
        folder_obj = _folder_by_name(ns, folder)
        if folder_obj is None:
            raise ValueError(f"未找到文件夹: {folder}")

        targets = list(_iter_subfolders(folder_obj, include_self=True, maxdepth=3)) if recursive else [folder_obj]

        def any_word_in(fulltext, words):
            return all(w.lower() in fulltext.lower() for w in words)

        cap = max(limit, 0) + max(offset, 0)
        result = []
        for fo in targets:
            if len(result) >= cap:
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
                    if len(result) >= cap:
                        break
                except Exception:
                    continue
        result.sort(key=lambda x: x.get("received", ""), reverse=True)
        return result[offset:offset + limit]
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def read_email(item_id: str, body_limit: int = 10000, save_attachments: bool = False,
               save_dir: str = "", include_html: bool = False) -> dict:
    """按 EntryID 读取单封邮件完整正文(可选导出附件)。

    Args:
        item_id: 邮件 EntryID(来自 search_emails 结果)
        body_limit: 正文截断长度
        save_attachments: True 时将附件/内嵌图片另存到本地(不入邮件目录)
        save_dir: 附件保存目录(默认 %TEMP%/opencode/outlook-attachments)
        include_html: True 时额外返回 body_html(保留表格/排版; 用 body_limit 截断)
    """
    import os
    try:
        app, ns, owned = _get_outlook()
        try:
            item = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 Item ID(可能邮件已不在本机或已删除)")

        d = _item_to_dict(item, include_body=True, body_limit=body_limit, include_html=include_html)
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
                import tempfile
                save_dir = os.path.join(tempfile.gettempdir(), "opencode", "outlook-attachments")
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
               use_html: bool = False, signature: str = "对外") -> dict:
    """新建邮件（非回复；回复请用 reply_email，且使用前先与用户确认原邮件）。

    Args:
        to: 收件人(多个用英文分号 ; 分隔)
        subject: 主题
        body: 正文(纯文本, 自动转 HTML 排版; use_html=True 时按 HTML 原样)
        cc: 抄送(可选; 英文分号分隔)
        bcc: 密送(可选; 英文分号分隔)
        attachments: 本地附件文件路径列表(可选, 可见附件)
        send: True=直接发送; False=仅保存草稿(默认, 避免误发)
        use_html: True=正文按 HTML 渲染(可含 <br>/<b> 等标签); False=纯文本
        signature: 签名名称(如 "对外"/"对内"; 优先 HTML 版, 含 logo 自动内嵌为
            "隐藏"图片; 无 HTML 版时回退 .txt); 默认 "对外"; 传空串/False 则不加签名
    """
    import html as _html
    try:
        app, ns, owned = _get_outlook()
        item = app.CreateItem(0)  # olMailItem
        item.To = to
        item.CC = cc or ""
        item.BCC = bcc or ""
        item.Subject = subject

        if use_html:
            block = body or ""
        else:
            text = (body or "").rstrip()
            body_html = _html.escape(text).replace("\r\n", "<br>").replace("\n", "<br>")
            block = ('<div style="font-family:等线;font-size:10.5pt;line-height:1.4;">'
                     + body_html + '</div>')

        if signature is True:
            sig_name = "对外"
        elif not signature:
            sig_name = ""
        else:
            sig_name = str(signature).strip()
        sig_logos = []
        if sig_name:
            sig_html, sig_logos = _build_signature(sig_name)
            if sig_html:
                block = block + "<br>" + sig_html

        item.HTMLBody = "<html><body>" + block + "</body></html>"

        # 签名 logo: 经 Word 编辑器通道插入为隐藏内嵌图
        logo_status = "none"
        if sig_logos:
            logo_status = _embed_signature_logos(item, sig_logos)

        for path in _validate_attachments(attachments):
            item.Attachments.Add(path)

        if send:
            item.Save()
            item.Send()
            return {"status": "sent", "to": to, "subject": subject, "logo": logo_status}
        item.Save()
        return {"status": "draft_saved", "to": to, "subject": subject, "logo": logo_status,
                "draft_id": str(item.EntryID),
                "hint": "已保存到草稿文件夹; 确认后可发送(或调用 send_draft), 也可将 send 设为 True 直接发送"}
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def reply_email(item_id: str, body: str, reply_all: bool = False, send: bool = False,
                signature: str = "", attachments: list = None,
                use_html: bool = False, cc: str = "") -> dict:
    """基于原邮件生成「真回复」(保留会话线程、原文引用与签名的 HTML 渲染)。

    Args:
        item_id: 原邮件 EntryID(来自 search_emails 结果)
        body: 回复正文(纯文本自动转 HTML; use_html=True 时按 HTML 原样)
        reply_all: True=回复全部(ReplyAll); False=仅回复发件人(默认)
        send: True=直接发送; False=仅保存草稿(默认, 避免误发)
        signature: 签名名称, 读取 Outlook 预设签名(如 "对外"/"对内"; 优先 HTML 版,
            含 logo 图片会自动内嵌为"隐藏"图片, 无 HTML 版时回退 .txt);
            默认读取环境变量 OUTLOOK_MCP_DEFAULT_SIGNATURE; 均为空则不追加签名
        attachments: 本地附件文件路径列表(可选, 可见附件; 路径不存在会直接报错)
        use_html: True=正文按 HTML 渲染(可放表格/B 标签); False=纯文本
        cc: 追加抄送(英文分号分隔; 保留 ReplyAll 已有的抄送人)
    """
    import html as _html
    try:
        app, ns, owned = _get_outlook()
        try:
            src = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 EntryID(邮件可能已被移动或删除)")
        reply = src.ReplyAll() if reply_all else src.Reply()

        if use_html:
            block = body or ""
        else:
            text = body.rstrip()
            body_html = _html.escape(text).replace("\r\n", "<br>").replace("\n", "<br>")
            block = ('<div style="font-family:等线;font-size:10.5pt;line-height:1.4;">'
                     + body_html + '</div>')

        sig_logos = []
        sig_name = (signature or os.environ.get("OUTLOOK_MCP_DEFAULT_SIGNATURE", "")).strip()
        if sig_name:
            sig_html, sig_logos = _build_signature(sig_name)
            if sig_html:
                block = block + "<br>" + sig_html

        if cc:
            existing = _unwrap(reply.CC) or ""
            reply.CC = f"{existing}; {cc}" if existing else cc

        h = reply.HTMLBody or ""
        idx = h.lower().find("<body")
        if idx >= 0:
            j = h.find(">", idx)
            h = h[:j + 1] + block + "<br><br>" + h[j + 1:]
        else:
            h = block + "<br><br>" + h
        reply.HTMLBody = h

        # 签名 logo: 经 Word 编辑器通道插入为隐藏内嵌图
        logo_status = "none"
        if sig_logos:
            logo_status = _embed_signature_logos(reply, sig_logos)

        # 普通附件(可见)
        for path in _validate_attachments(attachments):
            reply.Attachments.Add(path)

        to = ""
        try:
            to = "; ".join(_unwrap(r.Address) for r in reply.Recipients)
        except Exception:
            pass
        if send:
            reply.Save()
            reply.Send()
            return {"status": "sent", "to": to, "logo": logo_status}
        reply.Save()
        return {"status": "draft_saved", "to": to, "draft_id": str(reply.EntryID),
                "logo": logo_status,
                "hint": "真回复草稿已保存(HTML 渲染: 彩色签名+原文引用); 确认后可发送, 或调用 send_draft"}
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def send_draft(item_id: str) -> dict:
    """发送草稿箱中已保存的邮件(如用户确认后的回复草稿)。

    Args:
        item_id: 草稿的 EntryID(来自 reply_email/send_email 返回的 draft_id)
    """
    try:
        app, ns, owned = _get_outlook()
        try:
            item = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 EntryID(草稿可能已被删除或已发送)")
        subject = _unwrap(item.Subject)
        drafts = ns.GetDefaultFolder(16)
        try:
            parent_id = str(item.Parent.EntryID)
        except Exception:
            parent_id = ""
        if parent_id != str(drafts.EntryID):
            raise ValueError("只能发送草稿箱中的邮件; 该邮件不在草稿箱")
        item.Send()
        return {"status": "sent", "subject": subject}
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def list_drafts(subject_contains: str = "", limit: int = 20) -> list:
    """列出草稿箱(可按主题模糊过滤), 用于发信前查重/清理旧草稿。

    Args:
        subject_contains: 主题包含的文字(不区分大小写; 空 = 全部)
        limit: 最多返回条数
    """
    try:
        app, ns, owned = _get_outlook()
        drafts = ns.GetDefaultFolder(16)
        items = drafts.Items
        try:
            items.Sort("[LastModificationTime]", True)
        except Exception:
            pass
        out = []
        for item in items:
            try:
                if item.Class != 43:
                    continue
                subj = _unwrap(item.Subject)
                if subject_contains and subject_contains.lower() not in subj.lower():
                    continue
                d = _item_to_dict(item, include_body=False)
                d["folder"] = "草稿"
                try:
                    d["to"] = "; ".join(_unwrap(r.Address) for r in item.Recipients)
                except Exception:
                    d["to"] = ""
                out.append(d)
                if len(out) >= limit:
                    break
            except Exception:
                continue
        return out
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def delete_draft(item_id: str) -> dict:
    """删除草稿箱中的草稿(只允许草稿, 防止误删正式邮件)。

    Args:
        item_id: 草稿 EntryID(来自 list_drafts / send_email / reply_email)
    """
    try:
        app, ns, owned = _get_outlook()
        try:
            item = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 EntryID(草稿可能已被删除)")
        subject = _unwrap(item.Subject)
        drafts = ns.GetDefaultFolder(16)
        try:
            parent_id = str(item.Parent.EntryID)
        except Exception:
            parent_id = ""
        if parent_id != str(drafts.EntryID):
            raise ValueError("只能删除草稿箱中的邮件; 该邮件不在草稿箱")
        item.Delete()
        return {"status": "deleted", "subject": subject}
    finally:
        _safe_quit(app, owned)


@mcp.tool()
def forward_email(item_id: str, to: str, cc: str = "", body: str = "",
                  signature: str = "", attachments: list = None,
                  send: bool = False, use_html: bool = False) -> dict:
    """转发邮件(保留原文), 默认存草稿。

    Args:
        item_id: 原邮件 EntryID
        to: 收件人(英文分号分隔)
        cc: 抄送(英文分号分隔)
        body: 转发说明(可空; 纯文本自动转 HTML, use_html=True 时按 HTML 原样)
        signature: 签名名称(对外/对内)
        attachments: 追加附件路径列表(路径不存在会报错)
        send: True=直接发送; False=存草稿(默认)
        use_html: 正文是否按 HTML 渲染
    """
    import html as _html
    try:
        app, ns, owned = _get_outlook()
        try:
            src = ns.GetItemFromID(item_id)
        except Exception:
            raise ValueError("无效的 EntryID(邮件可能已被移动或删除)")
        fwd = src.Forward()
        fwd.To = to
        fwd.CC = cc or ""

        if body:
            if use_html:
                block = body
            else:
                text = body.rstrip()
                body_html = _html.escape(text).replace("\r\n", "<br>").replace("\n", "<br>")
                block = ('<div style="font-family:等线;font-size:10.5pt;line-height:1.4;">'
                         + body_html + '</div>')
        else:
            block = ""

        sig_logos = []
        sig_name = (signature or os.environ.get("OUTLOOK_MCP_DEFAULT_SIGNATURE", "")).strip()
        if sig_name:
            sig_html, sig_logos = _build_signature(sig_name)
            if sig_html:
                block = (block + "<br>" + sig_html) if block else sig_html

        if block:
            h = fwd.HTMLBody or ""
            idx = h.lower().find("<body")
            if idx >= 0:
                j = h.find(">", idx)
                h = h[:j + 1] + block + "<br><br>" + h[j + 1:]
            else:
                h = block + "<br><br>" + h
            fwd.HTMLBody = h

        logo_status = "none"
        if sig_logos:
            logo_status = _embed_signature_logos(fwd, sig_logos)

        for path in _validate_attachments(attachments):
            fwd.Attachments.Add(path)

        if send:
            fwd.Save()
            fwd.Send()
            return {"status": "sent", "to": to, "logo": logo_status}
        fwd.Save()
        return {"status": "draft_saved", "to": to, "draft_id": str(fwd.EntryID),
                "logo": logo_status, "hint": "转发草稿已保存; 确认后调用 send_draft"}
    finally:
        _safe_quit(app, owned)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
