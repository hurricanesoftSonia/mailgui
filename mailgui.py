"""
MailGUI - Hurricane Software Mail Client
A Windows-friendly desktop email client with SMTP send and IMAP receive.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import imaplib
import smtplib
import email
import email.utils
import ssl
import os
import json
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
from email.utils import parseaddr

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def _decode_header(s):
    if not s:
        return ""
    parts = decode_header(s)
    result = []
    for data, charset in parts:
        if isinstance(data, bytes):
            result.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(data)
    return "".join(result)


def _get_body(msg):
    text_body = ""
    html_body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if "attachment" in disp:
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            decoded = payload.decode(charset, errors="replace")
            if ct == "text/plain" and not text_body:
                text_body = decoded
            elif ct == "text/html" and not html_body:
                html_body = decoded
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            decoded = payload.decode(charset, errors="replace")
            if msg.get_content_type() == "text/html":
                html_body = decoded
            else:
                text_body = decoded
    return text_body, html_body


def _get_attachments(msg):
    attachments = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        if "attachment" not in disp and part.get_content_maintype() != "application":
            continue
        filename = part.get_filename()
        if filename:
            filename = _decode_header(filename)
            data = part.get_payload(decode=True)
            attachments.append({
                "filename": filename,
                "content_type": part.get_content_type(),
                "size": len(data) if data else 0,
                "data": data,
            })
    return attachments


class MailConfig:
    def __init__(self):
        self.data = {}
        self.load()

    def load(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                self.data = json.load(f)

    def save(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def get(self, key, default=""):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("帳號設定 Account Settings")
        self.config = config
        self.geometry("480x520")
        self.resizable(False, False)
        self.grab_set()

        pad = {"padx": 10, "pady": 4, "sticky": "ew"}

        # Account
        ttk.Label(self, text="帳號設定", font=("", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=(10, 5))

        ttk.Label(self, text="Email:").grid(row=1, column=0, padx=10, sticky="w")
        self.email_var = tk.StringVar(value=config.get("email"))
        ttk.Entry(self, textvariable=self.email_var, width=40).grid(row=1, column=1, **pad)

        ttk.Label(self, text="顯示名稱:").grid(row=2, column=0, padx=10, sticky="w")
        self.name_var = tk.StringVar(value=config.get("name"))
        ttk.Entry(self, textvariable=self.name_var, width=40).grid(row=2, column=1, **pad)

        ttk.Label(self, text="密碼:").grid(row=3, column=0, padx=10, sticky="w")
        self.pass_var = tk.StringVar(value=config.get("password"))
        ttk.Entry(self, textvariable=self.pass_var, show="*", width=40).grid(row=3, column=1, **pad)

        # SMTP
        ttk.Label(self, text="SMTP 寄信", font=("", 11, "bold")).grid(row=4, column=0, columnspan=2, pady=(15, 5))

        smtp = config.get("smtp", {})
        ttk.Label(self, text="主機:").grid(row=5, column=0, padx=10, sticky="w")
        self.smtp_host = tk.StringVar(value=smtp.get("host", "mx.hurricanesoft.com.tw"))
        ttk.Entry(self, textvariable=self.smtp_host, width=40).grid(row=5, column=1, **pad)

        ttk.Label(self, text="Port:").grid(row=6, column=0, padx=10, sticky="w")
        self.smtp_port = tk.StringVar(value=str(smtp.get("port", 25)))
        ttk.Entry(self, textvariable=self.smtp_port, width=40).grid(row=6, column=1, **pad)

        self.smtp_starttls = tk.BooleanVar(value=smtp.get("starttls", True))
        ttk.Checkbutton(self, text="STARTTLS", variable=self.smtp_starttls).grid(row=7, column=1, sticky="w", padx=10)

        # IMAP
        ttk.Label(self, text="IMAP 收信", font=("", 11, "bold")).grid(row=8, column=0, columnspan=2, pady=(15, 5))

        imap = config.get("imap", {})
        ttk.Label(self, text="主機:").grid(row=9, column=0, padx=10, sticky="w")
        self.imap_host = tk.StringVar(value=imap.get("host", "webmail.hurricanesoft.com.tw"))
        ttk.Entry(self, textvariable=self.imap_host, width=40).grid(row=9, column=1, **pad)

        ttk.Label(self, text="Port:").grid(row=10, column=0, padx=10, sticky="w")
        self.imap_port = tk.StringVar(value=str(imap.get("port", 993)))
        ttk.Entry(self, textvariable=self.imap_port, width=40).grid(row=10, column=1, **pad)

        self.imap_ssl = tk.BooleanVar(value=imap.get("ssl", True))
        ttk.Checkbutton(self, text="SSL", variable=self.imap_ssl).grid(row=11, column=1, sticky="w", padx=10)

        # Signature
        ttk.Label(self, text="簽名檔:").grid(row=12, column=0, padx=10, sticky="nw", pady=(10, 0))
        self.sig_text = tk.Text(self, width=40, height=4)
        self.sig_text.grid(row=12, column=1, padx=10, pady=(10, 5))
        self.sig_text.insert("1.0", config.get("signature", ""))

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=13, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="儲存", command=self.save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="left", padx=5)

        self.columnconfigure(1, weight=1)

    def save(self):
        self.config.set("email", self.email_var.get().strip())
        self.config.set("name", self.name_var.get().strip())
        self.config.set("password", self.pass_var.get().strip())
        self.config.set("signature", self.sig_text.get("1.0", "end-1c").strip())
        self.config.set("smtp", {
            "host": self.smtp_host.get().strip(),
            "port": int(self.smtp_port.get().strip()),
            "starttls": self.smtp_starttls.get(),
            "verify_ssl": False,
        })
        self.config.set("imap", {
            "host": self.imap_host.get().strip(),
            "port": int(self.imap_port.get().strip()),
            "ssl": self.imap_ssl.get(),
        })
        self.config.save()
        messagebox.showinfo("設定", "設定已儲存！")
        self.destroy()


class ComposeDialog(tk.Toplevel):
    def __init__(self, parent, config, reply_to=None, reply_all=False):
        super().__init__(parent)
        self.title("撰寫新郵件 Compose")
        self.config = config
        self.geometry("650x550")
        self.attachments = []
        self.grab_set()

        pad = {"padx": 8, "pady": 3, "sticky": "ew"}

        ttk.Label(self, text="收件人 To:").grid(row=0, column=0, padx=8, sticky="w")
        self.to_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.to_var, width=60).grid(row=0, column=1, **pad)

        ttk.Label(self, text="CC:").grid(row=1, column=0, padx=8, sticky="w")
        self.cc_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.cc_var, width=60).grid(row=1, column=1, **pad)

        ttk.Label(self, text="主旨 Subject:").grid(row=2, column=0, padx=8, sticky="w")
        self.subj_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.subj_var, width=60).grid(row=2, column=1, **pad)

        # Attachment bar
        att_frame = ttk.Frame(self)
        att_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        ttk.Button(att_frame, text="📎 附件", command=self.add_attachment).pack(side="left")
        self.att_label = ttk.Label(att_frame, text="無附件")
        self.att_label.pack(side="left", padx=10)

        # Body
        self.body_text = tk.Text(self, wrap="word")
        self.body_text.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=8, pady=5)

        # Pre-fill signature
        sig = config.get("signature", "")
        if sig:
            self.body_text.insert("end", f"\n\n{sig}")
            self.body_text.mark_set("insert", "1.0")

        # Pre-fill reply
        if reply_to:
            from_addr = _decode_header(reply_to.get("From", ""))
            subject = _decode_header(reply_to.get("Subject", ""))
            self.to_var.set(parseaddr(from_addr)[1])
            if not subject.lower().startswith("re:"):
                subject = f"Re: {subject}"
            self.subj_var.set(subject)
            if reply_all:
                cc_addrs = _decode_header(reply_to.get("Cc", ""))
                self.cc_var.set(cc_addrs)

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=8)
        ttk.Button(btn_frame, text="📤 寄出", command=self.send).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="left", padx=5)

        self.rowconfigure(4, weight=1)
        self.columnconfigure(1, weight=1)

    def add_attachment(self):
        paths = filedialog.askopenfilenames(title="選擇附件")
        for p in paths:
            self.attachments.append(p)
        if self.attachments:
            names = [os.path.basename(p) for p in self.attachments]
            self.att_label.config(text=", ".join(names))

    def send(self):
        to = self.to_var.get().strip()
        subject = self.subj_var.get().strip()
        body = self.body_text.get("1.0", "end-1c")
        cc = self.cc_var.get().strip() or None

        if not to:
            messagebox.showwarning("錯誤", "請填寫收件人！")
            return
        if not subject:
            messagebox.showwarning("錯誤", "請填寫主旨！")
            return

        self.title("寄送中...")
        threading.Thread(target=self._send_thread, args=(to, subject, body, cc), daemon=True).start()

    def _send_thread(self, to, subject, body, cc):
        try:
            cfg = self.config.data
            display_name = cfg.get("name", "") or cfg.get("email", "").split("@")[0]

            if self.attachments:
                msg = MIMEMultipart()
                msg.attach(MIMEText(body, "plain", "utf-8"))
                for path in self.attachments:
                    with open(path, "rb") as f:
                        data = f.read()
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(data)
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
                    msg.attach(part)
            else:
                msg = MIMEText(body, "plain", "utf-8")

            msg["From"] = f"{display_name} <{cfg['email']}>"
            msg["To"] = to
            msg["Subject"] = subject
            msg["Date"] = email.utils.formatdate(localtime=True)
            if cc:
                msg["Cc"] = cc

            smtp_cfg = cfg["smtp"]
            use_starttls = smtp_cfg.get("starttls", False)

            if use_starttls:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                s = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"])
                s.ehlo()
                s.starttls(context=ctx)
                s.ehlo()
            else:
                s = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"])

            with s:
                s.login(cfg["email"], str(cfg["password"]))
                recipients = [to]
                if cc:
                    recipients.extend([a.strip() for a in cc.split(",")])
                s.send_message(msg, to_addrs=recipients)

            self.after(0, lambda: self._send_done(True))
        except Exception as e:
            self.after(0, lambda: self._send_done(False, str(e)))

    def _send_done(self, success, error=""):
        if success:
            messagebox.showinfo("成功", "郵件已寄出！✉️")
            self.destroy()
        else:
            messagebox.showerror("寄信失敗", f"錯誤：{error}")
            self.title("撰寫新郵件 Compose")


class MailGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📧 MailGUI - Hurricane Software Mail Client")
        self.root.geometry("1000x650")
        self.config = MailConfig()
        self.messages = []  # list of (uid, flags, parsed_msg, raw)
        self.current_folder = "INBOX"

        self._build_ui()

    def _build_ui(self):
        # Menu
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="⚙ 帳號設定", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self.root.quit)
        menubar.add_cascade(label="檔案", menu=file_menu)

        mail_menu = tk.Menu(menubar, tearoff=0)
        mail_menu.add_command(label="📥 收信", command=self.fetch_mail)
        mail_menu.add_command(label="📝 撰寫新郵件", command=self.compose)
        menubar.add_cascade(label="郵件", menu=mail_menu)
        self.root.config(menu=menubar)

        # Toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill="x", padx=5, pady=3)
        ttk.Button(toolbar, text="📥 收信", command=self.fetch_mail).pack(side="left", padx=2)
        ttk.Button(toolbar, text="📝 新郵件", command=self.compose).pack(side="left", padx=2)
        ttk.Button(toolbar, text="↩ 回覆", command=self.reply).pack(side="left", padx=2)
        ttk.Button(toolbar, text="↩↩ 全部回覆", command=self.reply_all).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🗑 刪除", command=self.delete_mail).pack(side="left", padx=2)
        ttk.Button(toolbar, text="⚙ 設定", command=self.open_settings).pack(side="right", padx=2)

        # Search bar
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(search_frame, text="🔍 搜尋:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        ttk.Entry(search_frame, textvariable=self.search_var, width=40).pack(side="left", padx=5)

        # Folder selector
        ttk.Label(search_frame, text="資料夾:").pack(side="left", padx=(20, 5))
        self.folder_var = tk.StringVar(value="INBOX")
        folder_combo = ttk.Combobox(search_frame, textvariable=self.folder_var,
                                     values=["INBOX", "Sent", "Drafts", "Trash"], width=12, state="readonly")
        folder_combo.pack(side="left")
        folder_combo.bind("<<ComboboxSelected>>", lambda e: self.fetch_mail())

        # Paned window
        paned = ttk.PanedWindow(self.root, orient="vertical")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # Mail list
        list_frame = ttk.Frame(paned)
        paned.add(list_frame, weight=1)

        cols = ("status", "from", "subject", "date")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("status", text="")
        self.tree.heading("from", text="寄件人")
        self.tree.heading("subject", text="主旨")
        self.tree.heading("date", text="日期")
        self.tree.column("status", width=30, stretch=False)
        self.tree.column("from", width=200)
        self.tree.column("subject", width=400)
        self.tree.column("date", width=160)
        self.tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Mail detail
        detail_frame = ttk.Frame(paned)
        paned.add(detail_frame, weight=2)

        # Header info
        hdr_frame = ttk.Frame(detail_frame)
        hdr_frame.pack(fill="x", padx=5, pady=3)
        self.detail_from = ttk.Label(hdr_frame, text="", font=("", 10, "bold"))
        self.detail_from.pack(anchor="w")
        self.detail_subject = ttk.Label(hdr_frame, text="", font=("", 11, "bold"))
        self.detail_subject.pack(anchor="w")
        self.detail_date = ttk.Label(hdr_frame, text="", foreground="gray")
        self.detail_date.pack(anchor="w")
        self.detail_attachments = ttk.Label(hdr_frame, text="", foreground="blue")
        self.detail_attachments.pack(anchor="w")

        # Body
        self.body_text = tk.Text(detail_frame, wrap="word", state="disabled")
        self.body_text.pack(fill="both", expand=True, padx=5, pady=3)

        # Status bar
        self.status_var = tk.StringVar(value="就緒 Ready")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken").pack(fill="x", side="bottom")

    def open_settings(self):
        SettingsDialog(self.root, self.config)

    def compose(self):
        if not self.config.get("email"):
            messagebox.showwarning("提醒", "請先設定帳號！")
            self.open_settings()
            return
        ComposeDialog(self.root, self.config)

    def reply(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        _, _, msg, _ = self.messages[idx]
        ComposeDialog(self.root, self.config, reply_to=msg)

    def reply_all(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        _, _, msg, _ = self.messages[idx]
        ComposeDialog(self.root, self.config, reply_to=msg, reply_all=True)

    def delete_mail(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("確認", "確定要刪除這封郵件嗎？"):
            return
        idx = self.tree.index(sel[0])
        uid = self.messages[idx][0]
        threading.Thread(target=self._delete_thread, args=(uid,), daemon=True).start()

    def _delete_thread(self, uid):
        try:
            cfg = self.config.data
            imap_cfg = cfg["imap"]
            if imap_cfg.get("ssl", True):
                M = imaplib.IMAP4_SSL(imap_cfg["host"], imap_cfg["port"])
            else:
                M = imaplib.IMAP4(imap_cfg["host"], imap_cfg["port"])
            M.login(cfg["email"], str(cfg["password"]))
            M.select(self.current_folder)
            M.uid("STORE", uid, "+FLAGS", "(\\Deleted)")
            M.expunge()
            M.logout()
            self.root.after(0, self.fetch_mail)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("刪除失敗", str(e)))

    def fetch_mail(self):
        if not self.config.get("email"):
            messagebox.showwarning("提醒", "請先設定帳號！")
            self.open_settings()
            return
        self.current_folder = self.folder_var.get()
        self.status_var.set("收信中...")
        threading.Thread(target=self._fetch_thread, daemon=True).start()

    def _fetch_thread(self):
        try:
            cfg = self.config.data
            imap_cfg = cfg["imap"]
            if imap_cfg.get("ssl", True):
                M = imaplib.IMAP4_SSL(imap_cfg["host"], imap_cfg["port"])
            else:
                M = imaplib.IMAP4(imap_cfg["host"], imap_cfg["port"])
            M.login(cfg["email"], str(cfg["password"]))

            folder = self.current_folder
            status, _ = M.select(folder, readonly=True)
            if status != "OK":
                M.select("INBOX", readonly=True)

            # Search
            _, data = M.uid("SEARCH", None, "ALL")
            uids = data[0].split()
            uids = uids[-200:]  # Last 200 messages

            messages = []
            if uids:
                # Fetch in batch
                uid_str = b",".join(uids)
                _, fetch_data = M.uid("FETCH", uid_str, "(FLAGS BODY.PEEK[])")
                i = 0
                while i < len(fetch_data):
                    item = fetch_data[i]
                    if isinstance(item, tuple) and len(item) == 2:
                        resp_line = item[0]
                        raw = item[1]
                        # Parse UID and flags from response
                        resp_str = resp_line.decode("utf-8", errors="replace") if isinstance(resp_line, bytes) else str(resp_line)
                        uid_val = ""
                        flags = ""
                        import re
                        uid_match = re.search(r"UID (\d+)", resp_str)
                        if uid_match:
                            uid_val = uid_match.group(1)
                        flag_match = re.search(r"FLAGS \(([^)]*)\)", resp_str)
                        if flag_match:
                            flags = flag_match.group(1)
                        msg = email.message_from_bytes(raw)
                        messages.append((uid_val, flags, msg, raw))
                    i += 1

            M.logout()
            messages.reverse()  # newest first
            self.messages = messages
            self.root.after(0, self._update_list)
        except Exception as e:
            self.root.after(0, lambda: self._fetch_error(str(e)))

    def _fetch_error(self, err):
        self.status_var.set("收信失敗")
        messagebox.showerror("收信錯誤", f"無法連線：{err}")

    def _update_list(self):
        self.tree.delete(*self.tree.get_children())
        search = self.search_var.get().lower()
        count = 0
        for uid, flags, msg, raw in self.messages:
            from_str = _decode_header(msg.get("From", ""))
            subject = _decode_header(msg.get("Subject", ""))
            date = msg.get("Date", "")
            # Search filter
            if search and search not in from_str.lower() and search not in subject.lower():
                continue
            status = "●" if "\\Seen" not in flags else "○"
            self.tree.insert("", "end", values=(status, from_str[:40], subject[:80], date[:25]))
            count += 1
        self.status_var.set(f"共 {count} 封郵件 ({self.current_folder})")

    def _on_search(self, *args):
        if self.messages:
            self._update_list()

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        # Map display index to messages index (accounting for search filter)
        display_idx = self.tree.index(sel[0])
        search = self.search_var.get().lower()
        actual_idx = -1
        count = 0
        for i, (uid, flags, msg, raw) in enumerate(self.messages):
            from_str = _decode_header(msg.get("From", "")).lower()
            subject = _decode_header(msg.get("Subject", "")).lower()
            if search and search not in from_str and search not in subject:
                continue
            if count == display_idx:
                actual_idx = i
                break
            count += 1

        if actual_idx < 0:
            return

        uid, flags, msg, raw = self.messages[actual_idx]
        from_str = _decode_header(msg.get("From", ""))
        subject = _decode_header(msg.get("Subject", ""))
        date = msg.get("Date", "")
        text_body, html_body = _get_body(msg)
        attachments = _get_attachments(msg)

        self.detail_from.config(text=f"From: {from_str}")
        self.detail_subject.config(text=f"Subject: {subject}")
        self.detail_date.config(text=f"Date: {date}")

        if attachments:
            att_names = [a["filename"] for a in attachments]
            self.detail_attachments.config(text=f"📎 附件: {', '.join(att_names)}")
            self.detail_attachments.bind("<Button-1>", lambda e: self._save_attachments(attachments))
        else:
            self.detail_attachments.config(text="")

        display = text_body or html_body or "(無內容)"
        self.body_text.config(state="normal")
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", display)
        self.body_text.config(state="disabled")

    def _save_attachments(self, attachments):
        folder = filedialog.askdirectory(title="選擇附件儲存目錄")
        if not folder:
            return
        for att in attachments:
            path = os.path.join(folder, att["filename"])
            with open(path, "wb") as f:
                f.write(att["data"])
        messagebox.showinfo("附件", f"已儲存 {len(attachments)} 個附件！")

    def run(self):
        self.root.mainloop()


def main():
    app = MailGUI()
    app.run()


if __name__ == "__main__":
    main()
