"""
MailGUI - Hurricane Software Mail Client
A Windows-friendly desktop email client with SMTP send and IMAP receive.
"""
__version__ = "1.1.0"

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import imaplib
import poplib
import smtplib
import email
import email.utils
import ssl
import os
import sys
import json
import threading
import argparse
import base64
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
from email.utils import parseaddr
from cryptography.fernet import Fernet

# PyInstaller frozen exe: use exe directory; otherwise use script directory
if getattr(sys, 'frozen', False):
    # When frozen (packaged), check multiple possible locations
    if sys.platform == 'darwin':  # macOS .app bundle
        _base_dir = os.path.dirname(os.path.dirname(os.path.dirname(sys.executable)))  # Go up to .app
        # Try Contents/Resources first (where pyinstaller puts data files)
        _resources_dir = os.path.join(os.path.dirname(sys.executable), '..', 'Resources')
        if os.path.exists(os.path.join(_resources_dir, 'config.json')):
            _base_dir = _resources_dir
    else:
        _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(_base_dir, "config.json")

# If config doesn't exist in expected location, try current directory and home directory
if not os.path.exists(CONFIG_FILE):
    for try_path in [
        os.path.join(os.getcwd(), "config.json"),
        os.path.join(os.path.expanduser("~"), ".mailgui", "config.json"),
    ]:
        if os.path.exists(try_path):
            CONFIG_FILE = try_path
            break
    else:
        # Create default in home directory
        CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".mailgui", "config.json")
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

# Encryption key file (stored next to config)
KEY_FILE = os.path.join(os.path.dirname(CONFIG_FILE), ".mailgui.key")


def _get_or_create_key():
    """Get or create encryption key"""
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, 'rb') as f:
            return f.read()
    else:
        # Generate a new key
        key = Fernet.generate_key()
        os.makedirs(os.path.dirname(KEY_FILE), exist_ok=True)
        with open(KEY_FILE, 'wb') as f:
            f.write(key)
        # Set file permissions to read/write for owner only
        os.chmod(KEY_FILE, 0o600)
        return key


def _encrypt_password(password):
    """Encrypt password using Fernet"""
    if not password:
        return ""
    key = _get_or_create_key()
    f = Fernet(key)
    return f.encrypt(password.encode()).decode()


def _decrypt_password(encrypted_password):
    """Decrypt password using Fernet"""
    if not encrypted_password:
        return ""
    try:
        key = _get_or_create_key()
        f = Fernet(key)
        return f.decrypt(encrypted_password.encode()).decode()
    except Exception:
        # If decryption fails, assume it's plain text (for backward compatibility)
        return encrypted_password


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
                # Decrypt password when loading
                if self.data.get("password"):
                    self.data["password"] = _decrypt_password(self.data["password"])
        else:
            # 自動建立預設 config
            self.data = {
                "email": "",
                "name": "",
                "password": "",
                "signature": "",
                "recv_protocol": "pop3",
                "smtp": {"host": "mx.hurricanesoft.com.tw", "port": 25, "starttls": True, "verify_ssl": False},
                "imap": {"host": "webmail.hurricanesoft.com.tw", "port": 993, "ssl": True},
                "pop3": {"host": "webmail.hurricanesoft.com.tw", "port": 995, "ssl": True}
            }
            self.save()

    def save(self):
        # Create a copy of data for saving
        save_data = self.data.copy()
        # Encrypt password before saving
        if save_data.get("password"):
            save_data["password"] = _encrypt_password(save_data["password"])

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(save_data, f, indent=2, ensure_ascii=False)

        # Set file permissions to read/write for owner only
        os.chmod(CONFIG_FILE, 0o600)

    def get(self, key, default=""):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("帳號設定 Account Settings")
        self.config = config
        self.geometry("480x620")
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

        # Receive protocol selector
        ttk.Label(self, text="收信設定", font=("", 11, "bold")).grid(row=8, column=0, columnspan=2, pady=(15, 5))

        ttk.Label(self, text="協定:").grid(row=9, column=0, padx=10, sticky="w")
        self.recv_protocol = tk.StringVar(value=config.get("recv_protocol", "imap"))
        proto_frame = ttk.Frame(self)
        proto_frame.grid(row=9, column=1, sticky="w", padx=10)
        ttk.Radiobutton(proto_frame, text="IMAP", variable=self.recv_protocol, value="imap").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(proto_frame, text="POP3", variable=self.recv_protocol, value="pop3").pack(side="left")

        imap = config.get("imap", {})
        ttk.Label(self, text="主機:").grid(row=10, column=0, padx=10, sticky="w")
        self.imap_host = tk.StringVar(value=imap.get("host", "webmail.hurricanesoft.com.tw"))
        ttk.Entry(self, textvariable=self.imap_host, width=40).grid(row=10, column=1, **pad)

        ttk.Label(self, text="Port:").grid(row=11, column=0, padx=10, sticky="w")
        self.imap_port = tk.StringVar(value=str(imap.get("port", 993)))
        ttk.Entry(self, textvariable=self.imap_port, width=40).grid(row=11, column=1, **pad)

        self.imap_ssl = tk.BooleanVar(value=imap.get("ssl", True))
        ttk.Checkbutton(self, text="SSL", variable=self.imap_ssl).grid(row=12, column=1, sticky="w", padx=10)

        # POP3 settings (shared host/port/ssl with IMAP fields when protocol is pop3)
        pop3 = config.get("pop3", {})
        ttk.Label(self, text="POP3 主機:").grid(row=13, column=0, padx=10, sticky="w")
        self.pop3_host = tk.StringVar(value=pop3.get("host", "webmail.hurricanesoft.com.tw"))
        ttk.Entry(self, textvariable=self.pop3_host, width=40).grid(row=13, column=1, **pad)

        ttk.Label(self, text="POP3 Port:").grid(row=14, column=0, padx=10, sticky="w")
        self.pop3_port = tk.StringVar(value=str(pop3.get("port", 995)))
        ttk.Entry(self, textvariable=self.pop3_port, width=40).grid(row=14, column=1, **pad)

        self.pop3_ssl = tk.BooleanVar(value=pop3.get("ssl", True))
        ttk.Checkbutton(self, text="POP3 SSL", variable=self.pop3_ssl).grid(row=15, column=1, sticky="w", padx=10)

        # Signature
        ttk.Label(self, text="簽名檔:").grid(row=16, column=0, padx=10, sticky="nw", pady=(10, 0))
        self.sig_text = tk.Text(self, width=40, height=4)
        self.sig_text.grid(row=16, column=1, padx=10, pady=(10, 5))
        self.sig_text.insert("1.0", config.get("signature", ""))

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=17, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="儲存", command=self.save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=self.destroy).pack(side="left", padx=5)

        self.columnconfigure(1, weight=1)

    def save(self):
        # Disable save button to prevent double-click
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Frame):
                for btn in widget.winfo_children():
                    if isinstance(btn, ttk.Button) and btn.cget("text") == "儲存":
                        btn.configure(state="disabled")

        # Save in background thread
        threading.Thread(target=self._save_thread, daemon=True).start()

    def _save_thread(self):
        try:
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
            self.config.set("recv_protocol", self.recv_protocol.get())
            self.config.set("imap", {
                "host": self.imap_host.get().strip(),
                "port": int(self.imap_port.get().strip()),
                "ssl": self.imap_ssl.get(),
            })
            self.config.set("pop3", {
                "host": self.pop3_host.get().strip(),
                "port": int(self.pop3_port.get().strip()),
                "ssl": self.pop3_ssl.get(),
            })
            self.config.save()
            self.after(0, self._save_complete)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("儲存失敗", str(e)))

    def _save_complete(self):
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
        self.root.title(f"📧 MailGUI v{__version__} - Hurricane Software Mail Client")
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
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill="x", padx=5, pady=3)
        ttk.Button(self.toolbar, text="📥 收信", command=self.fetch_mail).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="📝 新郵件", command=self.compose).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="↩ 回覆", command=self.reply).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="↩↩ 全部回覆", command=self.reply_all).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="🗑 刪除", command=self.delete_mail).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="⚙ 設定", command=self.open_settings).pack(side="right", padx=2)

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
        self.status_var = tk.StringVar(value=f"✓ 就緒 | MailGUI v{__version__}")
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
            protocol = cfg.get("recv_protocol", "imap")
            if protocol == "pop3":
                pop3_cfg = cfg.get("pop3", {})
                host = pop3_cfg.get("host", "webmail.hurricanesoft.com.tw")
                port = pop3_cfg.get("port", 995)
                if pop3_cfg.get("ssl", True):
                    M = poplib.POP3_SSL(host, port)
                else:
                    M = poplib.POP3(host, port)
                M.user(cfg["email"])
                M.pass_(str(cfg["password"]))
                M.dele(int(uid))
                M.quit()
            else:
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

        # Disable buttons during fetch
        self._set_buttons_state("disabled")
        self.current_folder = self.folder_var.get()
        self.status_var.set("📨 收信中...")
        self.root.config(cursor="watch")

        threading.Thread(target=self._fetch_thread, daemon=True).start()

    def _set_buttons_state(self, state):
        """Enable or disable toolbar buttons"""
        for widget in self.toolbar.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=state)

    def _fetch_thread(self):
        try:
            cfg = self.config.data
            protocol = cfg.get("recv_protocol", "imap")
            if protocol == "pop3":
                self._fetch_pop3(cfg)
            else:
                self._fetch_imap(cfg)
            # Re-enable buttons after successful fetch
            self.root.after(0, self._fetch_complete)
        except Exception as e:
            self.root.after(0, lambda: self._fetch_error(str(e)))

    def _fetch_complete(self):
        """Called when fetch completes successfully"""
        self.status_var.set(f"✓ 已載入 {len(self.messages)} 封郵件")
        self.root.config(cursor="")
        self._set_buttons_state("normal")

    def _fetch_pop3(self, cfg):
        pop3_cfg = cfg.get("pop3", {})
        host = pop3_cfg.get("host", "webmail.hurricanesoft.com.tw")
        port = pop3_cfg.get("port", 995)
        use_ssl = pop3_cfg.get("ssl", True)

        if use_ssl:
            M = poplib.POP3_SSL(host, port)
        else:
            M = poplib.POP3(host, port)

        M.user(cfg["email"])
        M.pass_(str(cfg["password"]))

        # Get message count
        count, _ = M.stat()
        # Fetch last 50 messages (configurable limit for performance)
        fetch_limit = 50
        start = max(1, count - fetch_limit + 1)

        messages = []
        for i in range(count, start - 1, -1):
            try:
                resp, lines, octets = M.retr(i)
                raw = b"\r\n".join(lines)
                msg = email.message_from_bytes(raw)
                # POP3 doesn't have UIDs or flags; use message number as id
                messages.append((str(i), "", msg, raw))
            except Exception:
                continue

        M.quit()
        self.messages = messages  # already newest-first due to reversed range
        self.root.after(0, self._update_list)

    def _fetch_imap(self, cfg):
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
        fetch_limit = 50  # Limit for performance
        uids = uids[-fetch_limit:]  # Last N messages

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

    def _fetch_error(self, err):
        self.status_var.set("✗ 收信失敗")
        self.root.config(cursor="")
        self._set_buttons_state("normal")
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


def cli_send(args, config):
    """Send email via CLI"""
    cfg = config.data

    if not cfg.get("email") or not cfg.get("password"):
        print("Error: Email and password not configured. Run GUI mode to configure.")
        sys.exit(1)

    # Create message
    msg = MIMEMultipart()
    msg['From'] = f"{cfg.get('name', '')} <{cfg['email']}>"
    msg['To'] = args.to
    msg['Subject'] = args.subject

    if args.cc:
        msg['Cc'] = args.cc

    # Add body
    body = args.body
    if args.file:
        with open(args.file, 'r', encoding='utf-8') as f:
            body = f.read()

    msg.attach(MIMEText(body, 'plain'))

    # Add attachments
    if args.attach:
        for filepath in args.attach:
            if os.path.exists(filepath):
                with open(filepath, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filepath)}')
                    msg.attach(part)

    # Send email
    try:
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
            recipients = [args.to]
            if args.cc:
                recipients.extend([a.strip() for a in args.cc.split(",")])
            s.send_message(msg, to_addrs=recipients)

        print(f"✓ Email sent successfully to {args.to}")
    except Exception as e:
        print(f"✗ Failed to send email: {e}")
        sys.exit(1)


def cli_receive(args, config):
    """Receive emails via CLI"""
    cfg = config.data

    if not cfg.get("email") or not cfg.get("password"):
        print("Error: Email and password not configured. Run GUI mode to configure.")
        sys.exit(1)

    protocol = cfg.get("recv_protocol", "pop3")

    try:
        if protocol == "imap":
            imap_cfg = cfg["imap"]
            M = imaplib.IMAP4_SSL(imap_cfg["host"], imap_cfg["port"])
            M.login(cfg["email"], str(cfg["password"]))
            M.select("INBOX")

            _, data = M.search(None, "ALL")
            mail_ids = data[0].split()

            count = min(args.count, len(mail_ids))
            print(f"\n📬 Fetching {count} emails from INBOX...\n")

            for i in range(1, count + 1):
                mail_id = mail_ids[-i]
                _, msg_data = M.fetch(mail_id, "(RFC822)")
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                from_addr = parseaddr(msg.get("From", ""))[1]
                subject = _decode_header(msg.get("Subject", "(無主旨)"))
                date = msg.get("Date", "")

                print(f"[{i}] From: {from_addr}")
                print(f"    Subject: {subject}")
                print(f"    Date: {date}")
                print()

            M.close()
            M.logout()
        else:  # POP3
            pop3_cfg = cfg["pop3"]
            M = poplib.POP3_SSL(pop3_cfg["host"], pop3_cfg["port"])
            M.user(cfg["email"])
            M.pass_(str(cfg["password"]))

            num_messages = len(M.list()[1])
            count = min(args.count, num_messages)

            print(f"\n📬 Fetching {count} emails...\n")

            for i in range(1, count + 1):
                _, lines, _ = M.retr(num_messages - i + 1)
                raw_email = b"\r\n".join(lines)
                msg = email.message_from_bytes(raw_email)

                from_addr = parseaddr(msg.get("From", ""))[1]
                subject = _decode_header(msg.get("Subject", "(無主旨)"))
                date = msg.get("Date", "")

                print(f"[{i}] From: {from_addr}")
                print(f"    Subject: {subject}")
                print(f"    Date: {date}")
                print()

            M.quit()

    except Exception as e:
        print(f"✗ Failed to receive emails: {e}")
        sys.exit(1)


def cli_setup(args, config):
    """Setup email account via CLI"""
    print("📧 MailGUI Account Setup\n")

    # Email
    if args.email:
        email_addr = args.email
    else:
        email_addr = input("Email address: ").strip()

    # Name
    if args.name:
        name = args.name
    else:
        name = input("Display name (optional): ").strip()

    # Password
    if args.password:
        password = args.password
        print("⚠️  Warning: Passing password via --password is insecure!")
    else:
        password = getpass.getpass("Password: ")
        password_confirm = getpass.getpass("Confirm password: ")
        if password != password_confirm:
            print("✗ Passwords don't match!")
            sys.exit(1)

    # Server settings
    use_defaults = True
    if not args.smtp_host:
        use_default = input("Use default Hurricane Software mail servers? [Y/n]: ").strip().lower()
        use_defaults = use_default in ['', 'y', 'yes']

    if use_defaults and not args.smtp_host:
        smtp_host = "hurricanesoft.com.tw"
        smtp_port = 465
        smtp_ssl = True
        imap_host = "hurricanesoft.com.tw"
        imap_port = 993
        pop3_host = "hurricanesoft.com.tw"
        pop3_port = 995
        recv_protocol = "pop3"
    else:
        smtp_host = args.smtp_host or input("SMTP server: ").strip()
        smtp_port = args.smtp_port or int(input("SMTP port [465]: ").strip() or "465")
        smtp_ssl = smtp_port == 465
        recv_protocol = args.protocol or input("Receive protocol (imap/pop3) [pop3]: ").strip().lower() or "pop3"

        if recv_protocol == "imap":
            imap_host = args.imap_host or input("IMAP server: ").strip()
            imap_port = args.imap_port or int(input("IMAP port [993]: ").strip() or "993")
            pop3_host = imap_host
            pop3_port = 995
        else:
            pop3_host = args.pop3_host or input("POP3 server: ").strip()
            pop3_port = args.pop3_port or int(input("POP3 port [995]: ").strip() or "995")
            imap_host = pop3_host
            imap_port = 993

    # Update config
    config.set("email", email_addr)
    config.set("name", name)
    config.set("password", password)
    config.set("recv_protocol", recv_protocol)
    config.set("smtp", {
        "host": smtp_host,
        "port": smtp_port,
        "starttls": not smtp_ssl,
        "verify_ssl": False
    })
    config.set("imap", {
        "host": imap_host,
        "port": imap_port,
        "ssl": True
    })
    config.set("pop3", {
        "host": pop3_host,
        "port": pop3_port,
        "ssl": True
    })

    config.save()

    print(f"\n✅ Account configured successfully!")
    print(f"📧 Email: {email_addr}")
    print(f"📁 Config file: {CONFIG_FILE}")
    print(f"🔐 Password encrypted and saved securely")


def main():
    parser = argparse.ArgumentParser(
        description='MailGUI - Hurricane Software Mail Client',
        epilog='Run without arguments to launch GUI mode'
    )

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Send command
    send_parser = subparsers.add_parser('send', help='Send an email')
    send_parser.add_argument('--to', required=True, help='Recipient email address')
    send_parser.add_argument('--subject', required=True, help='Email subject')
    send_parser.add_argument('--body', help='Email body text')
    send_parser.add_argument('--file', help='Read email body from file')
    send_parser.add_argument('--cc', help='CC recipients (comma-separated)')
    send_parser.add_argument('--attach', nargs='+', help='Attachment file paths')

    # Receive command
    recv_parser = subparsers.add_parser('receive', help='Receive emails')
    recv_parser.add_argument('--count', type=int, default=10, help='Number of emails to fetch (default: 10)')

    # Setup command
    setup_parser = subparsers.add_parser('setup', help='Configure email account (interactive)')
    setup_parser.add_argument('--email', help='Email address')
    setup_parser.add_argument('--name', help='Display name')
    setup_parser.add_argument('--password', help='Password (insecure - prompt recommended)')
    setup_parser.add_argument('--smtp-host', help='SMTP server')
    setup_parser.add_argument('--smtp-port', type=int, help='SMTP port')
    setup_parser.add_argument('--imap-host', help='IMAP server')
    setup_parser.add_argument('--imap-port', type=int, help='IMAP port')
    setup_parser.add_argument('--pop3-host', help='POP3 server')
    setup_parser.add_argument('--pop3-port', type=int, help='POP3 port')
    setup_parser.add_argument('--protocol', choices=['imap', 'pop3'], help='Receive protocol')

    # Config command
    config_parser = subparsers.add_parser('config', help='Show configuration file location')

    args = parser.parse_args()

    # Load config
    config = MailConfig()

    if args.command == 'send':
        cli_send(args, config)
    elif args.command == 'receive':
        cli_receive(args, config)
    elif args.command == 'setup':
        cli_setup(args, config)
    elif args.command == 'config':
        print(f"Configuration file: {CONFIG_FILE}")
        print(f"Exists: {os.path.exists(CONFIG_FILE)}")
        if os.path.exists(CONFIG_FILE):
            print(f"Encryption key: {KEY_FILE}")
            print(f"Password encrypted: Yes")
    else:
        # No command = GUI mode
        app = MailGUI()
        app.run()


if __name__ == "__main__":
    main()
