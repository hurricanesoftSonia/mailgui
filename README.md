# 📧 MailGUI - Hurricane Software Mail Client

Windows 桌面郵件收發軟體，支援 SMTP 寄信 / IMAP 收信。

## 功能

- 📥 **收信 (IMAP)** — 收件匣列表、讀信、附件下載
- 📤 **寄信 (SMTP)** — 撰寫新信、回覆、全部回覆、CC、附件
- ⚙️ **帳號設定** — SMTP/IMAP 主機、帳號密碼
- 🔍 **搜尋** — 即時搜尋寄件人/主旨
- 📁 **資料夾** — INBOX、Sent、Drafts、Trash
- ✉️ **已讀/未讀標記** — ● 未讀 / ○ 已讀
- 📎 **附件** — 寄送和下載附件

## 快速開始

### GUI 模式（圖形介面）

```bash
python mailgui.py
```

### CLI 模式（命令列）

**查看配置檔位置：**
```bash
python mailgui.py config
```

**發送郵件：**
```bash
# 簡單發送
python mailgui.py send --to user@example.com --subject "測試" --body "Hello"

# 從檔案讀取內容
python mailgui.py send --to user@example.com --subject "報告" --file report.txt

# 加上 CC 和附件
python mailgui.py send --to user@example.com --subject "文件" --body "請查收" \
  --cc other@example.com --attach file1.pdf file2.jpg
```

**接收郵件：**
```bash
# 接收最新 10 封郵件（預設）
python mailgui.py receive

# 接收最新 20 封郵件
python mailgui.py receive --count 20
```

### 打包成 .exe (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name MailGUI mailgui.py
```

產出的 `MailGUI.exe` 在 `dist/` 目錄下。

## 設定

首次啟動請點「⚙ 設定」填入帳號資訊：

| 欄位 | 預設值 |
|------|--------|
| SMTP 主機 | mx.hurricanesoft.com.tw |
| SMTP Port | 25 (STARTTLS) |
| IMAP 主機 | webmail.hurricanesoft.com.tw |
| IMAP Port | 993 (SSL) |

設定會儲存在 `config.json`。

## 公司信箱

預設已配置颶風軟體公司郵件伺服器，只需填入 email 和密碼即可使用。

## 技術

- Python 3.9+
- tkinter (內建 GUI)
- imaplib / smtplib (標準庫)
- 無需額外安裝套件

## License

MIT © Hurricane Software
