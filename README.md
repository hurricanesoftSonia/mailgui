# 📧 MailGUI - Hurricane Software Mail Client

颶風軟體桌面郵件收發軟體，支援 Windows / macOS。

## 功能

- 📥 **收信** — 支援 POP3（預設）和 IMAP 兩種協定
- 📤 **寄信 (SMTP)** — 撰寫新信、回覆、全部回覆、CC、附件
- ⚙️ **帳號設定** — GUI 設定 + config.json 檔案設定
- 🔍 **搜尋** — 即時搜尋寄件人/主旨
- 📁 **資料夾** — INBOX、Sent、Drafts、Trash
- ✉️ **已讀/未讀標記** — ● 未讀 / ○ 已讀
- 📎 **附件** — 寄送和下載附件
- 🔧 **自動設定** — 首次啟動自動建立預設設定檔

## 系統需求

- Python 3.9 以上
- tkinter（Python 內建 GUI 庫）
- 無需額外安裝套件

### Windows
- Python 安裝時勾選 tkinter（預設已勾）
- 或直接使用打包好的 `MailGUI.exe`

### macOS
- macOS 內建 Python 3 + tkinter
- 如果 tkinter 缺失：`brew install python-tk`

### Linux
- `sudo apt install python3-tk`（Debian/Ubuntu）
- `sudo dnf install python3-tkinter`（Fedora）

## 快速開始

### 方法 1：直接執行（所有平台）

```bash
git clone git@github.com:hurricanesoftSonia/mailgui.git
cd mailgui
python3 mailgui.py
```

### 方法 2：打包成執行檔

**Windows（.exe）：**
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name MailGUI mailgui.py
# 產出 dist/MailGUI.exe
```

**macOS（.app）：**
```bash
pip3 install pyinstaller
pyinstaller --onefile --windowed --name MailGUI mailgui.py
# 產出 dist/MailGUI.app
```

> ⚠️ macOS 首次開啟可能會被 Gatekeeper 擋住，右鍵 → 打開 即可。

## 設定

### GUI 設定
啟動後點「⚙ 設定」，填入帳號資訊即可。

### config.json 檔案設定
程式會在同目錄自動建立 `config.json`，也可以手動編輯：

```json
{
  "email": "yourname@hurricanesoft.com.tw",
  "name": "Your Name",
  "password": "your-password",
  "signature": "--\nYour Name\n颶風軟體有限公司",
  "recv_protocol": "pop3",
  "smtp": {
    "host": "mx.hurricanesoft.com.tw",
    "port": 25,
    "starttls": true,
    "verify_ssl": false
  },
  "imap": {
    "host": "webmail.hurricanesoft.com.tw",
    "port": 993,
    "ssl": true
  },
  "pop3": {
    "host": "webmail.hurricanesoft.com.tw",
    "port": 995,
    "ssl": true
  }
}
```

### 設定欄位說明

| 欄位 | 說明 | 預設值 |
|------|------|--------|
| email | 公司信箱 | （必填） |
| name | 顯示名稱 | （選填） |
| password | 信箱密碼 | （必填） |
| signature | 簽名檔 | （選填） |
| recv_protocol | 收信協定 | `pop3` |
| smtp.host | SMTP 主機 | mx.hurricanesoft.com.tw |
| smtp.port | SMTP 埠號 | 25 |
| smtp.starttls | 啟用 STARTTLS | true |
| imap.host | IMAP 主機 | webmail.hurricanesoft.com.tw |
| imap.port | IMAP 埠號 | 993 |
| pop3.host | POP3 主機 | webmail.hurricanesoft.com.tw |
| pop3.port | POP3 埠號 | 995 |

## 公司信箱

預設已配置颶風軟體公司郵件伺服器，只需填入 **email** 和 **密碼** 即可使用。

公司信箱使用 POP3 收信（預設），如需 IMAP 請在設定中切換協定。

## 部署方式

### Windows 使用者
1. 從 Samba `\\192.168.50.32\shared\工具\releases\` 取得 `MailGUI.exe` 和 `config.json`
2. 放在同一個目錄
3. 編輯 `config.json` 填入帳密（或啟動後在 GUI 設定）
4. 雙擊 `MailGUI.exe` 執行

### macOS 使用者
1. 確認有 Python 3.9+：`python3 --version`
2. Clone 或下載原始碼
3. 執行 `python3 mailgui.py`

## 目錄結構

```
mailgui/
├── mailgui.py          # 主程式
├── config.json         # 設定檔（自動建立）
├── README.md           # 本文件
└── .github/            # CI 設定（已停用）
```

## 常見問題

**Q: 啟動後看不到信？**
A: 確認 config.json 的 email 和 password 正確，收信協定選 POP3。

**Q: macOS 打不開 .app？**
A: 右鍵 → 打開，或到系統偏好設定 → 安全性與隱私 → 允許。

**Q: Windows exe 啟動找不到設定？**
A: 確認 config.json 和 MailGUI.exe 在同一個目錄。

**Q: 寄信失敗？**
A: 確認 SMTP 設定，公司信箱用 STARTTLS port 25。

## 版本紀錄

- **v1.3.0** — 找不到 config.json 自動建立預設設定
- **v1.2.0** — 修正 PyInstaller 打包 config 路徑問題
- **v1.1.0** — 新增 POP3 收信、CLI 設定
- **v1.0.0** — 初版（IMAP 收信、SMTP 寄信）

## 技術

- Python 3.9+
- tkinter（內建 GUI，零外部依賴）
- poplib / imaplib / smtplib（標準庫）

## License

MIT © Hurricane Software
