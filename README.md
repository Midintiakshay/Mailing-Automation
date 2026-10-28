# Bulk Mailer — Setup & User Guide

Automated email campaign mailer for Bigrock / Roundcube / MailHostBox webmail.

---

## 🚀 Quick Start (Web Dashboard — Recommended for Non-IT Users)

### Option 1: One-Click Desktop Launcher (Windows)
Double-click **`run_ui.bat`** in this folder.
- It will verify the required packages and automatically open your default browser at:
  👉 **http://127.0.0.1:8000**

### Option 2: Command Line Launch
```bash
pip install -r requirements.txt
python app.py
```

### In the Web Dashboard:
1. **Enter Credentials**:
   - Sender Email (e.g., `oralhealth@scienceresearchpub.org`)
   - Email Password (with toggleable show/hide eye icon). *Kept securely in session memory, never written to disk.*
2. **Review Campaign Files**:
   - **Recipients**: Defaults to `recipients.csv` (or upload new `.csv` / `.xlsx` with column mapping preview).
   - **Template**: Defaults to `template.docx` (with live rich HTML preview and `{name}` / `{paper_title}` placeholders).
   - **Subjects**: Defaults to `subject.docx` (alternates between subjects, or uses built-in templates).
3. **Control Campaign**:
   - Click **Start Campaign** to begin.
   - Use **Pause** and **Stop** buttons whenever needed.
4. **Live Monitoring**:
   - Real-time progress bar and metrics (Total, Sent, Failed, Remaining).
   - Live **Sent Emails Tracker** table with search and status filters.
   - Live **Automation Activity Console** streaming browser actions.
   - Click **Export CSV** to download the latest `send_log.csv`.

---

## 🖥️ Command Line (CLI) Fallback

If you prefer to run directly from the terminal without the web interface:
```bash
python bigrock.py
```

---

## 📋 Campaign Files Format

### 1. Recipients List (`recipients.csv` or `.xlsx`)
Columns supported (case-insensitive):
- `name` / `author` — recipient's name
- `email` / `e-mail` / `mail id` — recipient's email address
- `topic` / `subject` — optional topic for subject personalization
- `paper_title` / `title` — optional paper title substituted into `{paper_title}`

### 2. Email Template (`template.docx` or `.txt`)
- Supports rich text formatting (bold, italic, fonts, sizes, paragraph margins).
- Use `{name}` anywhere you want the author's name replaced.
- Use `{paper_title}` anywhere you want the paper title replaced.
- Also automatically personalizes `Dear Researcher` to `Dear [Name]`.

### 3. Subject Lines (`subject.docx` or `.txt`)
- Put each alternative subject line on a new paragraph/line.
- The mailer alternates or randomly selects from these subjects.

---

## 🔒 Safety & Best Practices
- **Password Privacy**: Passwords are only used to log into Roundcube locally via your browser session and are never saved to disk.
- **Rate Limiting**: Keep the built-in delay (default: 30 seconds) to ensure optimal delivery and prevent server rate-limiting.
- **Log Tracking**: All results are continuously written to `send_log.csv` with timestamps and error codes.
