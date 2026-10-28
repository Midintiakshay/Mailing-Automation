#!/usr/bin/env python3
"""
Browser Automation Mailer (Roundcube / MailHostBox webmail)
===========================================================
Opens a real Chrome window, logs into your webmail, and for every recipient
in your Excel (or CSV) sheet it:

    finds Compose  ->  waits for the compose screen (new window OR same tab)
    ->  fills To  ->  fills Subject (auto-alternating built-in subjects)
    ->  fills the message body  ->  finds and clicks Send

SETUP
-----
1) pip install selenium webdriver-manager openpyxl --break-system-packages
2) Put next to this script:
     recipients.xlsx   (or recipients.csv)  with columns: name, email, paper_title
     email_template.txt                     with {name} and {paper_title} placeholders
3) python3 bulk_mailer.py

Every step prints its name. On failure it saves error_<...>.png plus
error_<...>.html so the exact page state can be inspected.
"""

import csv
import getpass
import html
import os
import random
import re
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
WEBMAIL_URL = "https://us3.webmail.mailhostbox.com"
FROM_EMAIL = "oralhealth@scienceresearchpub.org"
PASSWORD = ""   # Leave empty to prompt in terminal, or enter password here for auto-login

RECIPIENTS_FILE = "recipients.csv"   # explicit CSV file to use
TEMPLATE_FILE = "template"              # auto-detect .txt, .docx, .doc, .word, etc.
SUBJECTS_FILE = "subject"              # auto-detect .txt, .docx, .doc, .word, etc.
LOG_FILE = "send_log.csv"
DELAY_BETWEEN_EMAILS_SECONDS = 30

# Compose opens as a separate screen in Roundcube. If the Compose button
# cannot be clicked for any reason, we fall back to this direct URL.
COMPOSE_URL = WEBMAIL_URL.rstrip("/") + "/?_task=mail&_action=compose"


# ---------------------------------------------------------------------------
# BUILT-IN SUBJECT LINE  
# ---------------------------------------------------------------------------
def get_subject(name, topic, idx):
    topic = (topic or "").strip()
    author = (name or "").strip()
    prefixes = (
        "Invitation to Submit",
        "Research Article Invitation",
        "Research Publication Opportunity",
        "Invitation for Manuscript Consideration",
        "Call for Research Submissions",
    )
    prefix = prefixes[idx % len(prefixes)]

    if topic and len(topic) <= 80:
        return f"{prefix}: {topic}"
    if author:
        return f"{prefix} for {author}"
    if topic:
        return f"{prefix} for Researcher"
    return prefix


# ---------------------------------------------------------------------------
# DATA HELPERS  (Excel + CSV)
# ---------------------------------------------------------------------------
def _extract_field(record, *keys):
    for k in keys:
        norm = k.strip().lower()
        if norm in record and str(record[norm]).strip():
            return str(record[norm]).strip()
    norm_rec = {
        re.sub(r'[\s_\-\.]+', '', k.lower()): str(v).strip()
        for k, v in record.items() if v and str(v).strip()
    }
    for k in keys:
        norm_k = re.sub(r'[\s_\-\.]+', '', k.lower())
        if norm_k in norm_rec and norm_rec[norm_k]:
            return norm_rec[norm_k]
    return ""


_NAME_CANDIDATES = [
    "author_name", "author name", "name", "author", "authors", "author_names",
    "full_name", "fullname", "first_name", "first author", "corresponding author",
    "recipient", "recipient_name", "contact_name", "contact", "scholar", "researcher"
]

_EMAIL_CANDIDATES = [
    "email", "e-mail", "email id", "email_id", "mail", "mail id", "mail_id",
    "author_email", "author email", "contact_email", "recipient_email"
]

_TITLE_CANDIDATES = [
    "paper_title", "paper title", "title", "article_title", "article title",
    "paper", "article", "manuscript_title", "manuscript title", "work_title"
]

_TOPIC_CANDIDATES = [
    "topic", "subject", "journal", "category", "field", "research_area", "discipline"
]


def _clean_row(name, email, paper_title, topic, row_num=None):
    email = (email or "").strip()
    if not email or "@" not in email:
        if email:  # has text but no @ sign
            print(f"  [WARNING] Row {row_num}: invalid email '{email}' — skipped")
        return None
    return {
        "name": (name or "").strip(),
        "email": email,
        "paper_title": (paper_title or "").strip(),
        "topic": (topic or "").strip(),
    }


def load_recipients():
    """Load recipients directly from RECIPIENTS_FILE (no auto-discovery)."""
    path = Path(RECIPIENTS_FILE)
    if not path.exists():
        raise FileNotFoundError(f"Recipients file not found: {RECIPIENTS_FILE}")

    rows = []
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook

        ws = load_workbook(path, read_only=True, data_only=True).active
        header = None
        for row_num, raw in enumerate(ws.iter_rows(values_only=True), start=1):
            values = ["" if c is None else str(c).strip() for c in raw]
            if header is None:
                header = [v.lstrip('\ufeff').strip().lower() for v in values]
                print(f"  Excel header columns detected: {header}")
                continue
            record = {k: v for k, v in zip(header, values) if k}
            email = _extract_field(record, *_EMAIL_CANDIDATES)
            name = _extract_field(record, *_NAME_CANDIDATES)
            paper_title = _extract_field(record, *_TITLE_CANDIDATES)
            topic = _extract_field(record, *_TOPIC_CANDIDATES)

            row = _clean_row(name, email, paper_title, topic, row_num=row_num)
            if row:
                rows.append(row)
    else:
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f)
            header_fields = [h.strip().lower() for h in (reader.fieldnames or [])]
            print(f"  CSV header columns detected: {header_fields}")
            for row_num, raw_record in enumerate(reader, start=2):  # row 2 = first data row
                record = {
                    (k or "").lstrip('\ufeff').strip().lower(): (v or "").strip()
                    for k, v in raw_record.items() if k
                }
                email = _extract_field(record, *_EMAIL_CANDIDATES)
                name = _extract_field(record, *_NAME_CANDIDATES)
                paper_title = _extract_field(record, *_TITLE_CANDIDATES)
                topic = _extract_field(record, *_TOPIC_CANDIDATES)

                row = _clean_row(name, email, paper_title, topic, row_num=row_num)
                if row:
                    rows.append(row)

    print(f"Reading recipients from: {path.name}")
    if not rows:
        print("  [WARNING] No valid recipients found! Check your CSV columns and data.")
    else:
        print(f"  Found {len(rows)} valid recipient(s)")
        print(f"  First recipient: {rows[0]['name']} <{rows[0]['email']}>")
    return rows


def _read_docx_template(path):
    """Extract rich HTML and clean plain text from Word .docx file, preserving formatting.
    Preserves:
      - Paragraph boundaries
      - Bold, italic, underline
      - Font family and font size
      - Line breaks inside paragraphs (<br>)
      - Empty spacer paragraphs
    """
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "python-docx is required for .docx templates. Install it with: pip install python-docx"
        ) from exc

    import html
    import re

    doc = Document(str(path))
    paras = list(doc.paragraphs)
    # Strip trailing empty paragraphs
    while paras and not paras[-1].text.strip():
        paras.pop()

    html_paragraphs = []
    plain_paragraphs = []

    for p in paras:
        p_html_runs = []
        p_plain_runs = []
        for r in p.runs:
            raw_text = r.text
            escaped = html.escape(raw_text).replace('\n', '<br>')
            if r.bold:
                escaped = f'<b>{escaped}</b>'
            if r.italic:
                escaped = f'<i>{escaped}</i>'
            if r.underline:
                escaped = f'<u>{escaped}</u>'

            styles = []
            if r.font.name:
                styles.append(f"font-family: '{r.font.name}', Times, serif")
            if r.font.size:
                styles.append(f"font-size: {r.font.size.pt}pt")
            if r.font.color and r.font.color.rgb:
                styles.append(f"color: #{r.font.color.rgb}")

            if styles:
                escaped = f'<span style="{"; ".join(styles)}">{escaped}</span>'

            p_html_runs.append(escaped)
            p_plain_runs.append(raw_text)

        html_content = ''.join(p_html_runs)
        plain_content = ''.join(p_plain_runs)

        if not html_content.strip() or html_content.strip() == '<br>':
            html_paragraphs.append('<p style="margin: 0 0 12pt 0; line-height: 1.5;"><br></p>')
            plain_paragraphs.append('')
        else:
            align_style = ""
            if p.paragraph_format.alignment:
                align_name = str(p.paragraph_format.alignment).split('.')[-1].lower()
                if align_name in ("center", "right", "justify"):
                    align_style = f" text-align: {align_name};"

            html_paragraphs.append(
                f'<p style="margin: 0 0 12pt 0; line-height: 1.5;{align_style}">{html_content}</p>'
            )
            plain_paragraphs.append(plain_content)

    full_html = (
        '<div style="font-family: \'Times New Roman\', Times, serif; font-size: 12pt; color: #000000; line-height: 1.5;">\n'
        + '\n'.join(html_paragraphs)
        + '\n</div>'
    )
    full_plain = '\n\n'.join(plain_paragraphs)
    full_plain = re.sub(r'\n{3,}', '\n\n', full_plain).strip()

    return {"html": full_html, "plain": full_plain}


def _read_legacy_word_text(path):
    """Extract plain text from legacy Word files (.doc, .dot, .rtf, .word)."""
    if os.name != "nt":
        raise RuntimeError(
            f"Legacy Word template '{path.name}' requires Windows with Microsoft Word installed."
        )
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "For .doc templates, install pywin32: pip install pywin32"
        ) from exc

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        text = doc.Range().Text
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()
    return (text or "").strip()


def _text_to_template_dict(text):
    """Convert raw plain text into both HTML and clean plain text."""
    import html
    import re

    raw = text.strip().replace("\r\n", "\n")
    blocks = re.split(r'\n\s*\n', raw)
    html_paragraphs = []
    plain_paragraphs = []

    for b in blocks:
        lines = [l.strip() for l in b.split('\n') if l.strip()]
        if not lines:
            continue
        p_html = '<br>'.join(html.escape(l) for l in lines)
        html_paragraphs.append(f'<p style="margin: 0 0 12pt 0; line-height: 1.5;">{p_html}</p>')
        plain_paragraphs.append('\n'.join(lines))

    full_html = (
        '<div style="font-family: \'Times New Roman\', Times, serif; font-size: 12pt; color: #000000; line-height: 1.5;">\n'
        + '\n'.join(html_paragraphs)
        + '\n</div>'
    )
    full_plain = '\n\n'.join(plain_paragraphs)
    return {"html": full_html, "plain": full_plain}


def resolve_file_path(base_name):
    """Resolve the file even when the user gives a name without the exact extension."""
    requested = Path(base_name)
    if requested.exists():
        return requested

    if requested.suffix:
        stem = requested.with_suffix("")
        candidates = sorted(
            p for p in requested.parent.glob(f"{stem.name}.*")
            if p.suffix.lower() in {".txt", ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".word"}
        )
        if candidates:
            return candidates[0]

    # Fallback: if the configured name has no extension, search for any supported file near it.
    if not requested.suffix:
        candidates = sorted(
            p for p in requested.parent.glob(f"{requested.name}.*")
            if p.suffix.lower() in {".txt", ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".word"}
        )
        if candidates:
            return candidates[0]

    return requested


def load_template():
    """Load the email body template from TEMPLATE_FILE.
    Supports .docx (with full rich formatting), .txt, .doc, .rtf."""
    path = resolve_file_path(TEMPLATE_FILE)
    if not path.exists():
        raise FileNotFoundError(
            f"Template file not found: {TEMPLATE_FILE}\n"
            f"Please place a matching template next to this script, such as 'template.docx', 'template.doc', or 'template.txt'."
        )

    suffix = path.suffix.lower()
    if suffix in {".docx", ".docm", ".dotx", ".dotm"}:
        template_data = _read_docx_template(path)
    elif suffix in {".doc", ".dot", ".rtf", ".word"}:
        raw_text = _read_legacy_word_text(path)
        template_data = _text_to_template_dict(raw_text)
    else:
        raw_text = path.read_text(encoding="utf-8-sig").strip()
        template_data = _text_to_template_dict(raw_text)

    print(f"Loaded email template from: {path.name} (HTML: {len(template_data['html'])} chars, Plain: {len(template_data['plain'])} chars)")
    return template_data


def load_subjects():
    """Load subject lines from SUBJECTS_FILE.
    Supports .txt, .docx, .doc, .dot, .rtf, and .word templates."""
    path = resolve_file_path(SUBJECTS_FILE)
    if not path.exists():
        return []

    suffix = path.suffix.lower()
    if suffix in {".docx", ".docm", ".dotx", ".dotm"}:
        try:
            from docx import Document
            doc = Document(str(path))
            lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            text = "\n".join(lines)
        except Exception:
            text = ""
    elif suffix in {".doc", ".dot", ".rtf", ".word"}:
        text = _read_legacy_word_text(path)
    else:
        text = path.read_text(encoding="utf-8-sig").strip()

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if lines:
        print(f"Loaded {len(lines)} subjects from: {path.name}")
    return lines


def personalize_template(template_data, name, paper_title):
    """Replace placeholders in the template with actual recipient data.
    - {name} is replaced with author name (for 'Dear {name},' salutation)
    - {paper_title} is replaced with paper title
    - 'Dear Researcher' is also replaced with 'Dear {actual name}' as fallback
    """
    import html

    display_name = (name or "").strip()
    safe_name = html.escape(display_name) if display_name else "Researcher"
    fallback_name = display_name if display_name else "Researcher"
    paper_title_str = (paper_title or "").strip()
    safe_paper = html.escape(paper_title_str)

    if isinstance(template_data, dict):
        html_out = template_data.get("html", "")
        plain_out = template_data.get("plain", "")

        name_patterns = [
            r'\{\s*(name|author|author_name|authorname|recipient)\s*\}',
            r'\{name\}', r'\{NAME\}', r'\{Name\}',
            r'\{author_name\}', r'\{AUTHOR_NAME\}',
            r'\{author\}', r'\{AUTHOR\}',
        ]
        for pat in name_patterns:
            html_out = re.sub(pat, safe_name, html_out, flags=re.IGNORECASE)
            plain_out = re.sub(pat, fallback_name, plain_out, flags=re.IGNORECASE)

        title_patterns = [
            r'\{\s*(paper_title|papertitle|title|article_title)\s*\}',
            r'\{paper_title\}', r'\{PAPER_TITLE\}',
            r'\{title\}', r'\{TITLE\}',
        ]
        for pat in title_patterns:
            html_out = re.sub(pat, safe_paper, html_out, flags=re.IGNORECASE)
            plain_out = re.sub(pat, paper_title_str, plain_out, flags=re.IGNORECASE)

        if display_name:
            html_out = re.sub(r'Dear\s+Researcher\b', f'Dear {safe_name}', html_out, flags=re.IGNORECASE)
            plain_out = re.sub(r'Dear\s+Researcher\b', f'Dear {fallback_name}', plain_out, flags=re.IGNORECASE)

        return {"html": html_out, "plain": plain_out}
    else:
        # Fallback if string was passed
        text = str(template_data)
        text = text.replace("{name}", fallback_name)
        text = text.replace("{paper_title}", paper_title_str)
        text = text.replace("Dear Researcher", f"Dear {fallback_name}")
        return _text_to_template_dict(text)


def load_already_sent(log_path):
    sent = set()
    if Path(log_path).exists():
        try:
            with open(log_path, newline="", encoding="utf-8-sig", errors="replace") as f:
                for row in csv.DictReader(f):
                    if row and row.get("status") == "sent":
                        sent.add(row.get("email"))
        except Exception:
            pass
    return sent


def is_driver_alive(driver):
    if not driver:
        return False
    try:
        _ = driver.window_handles
        return True
    except Exception:
        return False


def is_fatal_browser_error(exc):
    msg = str(exc).lower()
    return any(p in msg for p in [
        "httpconnectionpool",
        "connection refused",
        "actively refused",
        "max retries exceeded",
        "newconnectionerror",
        "invalid session id",
        "no such window",
        "target window already closed",
        "session deleted",
        "chrome not reachable",
        "disconnected",
    ])


def log_result(log_path, email, name, subject, status, error=""):
    exists = Path(log_path).exists()
    try:
        with open(log_path, "a", newline="", encoding="utf-8", errors="replace") as f:
            w = csv.writer(f)
            if not exists:
                w.writerow(["timestamp", "email", "name", "subject", "status", "error"])
            w.writerow([time.strftime("%Y-%m-%d %H:%M:%S"), email, name, subject, status, error])
    except PermissionError:
        # Fallback if send_log.csv is locked by Microsoft Excel
        fb_path = Path("send_log_pending.csv")
        fb_exists = fb_path.exists()
        try:
            with open(fb_path, "a", newline="", encoding="utf-8", errors="replace") as f:
                w = csv.writer(f)
                if not fb_exists:
                    w.writerow(["timestamp", "email", "name", "subject", "status", "error"])
                w.writerow([time.strftime("%Y-%m-%d %H:%M:%S"), email, name, subject, status, error])
            print("  [WARNING] send_log.csv is open in Excel! Log saved to send_log_pending.csv.")
        except Exception:
            pass
    except Exception as e:
        print(f"  [WARNING] Could not write to log file: {e}")


# ---------------------------------------------------------------------------
# LOW-LEVEL BROWSER HELPERS
# ---------------------------------------------------------------------------
def start_browser():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_experimental_option(
        "prefs", {"profile.default_content_setting_values.popups": 1}
    )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    return driver


def visible_one(driver, css_list=(), xpath_list=(), timeout=10):
    """Return the first visible+enabled element matching any CSS or XPath."""
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        for by, selectors in ((By.CSS_SELECTOR, css_list), (By.XPATH, xpath_list)):
            for sel in selectors:
                try:
                    for el in driver.find_elements(by, sel):
                        if el.is_displayed() and el.is_enabled():
                            return el
                except Exception as exc:
                    last = exc
        time.sleep(0.4)
    raise Exception(f"No visible element for {list(css_list) + list(xpath_list)} ({last})")


def any_visible(driver, css_list=(), xpath_list=()):
    try:
        visible_one(driver, css_list, xpath_list, timeout=0.6)
        return True
    except Exception:
        return False


def safe_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


def dump(driver, tag):
    stamp = int(time.time())
    for name, action in (
        (f"error_{tag}_{stamp}.png", driver.save_screenshot),
        (f"error_{tag}_{stamp}.html", None),
    ):
        try:
            if action:
                action(name)
            else:
                Path(name).write_text(driver.page_source, encoding="utf-8")
            print(f"  saved {name}")
        except Exception:
            pass


def dismiss_any_dialogs(driver):
    try:
        buttons = driver.find_elements(
            By.CSS_SELECTOR,
            ".modal-dialog button, .modal-content button, #messagestack a, "
            "#confirm-dialog button, .ui-dialog-buttonset button, div.dialog button",
        )
        for btn in buttons:
            if btn.is_displayed() and any(
                k in (btn.text or "").lower() for k in ("discard", "yes", "ok", "close")
            ):
                safe_click(driver, btn)
                time.sleep(0.8)
                break
    except Exception:
        pass


# ---------------------------------------------------------------------------
# SELECTORS  (Roundcube Elastic + Larry + classic skins)
# ---------------------------------------------------------------------------
COMPOSE_CSS = [
    "a.compose", "a#rcmbtn107", "a#rcmbtn106", "#rcmbtn107", "#composebtn",
    "a[href*='_action=compose']", "a.button.compose", "a[onclick*='compose']",
    "[data-action='compose']",
]
COMPOSE_XPATH = [
    "//a[normalize-space(translate(@title,'COMPSE','compse'))='compose']",
    "//a[contains(translate(@title,'COMPSE','compse'),'compose')]",
    "//*[self::a or self::button][contains(translate(normalize-space(.),"
    "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'compose')]",
    "//*[self::a or self::button][contains(translate(normalize-space(.),"
    "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'new message')]",
]

TO_CSS = [
    "#compose_to input.recipient-input", "#compose_to input[role='combobox']",
    "#compose_to textarea", "#compose_to input", "textarea#_to", "input#_to",
    "textarea[name='_to']", "input[name='_to']", "input#rcmcomposeto",
    "input[role='combobox'][id*='to']",
]
TO_XPATH = ["//label[contains(.,'To')]/following::input[1]"]

SUBJECT_CSS = [
    "input#compose-subject", "input[name='_subject']", "#compose_subject input",
    "input#_subject", "input[placeholder*='ubject']",
]
SUBJECT_XPATH = ["//label[contains(.,'Subject')]/following::input[1]"]

BODY_IFRAME_CSS = ["iframe#composebody_ifr", "iframe.mce-edit-area", "iframe[id$='_ifr']"]
BODY_TEXTAREA_CSS = ["textarea#composebody", "textarea[name='_message']"]

# Priority dropdown in Roundcube compose form
PRIORITY_CSS = [
    "select[name='_priority']", "select#compose-priority", "select#_priority",
    "#compose_priority select",
]
PRIORITY_XPATH = [
    "//select[contains(@name,'priority')]",
    "//label[contains(.,'Priority')]/following::select[1]",
]

SEND_CSS = [
    "a.send", "button.send", "a#rcmbtn106", "button#rcmbtn106", "#rcmbtn106-clone",
    "button[data-action='send']", "a[onclick*='send']", "button.btn-primary.send",
]
SEND_XPATH = [
    "//a[contains(translate(@title,'SENDMSG','sendmsg'),'send')]",
    "//*[self::a or self::button][translate(normalize-space(.),"
    "'SEND','send')='send']",
    "//*[self::a or self::button][contains(translate(normalize-space(.),"
    "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'send')]",
]


# ---------------------------------------------------------------------------
# FLOW
# ---------------------------------------------------------------------------
def login(driver, wait, password):
    print("Step: login")
    driver.get(WEBMAIL_URL)
    user = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#rcmloginuser, input[name='_user']")
        )
    )
    user.clear()
    user.send_keys(FROM_EMAIL)
    pwd = driver.find_element(By.CSS_SELECTOR, "#rcmloginpwd, input[name='_pass']")

    if password and password.strip():
        safe_click(driver, pwd)
        pwd.clear()
        pwd.send_keys(password)
        # Verify password is in the input
        val = driver.execute_script("return arguments[0].value;", pwd)
        if not val:
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
            """, pwd, password)
        time.sleep(0.3)

        # Identify and click Login button
        submit_buttons = driver.find_elements(
            By.CSS_SELECTOR,
            "#rcmloginsubmit, button[type='submit'], input[type='submit'], button.btn-primary, #loginbtn, .formbuttons input"
        )
        clicked_btn = False
        for btn in submit_buttons:
            if btn.is_displayed() and btn.is_enabled():
                safe_click(driver, btn)
                clicked_btn = True
                break
        if not clicked_btn:
            pwd.send_keys(Keys.ENTER)
    else:
        print("\n[!] No password provided in terminal. Please type your password directly in Chrome and click Login.")

    # Wait for login completion or detect login errors (up to 60s)
    print("  Waiting for mailbox to load...")
    deadline = time.time() + 60
    while time.time() < deadline:
        if any_visible(driver, COMPOSE_CSS, COMPOSE_XPATH) or driver.find_elements(
            By.CSS_SELECTOR, "#mailboxlist, #messagelist, #layout, .folderlist"
        ):
            print("  Logged in successfully.\n")
            return
        err_elements = driver.find_elements(
            By.CSS_SELECTOR, "#messagestack div, #login-alert, .alert-danger, .boxerror"
        )
        for err in err_elements:
            if err.is_displayed() and err.text.strip():
                raise Exception(f"Login failed: {err.text.strip()}")
        time.sleep(1)

    raise Exception("Timed out waiting for login to complete.")


def open_compose(driver, wait, main_window):
    """Click Compose and return the window handle showing the compose form."""
    print("Step: find and click Compose")
    before = set(driver.window_handles)
    clicked = False
    try:
        safe_click(driver, visible_one(driver, COMPOSE_CSS, COMPOSE_XPATH, timeout=12))
        clicked = True
    except Exception as exc:
        print(f"  Compose button not clickable ({exc}) — using rcmail command")
        try:
            driver.execute_script("if (window.rcmail) rcmail.command('compose','',null,null);")
            clicked = True
        except Exception:
            pass

    print("Step: wait for compose screen")
    deadline = time.time() + 25
    while time.time() < deadline:
        new_handles = [h for h in driver.window_handles if h not in before]
        for handle in new_handles + [main_window]:
            try:
                driver.switch_to.window(handle)
            except Exception:
                continue
            if any_visible(driver, TO_CSS, TO_XPATH):
                print(f"  compose ready ({'new window' if handle != main_window else 'same window'})")
                return handle
        time.sleep(0.5)

    # Last resort: load the compose screen directly in a new tab.
    print("  compose form not detected — opening compose URL directly")
    driver.switch_to.new_window("tab")
    driver.get(COMPOSE_URL)
    WebDriverWait(driver, 20).until(lambda d: any_visible(d, TO_CSS, TO_XPATH))
    if not clicked:
        print("  (Compose button was never found; direct URL used instead)")
    return driver.current_window_handle


def fill_to(driver, to_email):
    print("Step: fill To")
    field = visible_one(driver, TO_CSS, TO_XPATH, timeout=15)
    safe_click(driver, field)
    try:
        field.clear()
    except Exception:
        pass
    field.send_keys(to_email)
    time.sleep(0.4)
    field.send_keys(Keys.ENTER)   # commits the recipient chip in Elastic
    for event in ("input", "change", "blur"):
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event(arguments[1],{bubbles:true}));",
            field, event,
        )


def fill_subject(driver, subject):
    print("Step: fill Subject")
    field = visible_one(driver, SUBJECT_CSS, SUBJECT_XPATH, timeout=12)
    safe_click(driver, field)
    try:
        field.clear()
    except Exception:
        pass
    field.send_keys(subject)
    if not (field.get_attribute("value") or "").strip():
        driver.execute_script(
            "arguments[0].value=arguments[1];"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));",
            field, subject,
        )


def fill_body(driver, body_data, compose_window):
    """Fill the email body with the personalized template content using TinyMCE API.
    Preserves exact paragraph spacing, fonts, bold, and lines from template.docx.
    """
    print("Step: fill body with template")

    if isinstance(body_data, dict):
        html_body = body_data.get("html", "")
        plain_body = body_data.get("plain", "")
    else:
        t_dict = _text_to_template_dict(str(body_data))
        html_body = t_dict["html"]
        plain_body = t_dict["plain"]

    # Proactively switch to HTML editor if Roundcube is in plain text mode
    try:
        driver.execute_script("""
            if (typeof rcmail !== 'undefined') {
                var isHtml = false;
                try {
                    if (typeof tinymce !== 'undefined' && tinymce.get('composebody') && tinymce.get('composebody').initialized) {
                        isHtml = true;
                    }
                } catch(e) {}
                if (!isHtml) {
                    try {
                        rcmail.command('toggle-editor', {id: 'composebody', html: true});
                    } catch(e) {}
                    var sel = document.querySelector('select[name="_is_html"], select#editor-selector, select[name="_editor_selector"]');
                    if (sel) {
                        for (var i = 0; i < sel.options.length; i++) {
                            if (sel.options[i].value === '1' || sel.options[i].value === 'html') {
                                if (sel.selectedIndex !== i) {
                                    sel.selectedIndex = i;
                                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                                break;
                            }
                        }
                    }
                }
            }
        """)
    except Exception:
        pass

    dismiss_any_dialogs(driver)

    # Wait for TinyMCE editor to be ready
    print("  Waiting for TinyMCE HTML editor to initialize...")
    deadline = time.time() + 20
    tinymce_ready = False
    while time.time() < deadline:
        try:
            ready = driver.execute_script("""
                return (typeof tinymce !== 'undefined'
                     && tinymce.get('composebody') !== null
                     && tinymce.get('composebody').initialized);
            """)
            if ready:
                tinymce_ready = True
                print("  TinyMCE HTML editor ready")
                break
        except Exception:
            pass
        time.sleep(0.5)

    time.sleep(1)

    if tinymce_ready:
        try:
            # Set HTML content in TinyMCE editor and sync to textarea
            driver.execute_script("""
                var editor = tinymce.get('composebody');
                editor.setContent(arguments[0]);
                editor.setDirty(true);
                if (editor.fire) { editor.fire('change'); }
                editor.save();
                var ta = document.getElementById('composebody');
                if (ta) {
                    ta.value = editor.getContent();
                    ta.dispatchEvent(new Event('input', {bubbles: true}));
                    ta.dispatchEvent(new Event('change', {bubbles: true}));
                }
            """, html_body)
            time.sleep(1)

            # Verify content
            verify = driver.execute_script("""
                var editor = tinymce.get('composebody');
                var editorContent = editor ? editor.getContent() : '';
                var ta = document.getElementById('composebody');
                var taContent = ta ? ta.value : '';
                return {editor: editorContent.length, textarea: taContent.length};
            """)
            print(f"  Body filled (HTML mode) — TinyMCE: {verify.get('editor', 0)} chars, Textarea: {verify.get('textarea', 0)} chars")

            if verify.get('editor', 0) > 10:
                return
            else:
                print("  TinyMCE content appears empty, trying textarea fallback...")
        except Exception as exc:
            print(f"  TinyMCE setContent failed: {exc}, trying fallback...")

    # FALLBACK: Plain text mode (clean paragraphs, not squashed)
    print("  Attempting plain text fallback...")
    try:
        driver.execute_script("""
            var ta = document.getElementById('composebody')
                  || document.querySelector("textarea[name='_message']");
            if (ta) {
                ta.value = arguments[0];
                ta.dispatchEvent(new Event('input', {bubbles: true}));
                ta.dispatchEvent(new Event('change', {bubbles: true}));
            }
        """, plain_body)
        print("  Body filled via plain text textarea fallback")
    except Exception as exc:
        print(f"  [!] ALL body fill methods failed: {exc}")


def set_priority_high(driver):
    """Set email priority to Highest in the compose form.
    The priority dropdown (#compose-priority) is inside the sidebar (#layout-sidebar)
    which is not always visible, so we use JavaScript to set it directly.
    From the actual HTML:
      <select name="_priority" id="compose-priority">
        <option value="5">Lowest</option>
        <option value="4">Low</option>
        <option value="0" selected>Normal</option>
        <option value="2">High</option>
        <option value="1">Highest</option>
      </select>
    """
    print("Step: set priority to Highest")
    try:
        # Use JavaScript directly — the dropdown is in the sidebar and may not be visible
        result = driver.execute_script("""
            var sel = document.querySelector('#compose-priority')
                   || document.querySelector("select[name='_priority']");
            if (!sel) return 'not_found';
            sel.value = '1';
            sel.dispatchEvent(new Event('change', {bubbles: true}));
            return sel.value;
        """)
        if result == '1':
            print("  Priority set to: Highest (value=1)")
        elif result == 'not_found':
            print("  [!] Priority dropdown not found in DOM")
        else:
            print(f"  [!] Priority value after set: {result} (expected '1')")
    except Exception as exc:
        print(f"  [!] Could not set priority: {exc}")


def click_send(driver):
    print("Step: find and click Send")
    try:
        safe_click(driver, visible_one(driver, SEND_CSS, SEND_XPATH, timeout=15))
    except Exception as exc:
        print(f"  Send button not found ({exc}) — using Ctrl+Enter shortcut")
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.CONTROL, Keys.ENTER)


def confirm_sent(driver, timeout=20):
    """Compose form disappearing (or a success toast) means the mail went out."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for el in driver.find_elements(By.CSS_SELECTOR, "#messagestack div, .ui.alert"):
                text = (el.text or "").lower()
                if "sent" in text:
                    return True
                if "error" in text or "failed" in text or "invalid" in text:
                    raise Exception(f"Webmail reported: {el.text.strip()}")
        except Exception as exc:
            if "Webmail reported" in str(exc):
                raise
        if not any_visible(driver, SUBJECT_CSS, SUBJECT_XPATH):
            return True
        time.sleep(0.5)
    return False


def send_one_email(driver, wait, to_email, subject, body_text, main_window):
    dismiss_any_dialogs(driver)
    compose_window = open_compose(driver, wait, main_window)
    driver.switch_to.window(compose_window)

    fill_to(driver, to_email)
    fill_subject(driver, subject)
    fill_body(driver, body_text, compose_window)  # paste the template into the body
    set_priority_high(driver)  # set email priority to High
    click_send(driver)

    if not confirm_sent(driver):
        raise Exception("Send clicked but the compose screen never closed")

    if compose_window != main_window and compose_window in driver.window_handles:
        try:
            driver.close()
        except Exception:
            pass
    driver.switch_to.window(main_window)
    dismiss_any_dialogs(driver)


def reset_to_inbox(driver, main_window):
    try:
        for handle in list(driver.window_handles):
            if handle != main_window:
                driver.switch_to.window(handle)
                driver.close()
        driver.switch_to.window(main_window)
        dismiss_any_dialogs(driver)
    except Exception:
        pass


# ---------------------------------------------------------------------------
def main():
    print("=== Browser Automation Mailer ===\n")

    # Load recipients — always process ALL from the beginning
    try:
        recipients = load_recipients()
    except Exception as exc:
        print(f"ERROR loading recipients: {exc}")
        return

    if not recipients:
        print("No valid recipients found. Nothing to send.")
        return

    # Load email template
    try:
        template_text = load_template()
    except Exception as exc:
        print(f"ERROR loading template: {exc}")
        return

    # Load custom subjects
    try:
        custom_subjects = load_subjects()
    except Exception as exc:
        print(f"ERROR loading subjects: {exc}")
        return

    # Always send to ALL recipients (no skipping based on send_log)
    todo = recipients
    print(f"\n{len(todo)} recipients to process (starting from the first).\n")

    password = (PASSWORD or "").strip()
    if not password:
        while not password:
            try:
                password = getpass.getpass(f"Password for {FROM_EMAIL} (hidden): ").strip()
            except Exception:
                pass
            if not password:
                ans = input(f"Password for {FROM_EMAIL} (or press Enter to type password directly in Chrome): ").strip()
                if ans:
                    password = ans
                else:
                    print("  You can type your password directly into the Chrome login screen.")
                    break

    driver = start_browser()
    wait = WebDriverWait(driver, 25)
    try:
        login(driver, wait, password)
        print("Logged in.\n")
    except Exception as exc:
        dump(driver, "login")
        print(f"LOGIN FAILED: {exc}")
        driver.quit()
        return

    main_window = driver.current_window_handle
    sent = 0
    for i, r in enumerate(todo):
        if custom_subjects:
            subject = random.choice(custom_subjects)
        else:
            subject = get_subject(r["name"], r["topic"], i)
        # Personalize the template: replace {name} and "Dear Researcher" with author name
        body_text = personalize_template(template_text, r["name"], r["paper_title"])
        try:
            send_one_email(driver, wait, r["email"], subject, body_text, main_window)
            log_result(LOG_FILE, r["email"], r["name"], subject, "sent")
            sent += 1
            print(f"[{i+1}/{len(todo)}] SENT -> {r['email']} | {subject}\n")
        except Exception as exc:
            err_str = str(exc)
            if is_fatal_browser_error(exc) or not is_driver_alive(driver):
                print(f"\n[CRITICAL ERROR] Chrome browser was closed or disconnected.")
                print(f"Stopping execution to protect your recipient list and prevent repeated failures.\n")
                log_result(LOG_FILE, r["email"], r["name"], subject, "failed", "Chrome closed or disconnected")
                break

            dump(driver, r["email"].replace("@", "_at_"))
            log_result(LOG_FILE, r["email"], r["name"], subject, "failed", err_str)
            print(f"[{i+1}/{len(todo)}] FAILED -> {r['email']}: {exc}\n")
            reset_to_inbox(driver, main_window)

        if i < len(todo) - 1:
            time.sleep(DELAY_BETWEEN_EMAILS_SECONDS)

    driver.quit()
    print(f"Done. {sent}/{len(todo)} sent this run. Log: {LOG_FILE}")


if __name__ == "__main__":
    main()
