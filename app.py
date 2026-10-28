#!/usr/bin/env python3
"""
Bigrock Bulk Mailer Web Application Server
FastAPI backend providing REST APIs, SSE real-time log streaming, and serving the Web UI.
"""

import threading
import asyncio
import csv
import json
import os
import shutil
import time
import webbrowser
from pathlib import Path
from typing import Optional, List, Dict, Any

import uvicorn
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from mailer_engine import (
    MailerEngine,
    parse_recipients_file,
    parse_template_file,
    parse_subjects_file,
)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

DEFAULT_RECIPIENTS = BASE_DIR / "recipients.csv"
DEFAULT_TEMPLATE = BASE_DIR / "template.docx"
DEFAULT_SUBJECT = BASE_DIR / "subject.docx"
SEND_LOG_FILE = BASE_DIR / "send_log.csv"

# Global Mailer Engine instance
engine = MailerEngine()

# In-memory log buffer & SSE subscriber queues
MAX_LOG_HISTORY = 300
log_history: List[Dict[str, str]] = []
log_subscribers: List[asyncio.Queue] = []
async_loop: Optional[asyncio.AbstractEventLoop] = None


def broadcast_log(message: str, level: str = "info"):
    entry = {"message": message, "level": level, "timestamp": time.strftime("%H:%M:%S")}
    log_history.append(entry)
    if len(log_history) > MAX_LOG_HISTORY:
        log_history.pop(0)

    # Dispatch to all SSE queues in the running event loop
    if async_loop and not async_loop.is_closed():
        for q in list(log_subscribers):
            try:
                async_loop.call_soon_threadsafe(q.put_nowait, entry)
            except Exception:
                pass


# Hook engine logs to SSE broadcast
engine.register_log_callback(broadcast_log)

# Active staged datasets (defaults or user-uploaded)
staged_data = {
    "recipients": [],
    "recipients_filename": "",
    "template_data": {"html": "", "plain": ""},
    "template_filename": "",
    "subjects": [],
    "subjects_filename": "",
}


def load_initial_defaults():
    # Priority 1: Check if user already uploaded recipients in uploads/
    uploaded_recipients = None
    for ext in (".csv", ".xlsx", ".xlsm"):
        cand = UPLOAD_DIR / f"uploaded_recipients{ext}"
        if cand.exists():
            uploaded_recipients = cand
            break

    target_recipients = uploaded_recipients if uploaded_recipients else DEFAULT_RECIPIENTS
    if target_recipients and target_recipients.exists():
        try:
            staged_data["recipients"] = parse_recipients_file(target_recipients)
            staged_data["recipients_filename"] = target_recipients.name
        except Exception as e:
            broadcast_log(f"Could not load recipients: {e}", "warning")

    # Priority 1: Check if user already uploaded template in uploads/
    uploaded_template = None
    for ext in (".docx", ".txt", ".doc", ".rtf"):
        cand = UPLOAD_DIR / f"uploaded_template{ext}"
        if cand.exists():
            uploaded_template = cand
            break

    target_template = uploaded_template if uploaded_template else DEFAULT_TEMPLATE
    if target_template and target_template.exists():
        try:
            staged_data["template_data"] = parse_template_file(target_template)
            staged_data["template_filename"] = target_template.name
        except Exception as e:
            broadcast_log(f"Could not load template: {e}", "warning")

    # Priority 1: Check if user already uploaded subjects in uploads/
    uploaded_subjects = None
    for ext in (".docx", ".txt", ".doc"):
        cand = UPLOAD_DIR / f"uploaded_subjects{ext}"
        if cand.exists():
            uploaded_subjects = cand
            break

    target_subjects = uploaded_subjects if uploaded_subjects else DEFAULT_SUBJECT
    if target_subjects and target_subjects.exists():
        try:
            staged_data["subjects"] = parse_subjects_file(target_subjects)
            staged_data["subjects_filename"] = target_subjects.name
        except Exception as e:
            broadcast_log(f"Could not load subjects: {e}", "warning")


load_initial_defaults()

app = FastAPI(title="Bigrock Bulk Mailer Web Dashboard", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# API MODELS
# ---------------------------------------------------------------------------
class StartJobRequest(BaseModel):
    from_email: str
    password: str
    webmail_url: Optional[str] = "https://us3.webmail.mailhostbox.com"
    delay_seconds: Optional[float] = 30.0
    headless: Optional[bool] = False


# ---------------------------------------------------------------------------
# REST ENDPOINTS
# ---------------------------------------------------------------------------
@app.get("/api/defaults")
def get_defaults():
    sample_recipients = staged_data["recipients"][:5]
    return {
        "from_email": "oralhealth@scienceresearchpub.org",
        "webmail_url": "https://us3.webmail.mailhostbox.com",
        "delay_seconds": 30,
        "recipients_count": len(staged_data["recipients"]),
        "recipients_filename": staged_data["recipients_filename"],
        "sample_recipients": sample_recipients,
        "template_filename": staged_data["template_filename"],
        "template_plain_preview": staged_data["template_data"].get("plain", "")[:400],
        "template_html_preview": staged_data["template_data"].get("html", ""),
        "subjects_count": len(staged_data["subjects"]),
        "subjects_filename": staged_data["subjects_filename"],
        "subjects_list": staged_data["subjects"][:5],
    }


@app.get("/api/status")
def get_status():
    status_dict = engine.get_status_dict()
    status_dict["recipients_loaded"] = len(staged_data["recipients"])
    status_dict["template_loaded"] = bool(staged_data["template_data"].get("plain"))
    status_dict["subjects_loaded"] = len(staged_data["subjects"])
    return status_dict


@app.post("/api/upload/recipients")
async def upload_recipients(file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".csv", ".xlsx", ".xlsm"):
        raise HTTPException(status_code=400, detail="Allowed file types: .csv, .xlsx, .xlsm")

    save_path = UPLOAD_DIR / f"uploaded_recipients{ext}"
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        rows = parse_recipients_file(save_path)
        if not rows:
            raise HTTPException(status_code=400, detail="No valid rows with emails found in file.")
        staged_data["recipients"] = rows
        staged_data["recipients_filename"] = file.filename
        broadcast_log(f"Uploaded recipients: {file.filename} ({len(rows)} recipients)", "success")
        return {
            "success": True,
            "filename": file.filename,
            "count": len(rows),
            "sample": rows[:5],
        }
    except Exception as exc:
        broadcast_log(f"Error parsing recipients file: {exc}", "error")
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/upload/template")
async def upload_template(file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".docx", ".txt", ".doc", ".rtf"):
        raise HTTPException(status_code=400, detail="Allowed file types: .docx, .txt")

    save_path = UPLOAD_DIR / f"uploaded_template{ext}"
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        t_data = parse_template_file(save_path)
        if not t_data.get("plain", "").strip():
            raise HTTPException(status_code=400, detail="Template file appears to be empty.")
        staged_data["template_data"] = t_data
        staged_data["template_filename"] = file.filename
        broadcast_log(f"Uploaded template: {file.filename}", "success")
        return {
            "success": True,
            "filename": file.filename,
            "plain_preview": t_data.get("plain", "")[:400],
            "html_preview": t_data.get("html", ""),
        }
    except Exception as exc:
        broadcast_log(f"Error parsing template: {exc}", "error")
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/upload/subjects")
async def upload_subjects(file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in (".docx", ".txt", ".doc"):
        raise HTTPException(status_code=400, detail="Allowed file types: .docx, .txt")

    save_path = UPLOAD_DIR / f"uploaded_subjects{ext}"
    with open(save_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        subjects = parse_subjects_file(save_path)
        staged_data["subjects"] = subjects
        staged_data["subjects_filename"] = file.filename
        broadcast_log(f"Uploaded subjects: {file.filename} ({len(subjects)} subjects)", "success")
        return {
            "success": True,
            "filename": file.filename,
            "count": len(subjects),
            "subjects": subjects[:5],
        }
    except Exception as exc:
        broadcast_log(f"Error parsing subjects: {exc}", "error")
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/start")
def start_sending(req: StartJobRequest):
    if not req.from_email or "@" not in req.from_email:
        raise HTTPException(status_code=400, detail="Please enter a valid sender email address.")
    if not req.password:
        raise HTTPException(status_code=400, detail="Please enter your email account password.")
    if not staged_data["recipients"]:
        raise HTTPException(status_code=400, detail="No recipients loaded. Please upload a recipients file.")
    if not staged_data["template_data"].get("plain"):
        raise HTTPException(status_code=400, detail="No template loaded. Please upload a template file.")

    config = {
        "from_email": req.from_email.strip(),
        "password": req.password,
        "webmail_url": req.webmail_url.strip(),
        "delay_seconds": req.delay_seconds,
        "headless": req.headless,
        "recipients": staged_data["recipients"],
        "template_data": staged_data["template_data"],
        "custom_subjects": staged_data["subjects"],
        "log_file": str(SEND_LOG_FILE),
    }

    try:
        engine.start(config)
        broadcast_log("Mailer process initiated from Web Dashboard.", "info")
        return {"success": True, "message": "Bulk mailer started."}
    except Exception as exc:
        broadcast_log(f"Failed to start mailer: {exc}", "error")
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/pause")
def pause_sending():
    engine.pause()
    return {"success": True, "message": "Pause requested."}


@app.post("/api/resume")
def resume_sending():
    engine.resume()
    return {"success": True, "message": "Process resumed."}


@app.post("/api/stop")
def stop_sending():
    engine.stop()
    return {"success": True, "message": "Stop requested."}


@app.get("/api/sent-log")
def get_sent_log(limit: int = 500, query: str = "", status: str = "", date: str = ""):
    files_to_check = [SEND_LOG_FILE, BASE_DIR / "send_log_pending.csv"]
    rows = []
    seen = set()
    available_dates = set()

    for fpath in files_to_check:
        if not fpath.exists():
            continue
        try:
            with open(fpath, newline="", encoding="utf-8", errors="replace") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    if not r or not r.get("email"):
                        continue
                    key = (r.get("timestamp", ""), r.get("email", ""), r.get("status", ""))
                    if key in seen:
                        continue
                    seen.add(key)

                    ts = (r.get("timestamp") or "").strip()
                    # Extract date token for available_dates
                    if ts:
                        d_token = ts.split(" ")[0].strip()
                        if d_token:
                            available_dates.add(d_token)

                    # Date filter
                    if date:
                        req_date = date.strip()
                        match_date = False
                        if req_date in ts:
                            match_date = True
                        elif "-" in req_date:
                            parts = req_date.split("-")
                            if len(parts) == 3:
                                # if YYYY-MM-DD -> match DD-MM-YYYY
                                alt1 = f"{parts[2]}-{parts[1]}-{parts[0]}"
                                # if DD-MM-YYYY -> match YYYY-MM-DD
                                alt2 = f"{parts[0]}-{parts[1]}-{parts[2]}"
                                if alt1 in ts or alt2 in ts:
                                    match_date = True
                        if not match_date:
                            continue

                    # Status filter
                    if status and (r.get("status") or "").lower() != status.lower():
                        continue

                    # Text search
                    if query:
                        q = query.lower()
                        if not (
                            q in (r.get("email") or "").lower()
                            or q in (r.get("name") or "").lower()
                            or q in (r.get("subject") or "").lower()
                            or q in ts.lower()
                        ):
                            continue

                    rows.append(r)
        except Exception as exc:
            broadcast_log(f"Error reading log file {fpath.name}: {exc}", "error")

    # Return newest first
    rows.reverse()
    sorted_dates = sorted(list(available_dates), reverse=True)
    return {"rows": rows[:limit], "total": len(rows), "available_dates": sorted_dates}


@app.get("/api/export-log")
def export_log():
    if not SEND_LOG_FILE.exists():
        raise HTTPException(status_code=404, detail="No send_log.csv found.")
    return FileResponse(
        SEND_LOG_FILE,
        media_type="text/csv",
        filename=f"send_log_{time.strftime('%Y%m%d_%H%M%S')}.csv",
    )


@app.get("/api/logs/stream")
async def stream_logs():
    q: asyncio.Queue = asyncio.Queue()
    # Send historical logs first
    for item in log_history[-50:]:
        await q.put(item)
    log_subscribers.append(q)

    async def event_generator():
        try:
            while True:
                item = await q.get()
                yield f"data: {json.dumps(item)}\n\n"
        except asyncio.CancelledError:
            pass
        finally:
            if q in log_subscribers:
                log_subscribers.remove(q)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


# Serve root and static assets
@app.get("/")
def serve_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return {"message": "Web UI frontend is building..."}


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


# ---------------------------------------------------------------------------
# LAUNCHER HELPER
# ---------------------------------------------------------------------------
def open_browser_delayed(url: str, delay: float = 1.2):
    def _open():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=_open, daemon=True).start()


if __name__ == "__main__":
    port = 8000
    host = "127.0.0.1"
    url = f"http://{host}:{port}"
    print(f"\n=======================================================")
    print(f"  Bigrock Bulk Mailer Web Interface")
    print(f"  Server running at: {url}")
    print(f"  Opening default web browser...")
    print(f"=======================================================\n")

    open_browser_delayed(url)

    # Capture current event loop reference for SSE broadcasting
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    async_loop = loop

    config = uvicorn.Config(app=app, host=host, port=port, log_level="info", loop="asyncio")
    server = uvicorn.Server(config)
    loop.run_until_complete(server.serve())
