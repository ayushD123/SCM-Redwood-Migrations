from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form, BackgroundTasks, Cookie
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import os
import sys
import shutil
import glob
import subprocess
import zipfile
from pathlib import Path
import asyncio
import uuid
import json
from typing import Dict, List, Optional, Tuple
import logging
from datetime import datetime
import traceback
import tempfile
import threading
import time
import pandas as pd
import re
import requests

APP_DIR = Path(__file__).resolve().parent


def _normalize_base_path(value: str) -> str:
    normalized = f"/{str(value or '').strip().strip('/')}"
    return "" if normalized == "/" else normalized


APP_BASE_PATH = _normalize_base_path(os.getenv("APP_BASE_PATH", ""))
ARTIFACT_RETENTION_MINUTES = max(5, int(os.getenv("ARTIFACT_RETENTION_MINUTES", "120")))
ARTIFACT_RETENTION_SECONDS = ARTIFACT_RETENTION_MINUTES * 60

df = pd.read_excel(APP_DIR / "static/Redwood_SCM_Features_V1.11.xlsx", sheet_name=0)
table_html = df.to_html(index=False, classes="excel-table")
excel_template = """<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"UTF-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>Redwood SCM Features</title>
  <style>
    body {
      margin: 0;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: #f4f4f4;
      color: #111;
      padding: 1.5rem;
    }
    h1 {
      font-size: 1.8rem;
      margin-bottom: 1rem;
    }
    .table-wrapper {
      background: white;
      border-radius: 12px;
      padding: 1rem;
      box-shadow: 0 12px 28px rgba(0,0,0,0.08);
      overflow-x: auto;
    }
    table.excel-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.95rem;
    }
    table.excel-table th,
    table.excel-table td {
      border: 1px solid #d0d0d0;
      padding: 0.65rem 0.8rem;
      text-align: left;
    }
    table.excel-table th {
      background: #333;
      color: white;
      font-weight: 600;
    }
  </style>
</head>
<body>
  <h1>Redwood SCM Features</h1>
  <div class=\"table-wrapper\">
    {table_html}
  </div>
  <script>window.REDWOOD_BASE_PATH = \"{app_base_path}\";</script>
  <script src=\"{app_base_path}/static/session.js\"></script>
  <script>window.sessionHelper.requireSession('/');</script>
</body>
</html>"""
class SaveSessionPayload(BaseModel):
    user: str
    customer: str


def _sanitize_filename(value: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9]+", "_", value or "")
    return safe.strip("_") or "session"


def _zip_with_customer_root(source_zip: Path, session_id: str, session: Dict, label: str = "output") -> Path:
    """Return a zip whose top-level folder is the sanitized customer name."""
    if not source_zip.exists():
        raise HTTPException(status_code=404, detail="Output ZIP not found")

    customer_root = _sanitize_filename(session.get("customer", ""))
    wrapped_zip = APP_DIR / f"{label}_{session_id}_customer_root.zip"
    if wrapped_zip.exists():
        wrapped_zip.unlink()

    with zipfile.ZipFile(source_zip, "r") as src, zipfile.ZipFile(wrapped_zip, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            if info.is_dir():
                continue

            name = info.filename.replace("\\", "/").lstrip("/")
            if not name or name.startswith("__MACOSX/"):
                continue

            parts = [part for part in name.split("/") if part and part != "."]
            if not parts:
                continue

            if parts[0] == customer_root:
                arcname = "/".join(parts)
            else:
                arcname = f"{customer_root}/{'/'.join(parts)}"

            dst.writestr(arcname, src.read(info.filename))

    return wrapped_zip



# --- force local folder onto sys.path so same-dir modules import reliably ---
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# Import modular components for Phase 3
try:
    from page_configurations import (
        get_page_config,
        get_composite_mapping,
        list_available_pages,
        validate_page_exists
    )
    #NON SSO
    # from automation_engine import (
    #     handle_sso_login,
    #     handle_mfa_login,
    #     navigate_to_page,
    #     open_visual_builder,
    #     select_vb_project,
    #     navigate_to_vb_page,
    #     capture_vb_urls,
    #     create_customization_rule,
    #     activate_advanced_mode,
    #     open_metadata_file,
    #     process_excel_mapping,
    #     inject_personalization_to_monaco
    # )
    PHASE3_AVAILABLE = True
except ImportError as e:
    logger.warning(f"Phase 3 modules not available: {e}")
    PHASE3_AVAILABLE = False

app = FastAPI(title="Oracle Migration Tool API", root_path=APP_BASE_PATH)

# Serve static HTML pages
app.mount("/static", StaticFiles(directory=APP_DIR / "static"), name="static")


def _html_page(filename: str) -> HTMLResponse:
    html = (APP_DIR / "static" / filename).read_text(encoding="utf-8")
    html = html.replace("__APP_BASE_PATH__", APP_BASE_PATH)
    return HTMLResponse(html)

@app.get("/")
async def root():
    return _html_page("index.html")

@app.get("/dashboard")
async def root():
    response = _html_page("dashboard.html")
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

@app.get("/phase2")
async def phase2_page():
    return _html_page("phase2.html")


@app.get("/phase1")
async def phase1_page():
    response = _html_page("phase1.html")
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.get("/excel")
async def excel_page():
    html = excel_template.replace("{table_html}", table_html).replace("{app_base_path}", APP_BASE_PATH)
    return HTMLResponse(html)


# Client-specific temporary removal: SCM Redwood Customizations Tool.
# Uncomment this route to restore the feature.
# @app.get("/phase3")
# async def phase3_page():
#     return FileResponse("static/phase3.html")

# Client-specific temporary removal: SCM Redwood Estimation Tool.
# Uncomment this route to restore the feature.
# @app.get("/estimation")
# async def estimation_page():
#     return FileResponse("static/estimation.html")

@app.get("/api/health")
async def health_check():
    """Health check endpoint to verify server and dependencies"""
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
        playwright_available = True
    except ImportError:
        playwright_available = False

    try:
        import pandas as pd
        pandas_available = True
    except ImportError:
        pandas_available = False

    current_job = _active_job_snapshot()
    return {
        "status": "healthy",
        "playwright_available": playwright_available,
        "pandas_available": pandas_available,
        "phase3_available": PHASE3_AVAILABLE,
        "active_sessions": len(sessions),
        "tool_jar_exists": (APP_DIR / TOOL_JAR).exists(),
        "busy": current_job is not None,
        "active_job": current_job,
        "base_path": APP_BASE_PATH,
        "artifact_retention_minutes": ARTIFACT_RETENTION_MINUTES,
    }


@app.get("/api/job/status")
async def job_status():
    current_job = _active_job_snapshot()
    return {"busy": current_job is not None, "active_job": current_job}

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def no_cache_phase1_assets(request: Request, call_next):
    response = await call_next(request)
    if request.url.path in {"/phase1", "/static/phase1.html"}:
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response

# # Proxy configuration - Oracle Corporate Proxy
# PROXY_SERVER = "http://www-proxy-hqdc.us.oracle.com:80"
# PROXY_BYPASS = "oraclecorp.com,oraclevcn.com,oraclecloud.com,169.254.169.254,localhost,127.0.0.1,0.0.0.0"

# # Set environment variables for requests/urllib (used by some libraries)
# os.environ['http_proxy'] = PROXY_SERVER
# os.environ['https_proxy'] = PROXY_SERVER
# os.environ['no_proxy'] = PROXY_BYPASS

# Tool JAR name
TOOL_JAR = "ScmRedwoodCustHelper-24.04.jar"

# In-memory storage
sessions: Dict[str, Dict] = {}
job_logs: Dict[str, List[str]] = {}
job_control_lock = threading.Lock()
active_job: Optional[Dict[str, str]] = None
cleanup_timers: Dict[str, threading.Timer] = {}
cleanup_stop_event = threading.Event()


def _active_job_snapshot() -> Optional[Dict[str, str]]:
    with job_control_lock:
        return dict(active_job) if active_job else None


def _reserve_job(session_id: str, phase: str) -> bool:
    global active_job
    with job_control_lock:
        if active_job is not None:
            return False
        previous_cleanup = cleanup_timers.pop(session_id, None)
        if previous_cleanup:
            previous_cleanup.cancel()
        active_job = {
            "session_id": session_id,
            "phase": phase,
            "started_at": datetime.now().isoformat(),
        }
        return True


def _release_job(session_id: str) -> None:
    global active_job
    with job_control_lock:
        if active_job and active_job.get("session_id") == session_id:
            active_job = None


def _remove_path(path: Path) -> None:
    try:
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        elif path.exists():
            path.unlink(missing_ok=True)
    except Exception as exc:
        logger.warning("Unable to remove temporary artifact %s: %s", path, exc)


def _cleanup_session_artifacts(session_id: str, remove_state: bool = True) -> None:
    session = sessions.get(session_id, {})
    explicit_paths = [
        session.get("upload_dir"),
        session.get("phase1_run_dir"),
        session.get("output_zip"),
        session.get("ppt_path"),
    ]
    for value in explicit_paths:
        if value:
            _remove_path(Path(value))

    for path in [
        APP_DIR / f"work_{session_id}",
        APP_DIR / f"output_{session_id}",
        APP_DIR / f"downloads_{session_id}",
        APP_DIR / f"uploads_{session_id}",
        APP_DIR / f"upload_{session_id}.jar",
        APP_DIR / f"output_{session_id}.zip",
        APP_DIR / f"phase1_logs_{session_id}.zip",
        APP_DIR / f"phase1_output_{session_id}.zip",
    ]:
        _remove_path(path)

    for wrapped_zip in APP_DIR.glob(f"*_{session_id}_customer_root.zip"):
        _remove_path(wrapped_zip)

    with job_control_lock:
        timer = cleanup_timers.pop(session_id, None)
        if timer and timer is not threading.current_thread():
            timer.cancel()
        if remove_state:
            sessions.pop(session_id, None)
            job_logs.pop(session_id, None)


def _schedule_session_cleanup(session_id: str) -> None:
    session = sessions.get(session_id)
    if session is not None:
        session["expires_at"] = datetime.fromtimestamp(
            time.time() + ARTIFACT_RETENTION_SECONDS
        ).isoformat()

    timer = threading.Timer(
        ARTIFACT_RETENTION_SECONDS,
        _cleanup_session_artifacts,
        args=(session_id,),
    )
    timer.daemon = True
    with job_control_lock:
        previous = cleanup_timers.pop(session_id, None)
        if previous:
            previous.cancel()
        cleanup_timers[session_id] = timer
    timer.start()


def _cleanup_stale_runtime_artifacts() -> None:
    cutoff = time.time() - ARTIFACT_RETENTION_SECONDS
    protected_session_ids = set(sessions)
    protected_paths = set()
    for session in sessions.values():
        for key in ("upload_dir", "phase1_run_dir", "output_zip", "ppt_path"):
            value = session.get(key)
            if value:
                try:
                    protected_paths.add(Path(value).resolve())
                except (OSError, RuntimeError):
                    continue

    patterns = [
        "work_*",
        "output_*",
        "downloads_*",
        "uploads_*",
        "upload_*.jar",
        "phase1_*.zip",
        "*_customer_root.zip",
        "logs/run_*",
    ]
    for pattern in patterns:
        for path in APP_DIR.glob(pattern):
            try:
                resolved_path = path.resolve()
                if any(session_id in str(resolved_path) for session_id in protected_session_ids):
                    continue
                if any(
                    resolved_path == protected_path or protected_path in resolved_path.parents
                    for protected_path in protected_paths
                ):
                    continue
                if path.stat().st_mtime < cutoff:
                    _remove_path(path)
            except FileNotFoundError:
                continue


def _periodic_artifact_cleanup() -> None:
    scan_interval_seconds = min(300, max(60, ARTIFACT_RETENTION_SECONDS // 4))
    while not cleanup_stop_event.wait(scan_interval_seconds):
        _cleanup_stale_runtime_artifacts()


@app.on_event("startup")
async def cleanup_stale_artifacts_on_startup() -> None:
    cleanup_stop_event.clear()
    _cleanup_stale_runtime_artifacts()
    cleanup_thread = threading.Thread(
        target=_periodic_artifact_cleanup,
        name="redwood-artifact-cleanup",
        daemon=True,
    )
    cleanup_thread.start()


@app.on_event("shutdown")
async def stop_artifact_cleanup() -> None:
    cleanup_stop_event.set()


@app.post("/api/save-session")
async def save_session(payload: SaveSessionPayload, response: Response):
    """Create a lightweight client session and persist it via cookie"""
    session_id = str(uuid.uuid4())
    sessions[session_id] = {
        "phase": "intro",
        "status": "initialized",
        "created_at": datetime.now(),
        "user": payload.user,
        "customer": payload.customer
    }
    # user = sessions[session_id]["user"]
    # customer = sessions[session_id]["customer"]
    job_logs[session_id] = [f"[{datetime.now()}] ✅ Session created for {payload.user}"]
    response.set_cookie(
        key="session_id",
        value=session_id,
        httponly=True,
        samesite="lax",
        path=APP_BASE_PATH or "/"
    )
    return {
        "has_session": True,
        "session_id": session_id,
        "status": "created",
        "message": "Session initialized"
    }


@app.get("/api/session/status")
async def session_status(request: Request):
    session_id = request.cookies.get("session_id") or request.headers.get("X-Session-Id")
    if not session_id or session_id not in sessions:
        return {"has_session": False}
    session_data = sessions[session_id].copy()
    session_data["session_id"] = session_id
    session_data["has_session"] = True
    return session_data

@app.get("/api/session-info")
async def get_session_info(session_id: str = Cookie(None)):
    if not session_id or session_id not in sessions:
        raise HTTPException(status_code=401, detail="Invalid session")

    session = sessions[session_id]
    print(session)
    return {
        "user": session.get("user"),
        "customer": session.get("customer")
    }
# ==================== GLOBAL HELPER FUNCTIONS ====================



def update_phase_in_sheet(customer_name, username, phase):
    WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyZvNj2NRxb66DbjxDRqZkZ9kZUCLo2hREqYe-NVKIVsUaAERLcct8Xs6LyE2hvC_OrrA/exec"
    try:
        response = requests.post(
            WEB_APP_URL,
            json={
                "customerName": customer_name,
                "username": username,
                "phase": phase
            },
            timeout=10
        )

        print("Sheet update response:", response.text)

    except Exception as e:
        print("Error updating sheet:", str(e))

def generate_field_mapping_report(
    session_id: str,
    excel_file: str,
    page_name: str,
    mapped_records: List[tuple],
    unmapped_records: List[tuple],
    log_queue: List[str]
) -> Path:
    """
    Global function to generate field mapping report for any VB page automation
    """
    mapped_count = len(mapped_records)
    unmapped_count = len(unmapped_records)

    log_queue.append("📝 Generating field mapping report...")

    safe_page_name = page_name.lower().replace(" ", "_").replace("-", "")
    report_filename = f"{safe_page_name}_mapping_report_{session_id}.txt"
    report_path = Path(report_filename)

    with open(report_path, 'w', encoding='utf-8') as log_file:
        log_file.write("=" * 120 + "\n")
        log_file.write(f"FIELD MAPPING REPORT - {page_name.upper()}\n")
        log_file.write("=" * 120 + "\n")
        log_file.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        log_file.write(f"Session ID: {session_id}\n")
        log_file.write(f"Excel File: {excel_file}\n")
        log_file.write(f"Target Page: {page_name}\n\n")

        log_file.write(f"MAPPED ({mapped_count}):\n")
        log_file.write(f"{'File':<30} {'Field':<25} {'Display Name':<30} {'To'}\n")
        log_file.write("-" * 120 + "\n")
        for file, field, disp, mapped_to in mapped_records:
            log_file.write(f"{file:<30} {field:<25} {disp:<30} {mapped_to}\n")

        if unmapped_records:
            log_file.write(f"\nUNMAPPED ({unmapped_count}):\n")
            log_file.write(f"{'File':<30} {'Field':<25} {'Display Name':<30} {'Personalization'}\n")
            log_file.write("-" * 120 + "\n")
            for file, field, disp, pers in unmapped_records:
                log_file.write(f"{file:<30} {field:<25} {disp:<30} {pers}\n")

        log_file.write("\n" + "=" * 120 + "\n")
        log_file.write("SUMMARY:\n")
        log_file.write(f"  Total Mapped Fields: {mapped_count}\n")
        log_file.write(f"  Total Unmapped Fields: {unmapped_count}\n")
        if (mapped_count + unmapped_count) > 0:
            success_rate = (mapped_count / (mapped_count + unmapped_count) * 100)
            log_file.write(f"  Success Rate: {success_rate:.1f}%\n")
        else:
            log_file.write("  Success Rate: N/A\n")
        log_file.write("=" * 120 + "\n")

    log_queue.append(f"✅ Mapping report saved: {report_filename}")
    return report_path


# ==================== PYDANTIC MODELS ====================

# class Phase3BrowserRequest(BaseModel):
#     mode: str = "browser"
#     url: str
#     username: str
#     password: str
#     project_name: str
#     page_id: str
#     excel_file_path: Optional[str] = None
#     requires_mfa: bool = False
#     mfa_code: Optional[str] = None

class Phase3ZipRequest(BaseModel):
    mode: str = "zip"
    page_id: str
    excel_filename: str
    zip_filename: str


class Phase1StartRequest(BaseModel):
    fusion_url: str
    username: str
    password: str
    input_excel_path: str
    run_only: Optional[str] = None
    auto_only: Optional[bool] = None


def _normalize_phase1_run_only(value: Optional[str]) -> str:
    normalized = " ".join(str(value or "").replace("_", " ").replace("-", " ").split()).casefold()
    aliases = {
        "": "all",
        "all": "all",
        "opt in": "optin",
        "optin": "optin",
        "profile options": "profile",
        "profile option": "profile",
        "profile": "profile",
        "ess jobs": "ess",
        "ess job": "ess",
        "ess": "ess",
        "auto": "auto",
    }
    if normalized not in aliases:
        raise HTTPException(status_code=400, detail="Invalid Run only option")
    return aliases[normalized]

class SessionResponse(BaseModel):
    session_id: str
    status: str
    message: str

class Phase3Response(BaseModel):
    session_id: str
    status: str
    mode: str
    message: str
    user: str
    customer: str

#

@app.post("/api/phase2/process")
async def process_phase2(
    request: Request,
    session_id: str = None,
    file: UploadFile = File(...),
    background_tasks: BackgroundTasks = None
):
    """Process uploaded JAR file in Phase 2"""
    if not (file.filename or "").lower().endswith('.jar'):
        raise HTTPException(status_code=400, detail="Only JAR files are allowed")

    sid = (
        session_id
        or request.cookies.get("session_id")
        or request.headers.get("X-Session-Id")
    )
    if not sid or sid not in sessions:
        raise HTTPException(status_code=401, detail="Invalid session")
    if not _reserve_job(sid, "phase2"):
        raise HTTPException(
            status_code=409,
            detail="Server is busy with another automation. Please try again after it finishes.",
        )

    work_dir = APP_DIR / f"work_{sid}"
    jar_path = work_dir / "temp.jar"
    try:
        existing_session = sessions[sid]
        user = existing_session.get("user")
        customer = existing_session.get("customer")
        existing_session.update({
            "phase": 2,
            "status": "processing",
            "created_at": datetime.now(),
            "filename": file.filename,
        })
        job_logs[sid] = []

        work_dir.mkdir(exist_ok=True)
        with jar_path.open("wb") as destination:
            shutil.copyfileobj(file.file, destination)
        if jar_path.stat().st_size == 0:
            raise ValueError("Uploaded JAR is empty")

        background_tasks.add_task(process_jar_background, sid, jar_path)
    except Exception as exc:
        _release_job(sid)
        _remove_path(work_dir)
        sessions[sid]["status"] = "error"
        sessions[sid]["error"] = str(exc)
        raise HTTPException(status_code=500, detail=f"Unable to start Phase 2: {exc}") from exc
    finally:
        await file.close()

    return {
        "session_id": sid,

        "status": "processing",
        "message": "JAR file uploaded and processing started",
        "user": user,
        "customer": customer

    }


async def process_jar_background(session_id: str, jar_path: Path):
    """Windows-optimized Phase 2 processing"""
    log_queue = job_logs[session_id]
    work_dir = (APP_DIR / f"work_{session_id}").resolve()
    output_dir = (APP_DIR / f"output_{session_id}").resolve()
    zip_path = (APP_DIR / f"output_{session_id}.zip").resolve()

    # Save uploaded JAR to a safe temp location BEFORE cleanup
    temp_jar_safe = (APP_DIR / f"upload_{session_id}.jar").resolve()

    try:
        log_queue.append(f"[{datetime.now()}] 🚀 Starting Phase 2 processing (Windows-optimized)")
        log_queue.append(f"📦 Source JAR: {jar_path.resolve()}")

        if not jar_path.exists():
            raise FileNotFoundError(f"Uploaded JAR not found: {jar_path}")
        if jar_path.stat().st_size == 0:
            raise ValueError("Uploaded JAR is empty")
        log_queue.append(f"✅ Validated source JAR ({jar_path.stat().st_size:,} bytes)")

        shutil.copy2(jar_path, temp_jar_safe)
        log_queue.append(f"✅ JAR backed up to: {temp_jar_safe.name}")

        # === CLEANUP PREVIOUS RUNS (WINDOWS-SAFE) ===
        for dir_path in [work_dir, output_dir]:
            if dir_path.exists():
                for attempt in range(3):
                    try:
                        shutil.rmtree(dir_path, ignore_errors=False)
                        break
                    except Exception as e:
                        log_queue.append(f"⚠️ Cleanup retry {attempt+1}/3 for {dir_path.name}: {type(e).__name__}")
                        await asyncio.sleep(0.3)

        work_dir.mkdir(exist_ok=True)
        output_dir.mkdir(exist_ok=True)
        log_queue.append(f"📁 Workspace created: {work_dir}")

        # === EXTRACT UPLOADED JAR ===
        log_queue.append(f"📦 Extracting {temp_jar_safe.name}...")
        try:
            with zipfile.ZipFile(temp_jar_safe, 'r') as zip_ref:
                for zip_info in zip_ref.infolist():
                    if zip_info.filename.startswith('__MACOSX') or zip_info.filename.endswith('/'):
                        continue
                    target_path = work_dir / Path(zip_info.filename.replace('\\', '/'))
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    if not zip_info.is_dir():
                        with zip_ref.open(zip_info) as src, open(target_path, 'wb') as dst:
                            shutil.copyfileobj(src, dst)
        except Exception as e:
            raise RuntimeError(f"JAR extraction failed: {str(e)}") from e
        log_queue.append("✅ JAR extracted successfully")

        # === LOCATE ADF DIRECTORY (CASE-INSENSITIVE) ===
        adf_dir = None
        for item in work_dir.iterdir():
            if item.is_dir() and item.name.lower() == "adf":
                adf_dir = item
                if item.name != "adf":
                    try:
                        new_path = work_dir / "adf"
                        item.rename(new_path)
                        adf_dir = new_path
                        log_queue.append(f"🔄 Normalized directory name: '{item.name}' → 'adf'")
                    except Exception as e:
                        log_queue.append(f"⚠️ Rename skipped: {e}")
                break

        if not adf_dir or not adf_dir.exists():
            candidates = [d for d in work_dir.rglob("*") if d.is_dir() and d.name.lower() == "adf"]
            if candidates:
                adf_dir = candidates[0]
                log_queue.append(f"🔍 Found adf in subpath: {adf_dir.relative_to(work_dir)}")
            else:
                raise FileNotFoundError("No 'adf' directory found in JAR contents")
        log_queue.append(f"✅ adf directory: {adf_dir.relative_to(work_dir)}")

        # === FIND OUTER JAR (CS_ADF_*.jar) ===
        outer_jars = [
            f for f in adf_dir.iterdir()
            if f.is_file() and f.name.lower().startswith("cs_adf_") and f.name.lower().endswith(".jar")
        ]
        if not outer_jars:
            raise FileNotFoundError("No CS_ADF_*.jar found in adf/ directory")
        outer_jar = outer_jars[0]
        log_queue.append(f"✅ Outer JAR: {outer_jar.name}")

        # === EXTRACT INNER MDS JAR ===
        inner_jar_basename = None
        with zipfile.ZipFile(outer_jar, 'r') as zip_ref:
            for name in zip_ref.namelist():
                if '_MDS_' in name and name.lower().endswith('.jar'):
                    inner_jar_basename = Path(name).name
                    inner_jar_path = adf_dir / inner_jar_basename
                    with zip_ref.open(name) as src, open(inner_jar_path, 'wb') as dst:
                        shutil.copyfileobj(src, dst)
                    log_queue.append(f"✅ Extracted inner MDS JAR: {inner_jar_basename}")
                    break

        if not inner_jar_basename:
            raise FileNotFoundError("No *_MDS_*.jar found inside outer JAR")

        # === PRE-CREATE OUTPUT DIRECTORY STRUCTURE ===
        tool_output_dir = work_dir / "output"
        tool_output_adf_dir = tool_output_dir / "adf"
        tool_output_dir.mkdir(exist_ok=True)
        tool_output_adf_dir.mkdir(exist_ok=True)
        log_queue.append(f"✅ Pre-created output directories: output/adf/")

        # === CREATE InputParams.txt ===
        mds_relative_path = f"adf/{inner_jar_basename}"
        input_params = """Modulename=Purchasing,Self Service Procurement,Sourcing,Spend Classification,Supplier Model,Supplier Portal,Supplier Qualification,Asset Tracking,Installed Base,Maintenance Management,Service Logistics,Cost Management,Fiscal Document Capture,Landed Cost Management,Quality Inspection Management,Receipt Accounting,Cloud ESG,Supply Chain Localization,Inventory Management,Receiving,Shipping,Common Work Execution,Common Work Setup,Discrete Manufacturing,E-Signatures and E-Records,Channel Revenue Management,Configurator,Configure To Order,Distributed Order Orchestration,Order Management,Pricing,Supply Chain Orchestration,Enterprise Catalog for Communications,Enterprise Visualization,Product and Catalog Management,Product Concept Design,Product Development,Product Hub,Product Hub Portal,Product Lifecycle Portfolio Management,Product Model,Product Requirements and Ideation Management,Quality Issue and Action Management,Supply Chain for Healthcare,Advanced Constraint Technology,Collaboration Messaging Framework,Supply Chain Management Common Components,Supply Chain Collaboration,Supply Chain Financial Orchestration Foundation,Backlog Management,Demand Management,Global Order Promising,Planning Central,Planning Collaboration,Planing Common,Production Scheduling,Replenishment Planning,Sales and Operations Planning,Supply Planning,Visual Information Builder,Visual Information Navigator
MDSFile={}
ReportOnly=Y
""".format(mds_relative_path)

        params_file = work_dir / "InputParams.txt"
        with open(params_file, 'w', encoding='utf-8', newline='\n') as f:
            f.write(input_params)

        if not params_file.exists() or params_file.stat().st_size < 50:
            raise RuntimeError(f"InputParams.txt creation failed (size: {params_file.stat().st_size} bytes)")
        log_queue.append(f"✅ InputParams.txt created ({params_file.stat().st_size} bytes)")

        # === COPY TOOL JAR ===
        tool_jar_src = (APP_DIR / TOOL_JAR).resolve()
        if not tool_jar_src.exists():
            raise FileNotFoundError(f"Tool JAR not found: {tool_jar_src}")
        tool_jar_dest = work_dir / TOOL_JAR
        shutil.copy2(tool_jar_src, tool_jar_dest)
        log_queue.append(f"✅ Tool JAR copied to workspace")

        # === JAVA EXECUTION ===
        java_exe = shutil.which("java.exe") or shutil.which("java")
        if not java_exe:
            raise RuntimeError("Java not found in PATH. Install Java 8/11 and add to PATH.")

        log_queue.append(f"⚙️ Using Java: {java_exe}")
        log_queue.append(f"⚙️ Executing from workspace: {work_dir}")
        log_queue.append("🚀 Running Redwood Helper Tool...")

        result = subprocess.run(
            [
                java_exe,
                "-cp", str(tool_jar_dest.absolute()),
                "oracle.apps.scm.cust.analyze.RedwoodCustHelper",
                str(work_dir.absolute())
            ],
            capture_output=True,
            text=True,
            cwd=str(work_dir.absolute()),
            timeout=300,
            encoding='utf-8',
            errors='replace'
        )

        if result.stdout.strip():
            log_queue.append("🔍 Java STDOUT (first 500 chars):")
            log_queue.append(result.stdout[:500])
        if result.stderr.strip():
            log_queue.append("⚠️ Java STDERR (first 800 chars):")
            log_queue.append(result.stderr[:800])

        # === READ JAVA TOOL LOG FILE ===
        log_files = list(work_dir.glob("CustMigration_*.log"))
        if log_files:
            log_file = log_files[0]
            log_queue.append(f"📋 Reading Java tool log: {log_file.name}")
            try:
                with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
                    log_content = f.read()
                    log_queue.append("📄 Java Tool Log (last 2000 chars):")
                    log_queue.append(log_content[-2000:])
            except Exception as e:
                log_queue.append(f"⚠️ Could not read log file: {e}")

        if result.returncode != 0:
            log_queue.append(f"⚠️ Java tool exit code: {result.returncode}")
            if tool_output_dir.exists() and any(tool_output_dir.iterdir()):
                log_queue.append("✅ Tool generated output despite non-zero exit code - proceeding")
            else:
                raise RuntimeError(
                    f"Java tool failed (exit {result.returncode}). "
                    f"Check the log file above for details."
                )
        else:
            log_queue.append("✅ Java tool executed successfully")

        # === VALIDATE OUTPUT DIRECTORY ===
        if not tool_output_dir.exists():
            log_queue.append("ℹ️ No 'output' directory created - likely no customizations found in the MDS JAR")
            summary_report = output_dir / "NoCustomizationsFound.txt"
            summary_report.parent.mkdir(parents=True, exist_ok=True)
            with open(summary_report, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("MIGRATION ANALYSIS REPORT\n")
                f.write("=" * 80 + "\n\n")
                f.write("STATUS: No customizations found\n\n")
                f.write("The uploaded migration JAR was successfully analyzed, but no customizations\n")
                f.write("were detected in the SCM modules.\n\n")
                f.write("=" * 80 + "\n")
            log_queue.append(f"📝 Created summary report: {summary_report.name}")
        else:
            output_files = [f for f in tool_output_dir.rglob("*") if f.is_file()]
            if not output_files:
                log_queue.append(f"ℹ️ 'output' directory exists but contains NO files")
                summary_report = output_dir / "NoCustomizationsFound.txt"
                summary_report.parent.mkdir(parents=True, exist_ok=True)
                with open(summary_report, 'w', encoding='utf-8') as f:
                    f.write("=" * 80 + "\n")
                    f.write("MIGRATION ANALYSIS REPORT\n")
                    f.write("=" * 80 + "\n\n")
                    f.write("STATUS: No customizations found\n\n")
                    f.write("The Java tool completed successfully but found no customizations.\n\n")
                    f.write("=" * 80 + "\n")
                log_queue.append(f"📝 Created summary report: {summary_report.name}")
            else:
                log_queue.append(f"✅ Java tool generated {len(output_files)} output files")
                log_queue.append("📤 Copying output files...")
                try:
                    shutil.copytree(tool_output_dir, output_dir, dirs_exist_ok=True)
                    log_queue.append(f"✅ Output files copied to: {output_dir}")
                except Exception as e:
                    log_queue.append(f"⚠️ shutil.copytree failed, using manual copy: {e}")
                    file_count = 0
                    for src_path in tool_output_dir.rglob("*"):
                        if src_path.is_file():
                            rel_path = src_path.relative_to(tool_output_dir)
                            dst_path = output_dir / rel_path
                            dst_path.parent.mkdir(parents=True, exist_ok=True)
                            shutil.copy2(src_path, dst_path)
                            file_count += 1
                    log_queue.append(f"✅ Manually copied {file_count} files")

        # === CREATE FINAL ZIP ===
        log_queue.append("🗜️ Creating output ZIP...")
        files_to_zip = [f for f in output_dir.rglob("*") if f.is_file()]
        if not files_to_zip:
            marker_file = output_dir / "ProcessingComplete.txt"
            with open(marker_file, 'w', encoding='utf-8') as f:
                f.write("Phase 2 processing completed successfully.\n")
                f.write("No output files were generated - check logs for details.\n")

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in output_dir.rglob("*"):
                if file_path.is_file():
                    arcname = file_path.relative_to(output_dir).as_posix()
                    zipf.write(file_path, arcname)

        log_queue.append(f"✅ Output ZIP created: {zip_path.name} ({zip_path.stat().st_size:,} bytes)")
        sessions[session_id]["status"] = "completed"
        sessions[session_id]["output_zip"] = str(zip_path)
        log_queue.append("🎉 PHASE 2 COMPLETED SUCCESSFULLY")

    except Exception as e:
        error_detail = f"{type(e).__name__}: {str(e)}"
        log_queue.append(f"💥 PHASE 2 FAILED: {error_detail}")
        log_queue.append(f"📋 Traceback:\n{traceback.format_exc()}")
        sessions[session_id]["status"] = "error"
        sessions[session_id]["error"] = error_detail

    finally:
        for dir_path in [work_dir, output_dir]:
            if dir_path.exists():
                try:
                    shutil.rmtree(dir_path, ignore_errors=True)
                except Exception as e:
                    log_queue.append(f"🧹 Cleanup warning for {dir_path.name}: {type(e).__name__}")
        if temp_jar_safe.exists():
            try:
                temp_jar_safe.unlink()
                log_queue.append(f"🧹 Cleaned up temp backup: {temp_jar_safe.name}")
            except Exception as e:
                log_queue.append(f"🧹 Cleanup warning for {temp_jar_safe.name}: {type(e).__name__}")
        _release_job(session_id)
        _schedule_session_cleanup(session_id)


def _phase1_run_dir_from_log_line(line: str) -> Optional[Path]:
    """Extract logs/run_* folder from combine.py output."""
    match = re.search(r"((?:[A-Za-z]:)?[/\\].*?[/\\]logs[/\\]run_\d{8}_\d{6})", line)
    if not match:
        match = re.search(r"(\blogs[/\\]run_\d{8}_\d{6})", line)
    if match:
        return Path(match.group(1)).resolve()
    return None


def _latest_phase1_run_dir(logs_root: Path, pre_runs: set[Path]) -> Optional[Path]:
    if not logs_root.exists():
        return None
    post_runs = {p.resolve() for p in logs_root.iterdir() if p.is_dir()}
    new_runs = sorted(post_runs - pre_runs, key=lambda p: p.stat().st_mtime)
    if new_runs:
        return new_runs[-1]
    all_runs = sorted(post_runs, key=lambda p: p.stat().st_mtime)
    return all_runs[-1] if all_runs else None


def _phase1_summary_ready(run_dir: Path) -> bool:
    summary_path = run_dir / "summary_log.txt"
    if not summary_path.exists() or not summary_path.is_file():
        return False
    try:
        summary_text = summary_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return False
    return "Completed At" in summary_text


def _mark_phase1_download_ready(session_id: str, run_dir: Optional[Path]):
    if not run_dir or not run_dir.exists() or not run_dir.is_dir():
        return
    session = sessions.get(session_id)
    if not session:
        return
    resolved_run_dir = str(run_dir.resolve())
    if session.get("phase1_run_dir") != resolved_run_dir:
        session["ppt_ready"] = False
        session["ppt_path"] = None
        session["output_zip"] = None
    session["phase1_run_dir"] = resolved_run_dir
    if _phase1_summary_ready(run_dir):
        session["summary_ready"] = True
        session["download_ready"] = True


def _refresh_phase1_download_ready(session_id: str):
    session = sessions.get(session_id)
    if not session or session.get("phase") != "phase1":
        return
    if session.get("download_ready"):
        return

    for line in reversed(job_logs.get(session_id, [])):
        detected_run_dir = _phase1_run_dir_from_log_line(str(line))
        if detected_run_dir:
            _mark_phase1_download_ready(session_id, detected_run_dir)
            if session.get("download_ready"):
                return

    if session.get("status") == "completed":
        _mark_phase1_download_ready(
            session_id,
            _latest_phase1_run_dir((APP_DIR / "logs").resolve(), set()),
        )


def _build_phase1_logs_zip(session_id: str, session: Dict) -> Path:
    run_dir_str = session.get("phase1_run_dir")
    run_dir = Path(run_dir_str) if run_dir_str else None
    if not run_dir or not run_dir.exists() or not run_dir.is_dir():
        raise HTTPException(status_code=404, detail="Phase 1 logs folder not found")

    log_files = sorted(p for p in run_dir.iterdir() if p.is_file() and p.suffix.lower() == ".txt")
    if not log_files:
        raise HTTPException(status_code=404, detail="No Phase 1 log files found")

    temp_zip = APP_DIR / f"phase1_logs_{session_id}.zip"
    if temp_zip.exists():
        temp_zip.unlink()

    with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for log_file in log_files:
            zf.write(log_file, arcname=f"logs/{log_file.name}")

    return temp_zip


def _build_phase1_output_zip(session_id: str, session: Dict) -> Path:
    run_dir_str = session.get("phase1_run_dir")
    run_dir = Path(run_dir_str) if run_dir_str else None
    if not run_dir or not run_dir.exists() or not run_dir.is_dir():
        raise HTTPException(status_code=404, detail="Phase 1 output folder not found")

    customer_safe = _sanitize_filename(session.get("customer", ""))
    root_name = customer_safe or run_dir.name
    temp_zip = APP_DIR / f"phase1_output_{session_id}.zip"
    if temp_zip.exists():
        temp_zip.unlink()

    files = sorted(p for p in run_dir.rglob("*") if p.is_file())
    if not files:
        raise HTTPException(status_code=404, detail="No Phase 1 output files found")

    with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in files:
            rel_path = file_path.relative_to(run_dir).as_posix()
            zf.write(file_path, arcname=f"{root_name}/{rel_path}")

    return temp_zip


def _generate_phase1_ppt(session_id: str, session: Dict) -> Path:
    run_dir_str = session.get("phase1_run_dir")
    run_dir = Path(run_dir_str) if run_dir_str else None
    if not run_dir or not run_dir.exists() or not run_dir.is_dir():
        raise HTTPException(status_code=404, detail="Phase 1 run folder not found")

    if not _phase1_summary_ready(run_dir):
        raise HTTPException(status_code=400, detail="Summary log is not ready yet")

    create_ppt_script = (APP_DIR / "create_ppt.py").resolve()
    if not create_ppt_script.exists():
        raise HTTPException(status_code=404, detail="create_ppt.py not found")

    customer_safe = _sanitize_filename(session.get("customer", ""))
    output_name = f"{customer_safe}.pptx" if customer_safe else f"{run_dir.name}_report.pptx"
    output_pptx = (run_dir / output_name).resolve()

    output_pptx.unlink(missing_ok=True)

    result = subprocess.run(
        [sys.executable, str(create_ppt_script), str(run_dir.resolve()), str(output_pptx)],
        capture_output=True,
        text=True,
        cwd=str(APP_DIR),
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "PPT generation failed").strip()
        if "Cannot find module 'pptxgenjs'" in detail:
            detail = "PPT dependency missing: run npm install pptxgenjs, then try Download PPT again."
        raise HTTPException(status_code=500, detail=detail[-2000:])
    if not output_pptx.exists() or output_pptx.stat().st_size == 0:
        raise HTTPException(status_code=500, detail="PPT file was not created")

    session["ppt_path"] = str(output_pptx)
    session["ppt_ready"] = True
    job_logs.setdefault(session_id, []).append(f"📊 Phase 1 PPT generated: {output_pptx.name}")
    return output_pptx


def run_phase1_background(session_id: str, payload: Phase1StartRequest):
    """Run combine.py as a subprocess and stream logs to session queue."""
    log_queue = job_logs.get(session_id, [])
    try:
        session = sessions[session_id]
        log_queue.append(f"[{datetime.now()}] 🚀 Starting Phase 1 automation")

        combine_script = (APP_DIR / "combine.py").resolve()
        if not combine_script.exists():
            raise FileNotFoundError("combine.py not found in project root")

        logs_root = (APP_DIR / "logs").resolve()
        pre_runs = set()
        if logs_root.exists():
            pre_runs = {p.resolve() for p in logs_root.iterdir() if p.is_dir()}

        env = os.environ.copy()
        env["FUSION_BASE_URL"] = payload.fusion_url
        env["FUSION_USERNAME"] = payload.username
        env["FUSION_PASSWORD"] = payload.password
        env["EXCEL_PATH"] = payload.input_excel_path
        env["RUN_ONLY"] = payload.run_only or "all"
        if payload.auto_only is not None:
            env["AUTO_ONLY"] = "true" if payload.auto_only else "false"

        cmd = [sys.executable, str(combine_script), "--run-only", env["RUN_ONLY"]]
        log_queue.append(f"⚙️ Executing: {' '.join(cmd)}")
        log_queue.append(f"🎯 Run only mode: {env['RUN_ONLY']}")

        process = subprocess.Popen(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(APP_DIR),
            env=env,
            bufsize=1,
        )

        if process.stdout:
            for line in process.stdout:
                line = line.rstrip("\n")
                if line:
                    log_queue.append(line)
                    detected_run_dir = _phase1_run_dir_from_log_line(line)
                    if detected_run_dir:
                        _mark_phase1_download_ready(session_id, detected_run_dir)

        return_code = process.wait()

        run_dir_str = session.get("phase1_run_dir")
        run_dir = Path(run_dir_str) if run_dir_str else _latest_phase1_run_dir(logs_root, pre_runs)
        _mark_phase1_download_ready(session_id, run_dir)

        output_zip = None
        if run_dir and run_dir.exists():
            output_zip = _build_phase1_output_zip(session_id, session).resolve()
            log_queue.append(f"📦 Output archive created: {output_zip.name}")

        if return_code != 0:
            session["status"] = "error"
            session["error"] = f"combine.py failed with exit code {return_code}"
            log_queue.append(f"💥 combine.py failed with exit code {return_code}")
            return

        session["status"] = "completed"
        session["summary_ready"] = bool(run_dir and run_dir.exists() and _phase1_summary_ready(run_dir))
        session["download_ready"] = bool(session.get("summary_ready"))
        if output_zip:
            session["output_zip"] = str(output_zip)
            session["phase1_run_dir"] = str(run_dir) if run_dir else None
        if session.get("download_ready"):
            log_queue.append("📦 Phase 1 output zip ready for download")
            if payload.run_only == "auto":
                try:
                    ppt_path = _generate_phase1_ppt(session_id, session)
                    log_queue.append(f"📊 Auto-only PPT ready: {ppt_path.name}")
                except Exception as ppt_exc:
                    session["ppt_ready"] = False
                    session["ppt_path"] = None
                    log_queue.append(f"⚠️ Auto-only PPT generation failed: {ppt_exc}")
        log_queue.append("✅ Phase 1 completed successfully")

    except Exception as e:
        sessions[session_id]["status"] = "error"
        sessions[session_id]["error"] = str(e)
        log_queue.append(f"💥 Phase 1 failed: {str(e)}")
        log_queue.append(f"📋 Traceback: {traceback.format_exc()}")
    finally:
        _release_job(session_id)
        _schedule_session_cleanup(session_id)


@app.post("/api/phase1/start")
async def start_phase1(
    request: Request,
    background_tasks: BackgroundTasks,
    fusion_url: str = Form(...),
    username: str = Form(...),
    password: str = Form(...),
    excel_file: UploadFile = File(...),
    run_only: Optional[str] = Form(None),
    auto_only: Optional[bool] = Form(None),
    session_id: str = Cookie(None)
):
    sid = session_id or request.cookies.get("session_id") or request.headers.get("X-Session-Id")
    if not sid or sid not in sessions:
        raise HTTPException(status_code=401, detail="Invalid session")

    existing = sessions[sid]
    user = existing.get("user")
    customer = existing.get("customer")

    original_filename = Path(excel_file.filename or "").name
    if not original_filename:
        raise HTTPException(status_code=400, detail="Excel file is required")

    file_ext = Path(original_filename).suffix.lower()
    if file_ext != ".xlsx":
        raise HTTPException(status_code=400, detail="Only .xlsx Excel files are allowed")
    if not _reserve_job(sid, "phase1"):
        raise HTTPException(
            status_code=409,
            detail="Server is busy with another automation. Please try again after it finishes.",
        )

    uploads_dir = APP_DIR / f"uploads_{sid}"
    shutil.rmtree(uploads_dir, ignore_errors=True)
    uploads_dir.mkdir(parents=True, exist_ok=True)
    saved_excel_path = uploads_dir / f"phase1_input{file_ext}"

    try:
        with saved_excel_path.open("wb") as buffer:
            shutil.copyfileobj(excel_file.file, buffer)
        if saved_excel_path.stat().st_size == 0:
            raise ValueError("Uploaded Excel file is empty")
        with saved_excel_path.open("rb") as buffer:
            if buffer.read(2) != b"PK":
                raise ValueError("Uploaded file is not a valid .xlsx workbook")
    except ValueError as exc:
        saved_excel_path.unlink(missing_ok=True)
        _release_job(sid)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        saved_excel_path.unlink(missing_ok=True)
        _release_job(sid)
        raise HTTPException(status_code=500, detail=f"Unable to save uploaded Excel file: {exc}") from exc
    finally:
        await excel_file.close()

    try:
        run_only_mode = _normalize_phase1_run_only(run_only)
    except Exception:
        _release_job(sid)
        _remove_path(uploads_dir)
        raise
    payload = Phase1StartRequest(
        fusion_url=fusion_url,
        username=username,
        password=password,
        input_excel_path=str(saved_excel_path.resolve()),
        run_only=run_only_mode,
        auto_only=auto_only,
    )

    sessions[sid] = {
        "phase": "phase1",
        "mode": "browser",
        "status": "processing",
        "created_at": datetime.now(),
        "user": user,
        "customer": customer,
        "fusion_url": fusion_url,
        "input_excel_path": payload.input_excel_path,
        "excel_filename": original_filename,
        "upload_dir": str(uploads_dir.resolve()),
        "summary_ready": False,
        "download_ready": False,
        "ppt_ready": False,
        "ppt_path": None,
        "run_only": payload.run_only,
        "auto_only": payload.auto_only,
    }
    job_logs[sid] = [
        f"[{datetime.now()}] ✅ Phase 1 session created",
        f"📄 Uploaded Excel saved: {original_filename}",
    ]

    background_tasks.add_task(run_phase1_background, sid, payload)

    return {
        "session_id": sid,
        "status": "processing",
        "message": "Phase 1 automation started",
        "user": user,
        "customer": customer,
    }


# ==================== PHASE 3 - DUAL MODE IMPLEMENTATION ====================

if PHASE3_AVAILABLE:

    # ── ZIP Mode helpers ──────────────────────────────────────────────────────

    def parse_personalization(p_str: str) -> Optional[Tuple[str, bool]]:
        s = str(p_str).strip()
        if s in ("Element re-ordered(mds:move)", ""):
            return None
        if "=" not in s:
            return None
        k, v = s.split("=", 1)
        k, v = k.strip().lower(), v.strip().lower()
        if k == "rendered":
            return ("hidden", v == "false")
        elif k == "readonly":
            return ("readonly", v == "true")
        elif k == "required":
            return ("required", v == "true")
        elif k == "showrequired":
            return ("required", v == "true")
        return None

    def find_metadata_json_in_zip(extracted_root: Path, page_id: str, log_queue: List[str]) -> Optional[Path]:
        config = get_page_config(page_id)
        if not config:
            return None
        vb_page_file = config.vb_page_file
        json_filename = "metadata-rules-x.json"
        log_queue.append(f"🔍 Searching for: {vb_page_file}/{json_filename}")
        search_patterns = [
            f"extension1/sources/dynamicLayouts/oracle_prc_procurementUI/{vb_page_file}/{json_filename}",
            f"extension1/sources/dynamicLayouts/**/{vb_page_file}/{json_filename}",
            f"webApps/**/pages/{vb_page_file}/{json_filename}",
            f"**/pages/{vb_page_file}/{json_filename}",
            f"**/{vb_page_file}/{json_filename}",
        ]
        for pattern in search_patterns:
            if "**" not in pattern:
                candidate = extracted_root / pattern
                if candidate.is_file():
                    log_queue.append(f"  ✓ Found: {candidate.relative_to(extracted_root)}")
                    return candidate
            else:
                matches = list(extracted_root.glob(pattern))
                if matches:
                    log_queue.append(f"  ✓ Found: {matches[0].relative_to(extracted_root)}")
                    return matches[0]
        log_queue.append(f"  ✗ Not found")
        return None

    def process_excel_for_zip(
        excel_path: Path,
        composite_mapping: Dict[Tuple[str, str], str],
        log_queue: List[str]
    ) -> Tuple[Dict, List[Tuple], List[Tuple]]:
        log_queue.append(f"📊 Processing Excel: {excel_path.name}")
        df = pd.read_excel(excel_path, dtype=str)
        required_cols = {"File", "Field", "Display Name", "Personalization"}
        if not required_cols <= set(df.columns):
            raise ValueError(f"Missing columns: {required_cols - set(df.columns)}")

        fields = {}
        mapped_records = []
        unmapped_records = []
        skip_display_names = {
            "Column", "Panel Form Layout", "Region", "Delivery", "Billing", "Tax",
            "NotesandAttachments", "Source", "Popup", "Output Text",
            "Panel Group Layout", "DocumentAttachments panelHeader"
        }

        for _, row in df.iterrows():
            file = str(row["File"]).strip()
            field = str(row["Field"]).strip()
            disp = str(row["Display Name"]).strip()
            pers = row["Personalization"]
            if file.lower() == "file":
                continue
            key = (file, field)
            new_key = composite_mapping.get(key)
            if new_key is None:
                if disp not in skip_display_names:
                    unmapped_records.append((file, field, disp, str(pers)))
                continue
            prop = parse_personalization(pers)
            if not prop:
                continue
            p_key, p_val = prop
            if new_key not in fields:
                fields[new_key] = {}
            fields[new_key][p_key] = {"value": p_val}
            mapped_records.append((file, field, disp, new_key))

        log_queue.append(f"  ✓ Mapped: {len(mapped_records)}, Unmapped: {len(unmapped_records)}")
        return fields, mapped_records, unmapped_records

    def inject_fields_into_json(json_path: Path, fields: Dict, log_queue: List[str]) -> bool:
        log_queue.append(f"💉 Updating: {json_path.name}")
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if "addMetadataRules" not in data:
                log_queue.append("  ✗ Missing 'addMetadataRules'"); return False
            if not isinstance(data["addMetadataRules"], list) or len(data["addMetadataRules"]) == 0:
                log_queue.append("  ✗ 'addMetadataRules' not a non-empty list"); return False
            if "overlay" not in data["addMetadataRules"][0]:
                log_queue.append("  ✗ Missing 'overlay'"); return False
            data["addMetadataRules"][0]["overlay"]["fields"] = fields
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            log_queue.append(f"  ✓ Injected {len(fields)} fields")
            return True
        except Exception as e:
            log_queue.append(f"  ✗ Error: {str(e)}")
            return False

    def process_zip_file(
        session_id: str,
        page_id: str,
        excel_path: Path,
        zip_in_path: Path,
        log_queue: List[str]
    ) -> Optional[Tuple[Path, Path]]:
        log_queue.append(f"[{datetime.now()}] 🚀 Starting ZIP Mode processing")
        config = get_page_config(page_id)
        composite_mapping = get_composite_mapping(page_id)
        if not config or not composite_mapping:
            log_queue.append(f"❌ Configuration not found for '{page_id}'"); return None
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)
                log_queue.append("📤 Extracting ZIP...")
                with zipfile.ZipFile(zip_in_path, 'r') as zf:
                    zf.extractall(tmp_path)
                json_path = find_metadata_json_in_zip(tmp_path, page_id, log_queue)
                if not json_path:
                    log_queue.append(f"❌ metadata-rules-x.json not found for '{config.vb_page_file}'")
                    return None
                fields, mapped_records, unmapped_records = process_excel_for_zip(
                    excel_path, composite_mapping, log_queue
                )
                report_path = generate_field_mapping_report(
                    session_id=session_id,
                    excel_file=str(excel_path),
                    page_name=config.page_name,
                    mapped_records=mapped_records,
                    unmapped_records=unmapped_records,
                    log_queue=log_queue
                )
                if not inject_fields_into_json(json_path, fields, log_queue):
                    return None
                log_queue.append("📦 Repackaging ZIP...")
                output_zip = Path(f"output_{page_id}_{session_id}.zip")
                with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for root, dirs, files in os.walk(tmp_path):
                        for file in files:
                            full_path = Path(root) / file
                            zf_out.write(full_path, full_path.relative_to(tmp_path))
                log_queue.append(f"✅ Output ZIP: {output_zip.name}")
                log_queue.append(f"✅ Mapping Report: {report_path.name}")
                return (output_zip, report_path)
        except Exception as e:
            log_queue.append(f"💥 Error: {str(e)}")
            log_queue.append(f"📋 Traceback: {traceback.format_exc()}")
            return None

    # ── Phase 3 endpoints ─────────────────────────────────────────────────────

    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.get("/api/phase3/pages")
    async def list_phase3_pages():
        return {
            "available_pages": list_available_pages(),
            "total_pages": len(list_available_pages()),
            "modes": {
                #NON SSo "browser": "Full automation with browser (requires credentials)",
                "zip": "File processing only (for SSO/MFA restricted accounts)"
            }
        }

    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.post("/api/phase3/upload-files/{page_id}")
    async def upload_phase3_files(
        page_id: str,
        excel_file: UploadFile = File(...),
        zip_file: UploadFile = File(None)
    ):
        if not validate_page_exists(page_id):
            raise HTTPException(status_code=404, detail=f"Page '{page_id}' not found")
        config = get_page_config(page_id)
        if not excel_file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")
        excel_path = Path(f"phase3_{page_id}_{excel_file.filename}")
        with open(excel_path, "wb") as f:
            content = await excel_file.read()
            f.write(content)
        response = {
            "status": "success",
            "page_id": page_id,
            "page_name": config.page_name,
            "excel_path": str(excel_path),
            "files": {"excel": excel_file.filename}
        }
        if zip_file:
            if not zip_file.filename.endswith('.zip'):
                raise HTTPException(status_code=400, detail="ZIP file must be .zip")
            zip_path = Path(f"phase3_{page_id}_{zip_file.filename}")
            with open(zip_path, "wb") as f:
                content = await zip_file.read()
                f.write(content)
            response["zip_path"] = str(zip_path)
            response["files"]["zip"] = zip_file.filename
        return response

    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.post("/api/phase3/start", response_model=Phase3Response)
    async def start_phase3(request: dict, background_tasks: BackgroundTasks,session_id: str = Cookie(None)):
        mode = request.get("mode", "browser")
        page_id = request.get("page_id")
        if not validate_page_exists(page_id):
            raise HTTPException(status_code=404, detail=f"Page '{page_id}' not found")
        config = get_page_config(page_id)
        # session_id = str(uuid.uuid4())

        # if mode == "browser":
        #     browser_req = Phase3BrowserRequest(**request)
        #     sessions[session_id] = {
        #         "phase": "phase3", "mode": "browser", "status": "running",
        #         "created_at": datetime.now(), "url": browser_req.url,
        #         "username": browser_req.username, "project_name": browser_req.project_name,
        #         "page_id": page_id, "page_name": config.page_name,
        #         "requires_mfa": browser_req.requires_mfa
        #     }
        #     job_logs[session_id] = []
        #     job_logs[session_id].append(f"[{datetime.now()}] ✅ Phase 3 session created (BROWSER MODE)")
        #     job_logs[session_id].append(f"📄 Page: {config.page_name}")
        #     background_tasks.add_task(
        #         run_phase3_browser_mode,
        #         session_id, browser_req.url, browser_req.username, browser_req.password,
        #         browser_req.project_name, page_id,
        #         browser_req.excel_file_path or config.excel_template,
        #         browser_req.requires_mfa, browser_req.mfa_code
        #     )
        #     return Phase3Response(
        #         session_id=session_id, status="started", mode="browser",
        #         message=f"Phase 3 (Browser Mode) started for '{config.page_name}'"
        #     )

        if mode == "zip":
            zip_req = Phase3ZipRequest(**request)
            excel_path = Path(f"phase3_{page_id}_{zip_req.excel_filename}")
            zip_path   = Path(f"phase3_{page_id}_{zip_req.zip_filename}")
            if not excel_path.exists():
                raise HTTPException(status_code=400, detail="Excel file not found. Upload first.")
            if not zip_path.exists():
                raise HTTPException(status_code=400, detail="ZIP file not found. Upload first.")
            session = sessions[session_id]

            user = session["user"]
            customer = session["customer"]
            print(user)
            print(customer)
            sessions[session_id] = {
                "phase": "phase3", "mode": "zip", "status": "processing",
                "created_at": datetime.now(), "page_id": page_id,
                "page_name": config.page_name,
                "excel_file": str(excel_path), "zip_file": str(zip_path),
                "user":user,
                "customer":customer
                
            }
            job_logs[session_id] = []
            job_logs[session_id].append(f"[{datetime.now()}] ✅ Phase 3 session created (ZIP MODE)")
            job_logs[session_id].append(f"📄 Page: {config.page_name}")
            background_tasks.add_task(run_phase3_zip_mode, session_id, page_id, excel_path, zip_path)
            print("RESPONSE:", {
                "user": user,
                "customer": customer
            })
            return Phase3Response(
                session_id=session_id, status="processing", mode="zip",
                message=f"Phase 3 (ZIP Mode) started for '{config.page_name}'",
                customer=customer,
                user=user,
            )

        else:
            raise HTTPException(status_code=400, detail=f"Invalid mode: {mode}")

    # ── Browser mode background task ──────────────────────────────────────────

    # def run_phase3_browser_mode(
    #     session_id: str, url: str, username: str, password: str,
    #     project_name: str, page_id: str, excel_file: str,
    #     requires_mfa: bool, mfa_code: Optional[str]
    # ):
    #     log_queue = job_logs[session_id]
    #     try:
    #         config = get_page_config(page_id)
    #         log_queue.append(f"🚀 Starting Browser Mode automation")
    #         parsed_url = url.lower()
    #         bypass_domains = PROXY_BYPASS.split(',')
    #         should_bypass = any(domain.strip() in parsed_url for domain in bypass_domains)

    #         with sync_playwright() as p:
    #             launch_options = {
    #                 "headless": True,
    #                 "args": ['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
    #             }
    #             if not should_bypass and PROXY_SERVER:
    #                 launch_options["proxy"] = {"server": PROXY_SERVER}
    #                 log_queue.append(f"📡 Using proxy: {PROXY_SERVER}")
    #             else:
    #                 log_queue.append(f"🔓 Bypassing proxy for Oracle domain")

    #             browser = p.chromium.launch(**launch_options)
    #             context = browser.new_context(ignore_https_errors=True)
    #             page = context.new_page()
    #             log_queue.append(f"🌐 Navigating to {url}")
    #             page.goto(url, timeout=60000)

    #             if requires_mfa:
    #                 handle_mfa_login(page, username, password, log_queue, mfa_code)
    #             else:
    #                 handle_sso_login(page, username, password, log_queue)

    #             navigate_to_page(page, config.navigation_path, log_queue)
    #             vb_page = open_visual_builder(page, log_queue)
    #             select_vb_project(vb_page, project_name, log_queue)
    #             navigate_to_vb_page(vb_page, config.vb_search_term, log_queue)

    #             rule_name = f"auto_{page_id}_{session_id[:8]}"
    #             create_customization_rule(vb_page, rule_name, log_queue)
    #             activate_advanced_mode(vb_page, log_queue)
    #             open_metadata_file(vb_page, config.vb_page_file, log_queue)

    #             fields, mapped_records, unmapped_records = process_excel_mapping(
    #                 excel_file, config.composite_mapping, log_queue
    #             )
    #             report_path = generate_field_mapping_report(
    #                 session_id, excel_file, config.page_name,
    #                 mapped_records, unmapped_records, log_queue
    #             )
    #             sessions[session_id]["mapping_report"] = str(report_path)

    #             success = inject_personalization_to_monaco(vb_page, fields, log_queue)
    #             if success:
    #                 log_queue.append("✅ Personalization injected!")
    #                 log_queue.append("📸 Capturing URLs...")
    #                 urls = capture_vb_urls(vb_page, log_queue)
    #                 if urls.get('editor_url'):
    #                     sessions[session_id]["editor_url"] = urls['editor_url']
    #                 if urls.get('preview_url'):
    #                     sessions[session_id]["preview_url"] = urls['preview_url']
    #                 log_queue.append("✅ Automation completed successfully")
    #                 sessions[session_id]["status"] = "completed"
    #             else:
    #                 raise Exception("Injection failed")

    #             browser.close()

    #     except Exception as e:
    #         log_queue.append(f"💥 Error: {str(e)}")
    #         sessions[session_id]["status"] = "error"
    #         sessions[session_id]["error"] = str(e)
    #         log_queue.append(f"📋 Traceback: {traceback.format_exc()}")

    # # ── ZIP mode background task ──────────────────────────────────────────────

    def run_phase3_zip_mode(session_id: str, page_id: str, excel_path: Path, zip_path: Path):
        log_queue = job_logs[session_id]
        try:
            result = process_zip_file(session_id, page_id, excel_path, zip_path, log_queue)
            if result:
                output_zip, report_path = result
                sessions[session_id]["status"] = "completed"
                sessions[session_id]["output_zip"] = str(output_zip)
                sessions[session_id]["mapping_report"] = str(report_path)
                log_queue.append("✅ ZIP Mode completed!")
                log_queue.append(f"📦 Output ZIP available for download")
                log_queue.append(f"📊 Mapping Report available for download")
            else:
                sessions[session_id]["status"] = "error"
                sessions[session_id]["error"] = "Processing failed"
        except Exception as e:
            log_queue.append(f"💥 Error: {str(e)}")
            sessions[session_id]["status"] = "error"
            sessions[session_id]["error"] = str(e)
            log_queue.append(f"📋 Traceback: {traceback.format_exc()}")

    # ── Phase 3 specific endpoints ────────────────────────────────────────────

    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.get("/api/phase3/session/{session_id}/urls")
    async def get_phase3_urls(session_id: str):
        if session_id not in sessions:
            raise HTTPException(status_code=404, detail="Session not found")
        session = sessions[session_id]
        if session.get("phase") != "phase3" or session.get("mode") != "browser":
            raise HTTPException(status_code=400, detail="URLs only available for Browser mode")
        return {
            "session_id": session_id,
            "page_id": session.get("page_id"),
            "page_name": session.get("page_name"),
            "editor_url": session.get("editor_url"),
            "preview_url": session.get("preview_url"),
            "status": session.get("status")
        }

    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.get("/api/session/{session_id}/download-mapping-report")
    async def download_mapping_report_phase3(session_id: str):
        if session_id not in sessions:
            raise HTTPException(status_code=404, detail="Session not found")
        session = sessions[session_id]
        if session.get("phase") != "phase3":
            raise HTTPException(status_code=400, detail="Mapping report only available for Phase 3")
        report_path = session.get("mapping_report")
        if not report_path or not Path(report_path).exists():
            raise HTTPException(status_code=404, detail="Mapping report not found")
        return FileResponse(
            path=report_path,
            filename=Path(report_path).name,
            media_type="text/plain"
        )

else:
    # Client-specific temporary removal: Phase 3 customizations API.
    # Uncomment this decorator to restore the endpoint.
    # @app.get("/api/phase3/pages")
    async def phase3_unavailable():
        raise HTTPException(
            status_code=503,
            detail="Phase 3 modules not available. Ensure page_configurations.py and automation_engine.py exist."
        )


# ==================== COMMON ENDPOINTS ====================

@app.get("/api/session/{session_id}/logs")
async def get_session_logs(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    _refresh_phase1_download_ready(session_id)
    return {
        "session_id": session_id,
        "status": sessions[session_id]["status"],
        "logs": job_logs.get(session_id, []),
        "phase": sessions[session_id].get("phase"),
        "mode": sessions[session_id].get("mode"),
        "summary_ready": bool(sessions[session_id].get("summary_ready")),
        "download_ready": bool(sessions[session_id].get("download_ready")),
        "ppt_ready": bool(sessions[session_id].get("ppt_ready")),
    }

@app.get("/api/session/{session_id}/status")
async def get_session_status(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    _refresh_phase1_download_ready(session_id)
    return sessions[session_id].copy()

@app.get("/api/session/{session_id}/download")
async def download_results(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    session = sessions[session_id]
    phase = session.get("phase")
    mode  = session.get("mode")

    # Phase 2
    if phase == 2:
        if session["status"] != "completed":
            raise HTTPException(status_code=400, detail="Processing not completed")

        zip_path = APP_DIR / f"output_{session_id}.zip"
        if not zip_path.exists():
            raise HTTPException(status_code=404, detail="Output file not found")

        customer_safe = _sanitize_filename(session.get("customer", ""))
        final_name = f"{customer_safe}.zip" if customer_safe else "output.zip"
        download_zip = _zip_with_customer_root(zip_path, session_id, session, "phase2_output")

        response = FileResponse(
            path=download_zip,
            filename=final_name,
            media_type="application/zip"
        )
        response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
        return response

    # Phase 3
    elif phase == "phase3":
        if session["status"] != "completed":
            raise HTTPException(status_code=400, detail="Processing not completed")
        # if mode == "browser":
        #     report_path = session.get("mapping_report")
        #     if report_path and Path(report_path).exists():
        #         return FileResponse(path=report_path, filename=Path(report_path).name, media_type="text/plain")
        #     raise HTTPException(status_code=404, detail="Mapping report not found")
        if mode == "zip":
            output_zip = session.get("output_zip")
            if output_zip and Path(output_zip).exists():
                customer_safe = _sanitize_filename(session.get("customer", ""))
                final_name = f"{customer_safe}_phase3_output.zip" if customer_safe else Path(output_zip).name
                download_zip = _zip_with_customer_root(Path(output_zip), session_id, session, "phase3_output")
                response = FileResponse(path=download_zip, filename=final_name, media_type="application/zip")
                response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
                return response
            raise HTTPException(status_code=404, detail="Output ZIP not found")

    # Phase 1
    elif phase == "phase1":
        _refresh_phase1_download_ready(session_id)
        if not session.get("download_ready"):
            raise HTTPException(status_code=400, detail="Summary log is not ready yet")

        output_zip = session.get("output_zip")
        if output_zip and Path(output_zip).exists():
            customer_safe = _sanitize_filename(session.get("customer", ""))
            final_name = f"{customer_safe}.zip" if customer_safe else Path(output_zip).name
            download_zip = _zip_with_customer_root(Path(output_zip), session_id, session, "phase1_output_download")
            response = FileResponse(path=download_zip, filename=final_name, media_type="application/zip")
            response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
            return response

        try:
            temp_zip = _build_phase1_output_zip(session_id, session)
            customer_safe = _sanitize_filename(session.get("customer", ""))
            final_name = f"{customer_safe}.zip" if customer_safe else "phase1_output.zip"
            download_zip = _zip_with_customer_root(temp_zip, session_id, session, "phase1_output_download")
            response = FileResponse(path=download_zip, filename=final_name, media_type="application/zip")
            response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
            return response
        except HTTPException:
            temp_zip = _build_phase1_logs_zip(session_id, session)
            customer_safe = _sanitize_filename(session.get("customer", ""))
            final_name = f"{customer_safe}.zip" if customer_safe else "phase1_logs.zip"
            download_zip = _zip_with_customer_root(temp_zip, session_id, session, "phase1_logs_download")
            response = FileResponse(path=download_zip, filename=final_name, media_type="application/zip")
            response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
            return response
        raise HTTPException(status_code=404, detail="Phase 1 output not found")

    raise HTTPException(status_code=404, detail="No downloadable file found")


@app.get("/api/session/{session_id}/download-ppt")
async def download_phase1_ppt(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    if session.get("phase") != "phase1":
        raise HTTPException(status_code=400, detail="PPT download is available only for Phase 1")

    _refresh_phase1_download_ready(session_id)
    if not session.get("download_ready"):
        raise HTTPException(status_code=400, detail="Summary log is not ready yet")

    output_pptx = _generate_phase1_ppt(session_id, session)
    customer_safe = _sanitize_filename(session.get("customer", ""))
    final_name = f"{customer_safe}.pptx" if customer_safe else output_pptx.name

    response = FileResponse(
        path=output_pptx,
        filename=final_name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    response.headers["Content-Disposition"] = f'attachment; filename="{final_name}"'
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

@app.delete("/api/session/{session_id}")
async def delete_session(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    current_job = _active_job_snapshot()
    if current_job and current_job.get("session_id") == session_id:
        raise HTTPException(status_code=409, detail="Cannot delete a session while its automation is running")
    page_id = sessions[session_id].get("page_id")
    if page_id:
        for pattern in [f"phase3_{page_id}_*", f"output_{page_id}_{session_id}*", f"*_mapping_report_{session_id}*"]:
            for file in glob.glob(pattern):
                Path(file).unlink(missing_ok=True)

    _cleanup_session_artifacts(session_id)

    return {"status": "success", "message": f"Session {session_id} deleted"}

@app.get("/api/sessions")
async def list_sessions():
    return {
        "sessions": [
            {
                "session_id": sid,
                "phase": data.get("phase"),
                "mode": data.get("mode"),
                "status": data.get("status"),
                "created_at": data.get("created_at").isoformat() if data.get("created_at") else None,
                "log_count": len(job_logs.get(sid, []))
            }
            for sid, data in sessions.items()
        ],
        "total_sessions": len(sessions)
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
