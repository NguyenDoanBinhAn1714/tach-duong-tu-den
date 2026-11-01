"""Bản server của công cụ tách Đường / Từ / Đến.

Chạy:  python run_server.py   rồi mở http://127.0.0.1:8000
Hoặc gọi trực tiếp qua API (xem /docs để thử).
"""

import base64
import binascii
import hashlib
import html
import json
import os
import secrets
import shutil
import tempfile
import threading
import time
import sys
from urllib.parse import quote
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.middleware.sessions import SessionMiddleware

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:  # .env không bắt buộc nếu đã khai báo biến môi trường
    pass

from core.ai_fallback import GeminiFallback
from core.generic import split_column_in_workbook
from core.processor import export_to_excel, process_file
from core.splitter import split_row
from server.jobs import STORE, TTL_SECONDS, JobError, inspect_workbook, start_job

try:
    from config import GEMINI_API_KEY as _CONFIG_KEY
except ImportError:
    _CONFIG_KEY = ""

STATIC_DIR = Path(__file__).parent / "static"
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "20"))
API_KEY = os.getenv("GEMINI_API_KEY") or _CONFIG_KEY
APP_ENV = os.getenv("APP_ENV", "development").strip().lower()
APP_USERNAME = os.getenv("APP_USERNAME", "admin").strip() or "admin"
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
SESSION_SECRET = os.getenv("SESSION_SECRET", "")

# Không cho ứng dụng chạy production với mật khẩu/khóa phiên chưa được cấu hình.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

if APP_ENV == "production":
    if len(APP_PASSWORD) < 12:
        raise RuntimeError("Production cần APP_PASSWORD dài tối thiểu 12 ký tự.")
    if len(SESSION_SECRET) < 32:
        raise RuntimeError("Production cần SESSION_SECRET dài tối thiểu 32 ký tự.")

# Chạy local lần đầu: sinh mật khẩu tạm và in ra terminal để người dùng đăng nhập.
_GENERATED_PASSWORD = not bool(APP_PASSWORD)
if _GENERATED_PASSWORD:
    APP_PASSWORD = secrets.token_urlsafe(12)
    print("\n[TACH-DUONG] Chưa đặt APP_PASSWORD. Tài khoản đăng nhập local:")
    print(f"  Tên đăng nhập: {APP_USERNAME}")
    print(f"  Mật khẩu tạm: {APP_PASSWORD}")
    print("  Hãy đặt APP_PASSWORD trong tệp .env trước khi chia sẻ hoặc triển khai công khai.\n")
if not SESSION_SECRET:
    SESSION_SECRET = secrets.token_urlsafe(48)

LOGIN_MAX_ATTEMPTS = int(os.getenv("LOGIN_MAX_ATTEMPTS", "5"))
LOGIN_LOCK_SECONDS = int(os.getenv("LOGIN_LOCK_SECONDS", "300"))
REGISTRATION_ENABLED = os.getenv("REGISTRATION_ENABLED", "true").strip().lower() in {"1", "true", "yes", "on"}
USER_STORE_PATH = Path(os.getenv("USER_STORE_PATH", "data/users.json"))
_login_failures: dict[str, list[float]] = {}
_user_store_lock = threading.Lock()


def _read_users() -> dict:
    if not USER_STORE_PATH.exists():
        return {}
    try:
        value = json.loads(USER_STORE_PATH.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def _write_users(users: dict) -> None:
    USER_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = USER_STORE_PATH.with_suffix(".tmp")
    temporary.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(USER_STORE_PATH)


def _hash_password(password: str, salt: bytes | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return (
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(digest).decode("ascii"),
    )


def _password_matches(password: str, stored: dict) -> bool:
    try:
        salt = base64.b64decode(stored["salt"])
        expected = base64.b64decode(stored["hash"])
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
        return secrets.compare_digest(actual, expected)
    except (KeyError, ValueError, TypeError, binascii.Error):
        return False


def _username_exists(username: str) -> bool:
    if secrets.compare_digest(username, APP_USERNAME):
        return True
    return username in _read_users()


COLUMN_ORDERS = {
    "duong-tu-den": ("Đường", "Từ", "Đến"),
    "tu-den-duong": ("Từ", "Đến", "Đường"),
}
MODES = ("add-columns", "price-table")

app = FastAPI(
    title="Tách Đường / Từ / Đến",
    version="1.1.0",
    docs_url=None if APP_ENV == "production" else "/docs",
    redoc_url=None if APP_ENV == "production" else "/redoc",
    openapi_url=None if APP_ENV == "production" else "/openapi.json",
)

_ai = None


def get_ai():
    """Tạo bộ gọi Gemini một lần, dùng lại cho các request sau. None nếu chưa có key."""
    global _ai
    if not API_KEY:
        return None
    if _ai is None:
        _ai = GeminiFallback(API_KEY)
    return _ai if _ai.available() else None


class SplitTextRequest(BaseModel):
    text: str = Field(..., min_length=1, max_length=2000)
    last_area_name: Optional[str] = None
    use_ai: bool = False


@app.get("/", include_in_schema=False)
def index():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return FileResponse(STATIC_DIR / "favicon.svg", media_type="image/svg+xml")


def _safe_next(value: str | None) -> str:
    """Chỉ cho phép điều hướng nội bộ, tránh open-redirect."""
    if value and value.startswith("/") and not value.startswith("//") and "\\" not in value:
        return value
    return "/"


def _login_page(error: str = "", next_path: str = "/") -> HTMLResponse:
    error_html = (f'<div class="error" role="alert">{html.escape(error)}</div>' if error else "")
    next_html = html.escape(_safe_next(next_path), quote=True)
    body = f'''<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Đăng nhập · Tách Đường / Từ / Đến</title>
<style>
:root{{--ink:#10233f;--blue:#1769aa;--muted:#61738a;--line:#dce5ef;--bg:#f1f7e8}}
*{{box-sizing:border-box}}body{{margin:0;min-height:100vh;display:grid;place-items:center;padding:22px;background:radial-gradient(ellipse at 10% 0%,#dceeff,transparent 40%),var(--bg);font:16px/1.5 'Segoe UI',Arial,sans-serif;color:var(--ink)}}
.card{{width:min(100%,420px);background:#fff;border:1px solid var(--line);border-radius:22px;padding:32px;box-shadow:0 18px 50px #10233f17}}.brand{{display:flex;gap:12px;align-items:center;margin-bottom:24px}}.logo{{width:44px;height:44px;border-radius:13px;background:linear-gradient(145deg,#79b983,#477451);display:grid;place-items:center;color:#fff;font-size:22px;font-weight:800}}
h1{{margin:0;font-size:25px;letter-spacing:-.03em}}.sub{{color:var(--muted);font-size:14px;margin-top:3px}}label{{display:block;margin:17px 0 7px;font-size:14px;font-weight:700}}input{{width:100%;height:46px;padding:10px 12px;border:1px solid #bdcad9;border-radius:10px;font:inherit}}input:focus{{outline:3px solid #dceeff;border-color:var(--blue)}}button{{width:100%;margin-top:22px;height:47px;border:0;border-radius:10px;background:var(--blue);color:#fff;font:700 16px 'Segoe UI',Arial,sans-serif;cursor:pointer}}button:hover{{background:#12548a}}.error{{padding:10px 12px;border-radius:9px;background:#fff0f0;color:#a92f36;margin:14px 0 0;font-size:14px}}.foot{{margin-top:22px;color:var(--muted);font-size:12px;text-align:center}}
</style></head><body><main class="card"><div class="brand"><div class="logo" aria-hidden="true">D</div><div><h1>Tách Đường · Từ · Đến</h1><div class="sub">Đăng nhập để sử dụng công cụ xử lý Excel</div></div></div>
{error_html}<form method="post" action="/login" autocomplete="on"><input type="hidden" name="next" value="{next_html}">
<label for="username">Tên đăng nhập</label><input id="username" name="username" type="text" autocomplete="username" maxlength="100" required autofocus>
<label for="password">Mật khẩu</label><input id="password" name="password" type="password" autocomplete="current-password" maxlength="256" required>
<button id="loginSubmit" type="submit">Đăng nhập</button></form><div class="foot">Chưa có tài khoản? <a href="/register">Đăng ký ngay</a><br>Phiên đăng nhập được bảo vệ bằng cookie có chữ ký.</div></main></body></html>'''
    return HTMLResponse(body)


def _locked_out(client_ip: str) -> bool:
    now = time.monotonic()
    recent = [t for t in _login_failures.get(client_ip, []) if now - t < LOGIN_LOCK_SECONDS]
    _login_failures[client_ip] = recent
    return len(recent) >= LOGIN_MAX_ATTEMPTS


@app.get("/login", response_class=HTMLResponse, include_in_schema=False)
def login_page(request: Request, next: str = "/"):
    if _username_exists(request.session.get("username", "")):
        return RedirectResponse(_safe_next(next), status_code=303)
    return _login_page(next_path=next)


@app.post("/login", response_class=HTMLResponse, include_in_schema=False)
def login(request: Request, username: str = Form(..., max_length=100),
          password: str = Form(..., max_length=256), next: str = Form("/")):
    ip = request.client.host if request.client else "unknown"
    if _locked_out(ip):
        return HTMLResponse(
            _login_page("Bạn đã nhập sai quá nhiều lần. Hãy thử lại sau 5 phút.", next).body,
            status_code=429,
        )

    username_ok = secrets.compare_digest(username, APP_USERNAME)
    password_ok = secrets.compare_digest(password, APP_PASSWORD)
    users = _read_users()
    user_record = users.get(username)
    if not (username_ok and password_ok) and not (user_record and _password_matches(password, user_record)):
        _login_failures.setdefault(ip, []).append(time.monotonic())
        return HTMLResponse(_login_page("Tên đăng nhập hoặc mật khẩu không chính xác.", next).body, status_code=401)

    _login_failures.pop(ip, None)
    request.session.clear()
    request.session["username"] = username if user_record else APP_USERNAME
    request.session["authenticated_at"] = int(time.time())
    return RedirectResponse(_safe_next(next), status_code=303)


def _register_page(error: str = "") -> HTMLResponse:
    error_html = (f'<div class="error" role="alert">{html.escape(error)}</div>' if error else "")
    return HTMLResponse(f'''<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Đăng ký · Tách Đường / Từ / Đến</title><style>
:root{{--ink:#10233f;--blue:#1769aa;--muted:#61738a;--line:#dce5ef;--bg:#f1f7e8}}*{{box-sizing:border-box}}body{{margin:0;min-height:100vh;display:grid;place-items:center;padding:22px;background:radial-gradient(ellipse at 10% 0%,#dceeff,transparent 40%),var(--bg);font:16px/1.5 'Segoe UI',Arial,sans-serif;color:var(--ink)}}.card{{width:min(100%,420px);background:#fff;border:1px solid var(--line);border-radius:22px;padding:32px;box-shadow:0 18px 50px #10233f17}}h1{{margin:0 0 20px;font-size:25px}}label{{display:block;margin:15px 0 7px;font-size:14px;font-weight:700}}input{{width:100%;height:46px;padding:10px 12px;border:1px solid #bdcad9;border-radius:10px;font:inherit}}input:focus{{outline:3px solid #dceeff;border-color:var(--blue)}}button{{width:100%;margin-top:22px;height:47px;border:0;border-radius:10px;background:var(--blue);color:#fff;font:700 16px 'Segoe UI',Arial,sans-serif;cursor:pointer}}.error{{padding:10px 12px;border-radius:9px;background:#fff0f0;color:#a92f36;margin-bottom:12px;font-size:14px}}.foot{{margin-top:18px;text-align:center;color:var(--muted);font-size:14px}}
</style></head><body><main class="card"><h1>Đăng ký tài khoản</h1>{error_html}<form method="post" action="/register" autocomplete="on">
<label for="username">Tên đăng nhập</label><input id="username" name="username" type="text" autocomplete="username" minlength="3" maxlength="100" required autofocus>
<label for="password">Mật khẩu</label><input id="password" name="password" type="password" autocomplete="new-password" minlength="8" maxlength="256" required>
<label for="confirm_password">Nhập lại mật khẩu</label><input id="confirm_password" name="confirm_password" type="password" autocomplete="new-password" minlength="8" maxlength="256" required>
<button type="submit">Tạo tài khoản</button></form><div class="foot"><a href="/login">Đã có tài khoản? Đăng nhập</a></div></main></body></html>''')


@app.get("/register", response_class=HTMLResponse, include_in_schema=False)
def register_page(request: Request):
    if not REGISTRATION_ENABLED:
        return HTMLResponse(_register_page("Chức năng đăng ký hiện đã tắt.").body, status_code=403)
    if request.session.get("username") and _username_exists(request.session["username"]):
        return RedirectResponse("/", status_code=303)
    return _register_page()


@app.post("/register", response_class=HTMLResponse, include_in_schema=False)
def register(request: Request, username: str = Form(..., min_length=3, max_length=100),
             password: str = Form(..., min_length=8, max_length=256),
             confirm_password: str = Form(..., min_length=8, max_length=256)):
    if not REGISTRATION_ENABLED:
        return HTMLResponse(_register_page("Chức năng đăng ký hiện đã tắt.").body, status_code=403)
    username = username.strip()
    if not username:
        return HTMLResponse(_register_page("Tên đăng nhập không được để trống.").body, status_code=400)
    if password != confirm_password:
        return HTMLResponse(_register_page("Mật khẩu nhập lại không khớp.").body, status_code=400)
    if username == APP_USERNAME:
        return HTMLResponse(_register_page("Tên đăng nhập này đã được hệ thống sử dụng.").body, status_code=409)

    with _user_store_lock:
        users = _read_users()
        if username in users:
            return HTMLResponse(_register_page("Tên đăng nhập đã tồn tại.").body, status_code=409)
        salt, digest = _hash_password(password)
        users[username] = {"salt": salt, "hash": digest, "created_at": int(time.time())}
        _write_users(users)

    request.session.clear()
    request.session["username"] = username
    request.session["authenticated_at"] = int(time.time())
    return RedirectResponse("/", status_code=303)


@app.post("/logout", include_in_schema=False)
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.middleware("http")
async def require_login(request: Request, call_next):
    path = request.url.path
    public_paths = {"/login", "/register", "/logout", "/favicon.ico", "/api/health"}
    if path not in public_paths and not _username_exists(request.session.get("username", "")):
        if path.startswith("/api/"):
            return JSONResponse({"detail": "Bạn cần đăng nhập để sử dụng chức năng này."}, status_code=401)
        redirect_to = "/login?next=" + quote(_safe_next(path), safe="/")
        return RedirectResponse(redirect_to, status_code=303)
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Cache-Control", "no-store" if path in public_paths or path.startswith("/api/") else "private, no-cache")
    return response


@app.get("/api/health")
def health():
    return {"status": "ok", "ai_enabled": bool(API_KEY), "max_upload_mb": MAX_UPLOAD_MB,
            "session_ttl_minutes": TTL_SECONDS // 60}


@app.post("/api/split-text")
def split_text(req: SplitTextRequest):
    text = req.text.strip()
    if not text:
        raise HTTPException(400, "Câu cần tách đang để trống.")
    ai = get_ai() if req.use_ai else None
    duong, tu, den, source = split_row(text, req.last_area_name, ai_fallback=ai)
    return {"duong": duong, "tu": tu, "den": den, "source": source}


def _save_upload(upload: UploadFile, workdir: str) -> str:
    name = upload.filename or "upload.xlsx"
    if not name.lower().endswith(".xlsx"):
        raise HTTPException(400, "Chỉ nhận file Excel định dạng .xlsx.")
    path = os.path.join(workdir, "input.xlsx")
    limit = MAX_UPLOAD_MB * 1024 * 1024
    size = 0
    with open(path, "wb") as f:
        while chunk := upload.file.read(1024 * 1024):
            size += len(chunk)
            if size > limit:
                raise HTTPException(413, f"File lớn hơn giới hạn {MAX_UPLOAD_MB} MB.")
            f.write(chunk)
    return path


@app.post("/api/split-file")
def split_file(
    file: UploadFile = File(...),
    mode: str = Form("add-columns"),
    sheet: str = Form(""),
    column: str = Form("Tên đường"),
    order: str = Form("duong-tu-den"),
    use_ai: bool = Form(False),
    hide_source: bool = Form(False),
):
    if mode not in MODES:
        raise HTTPException(400, f"mode phải là một trong {MODES}.")
    if order not in COLUMN_ORDERS:
        raise HTTPException(400, f"order phải là một trong {tuple(COLUMN_ORDERS)}.")

    column_order = COLUMN_ORDERS[order]
    ai = get_ai() if use_ai else None
    workdir = tempfile.mkdtemp(prefix="tach_duong_")
    try:
        in_path = _save_upload(file, workdir)
        out_path = os.path.join(workdir, "output.xlsx")

        if mode == "add-columns":
            res = split_column_in_workbook(
                in_path, out_path, sheet_name=sheet.strip() or None,
                column_spec=column.strip() or "Tên đường", ai_fallback=ai,
                column_order=column_order, hide_source=hide_source,
            )
            stats = res.as_dict()
        else:
            res = process_file(
                in_path, ai_fallback=ai, column_order=column_order,
                sheet_name=sheet.strip() or "Bảng giá",
            )
            export_to_excel(res.df, out_path)
            stats = {
                "source_rows": res.total_source_rows,
                "result_rows": len(res.df),
                "rule": res.count_rule,
                "rule_parent": res.count_rule_parent,
                "verified_override": res.count_verified_override,
                "ai": res.count_ai,
                "unresolved": res.count_unresolved,
            }
    except HTTPException:
        shutil.rmtree(workdir, ignore_errors=True)
        raise
    except KeyError as e:
        shutil.rmtree(workdir, ignore_errors=True)
        raise HTTPException(400, f"Không tìm thấy sheet hoặc cột cần thiết: {e}")
    except ValueError as e:
        shutil.rmtree(workdir, ignore_errors=True)
        raise HTTPException(400, str(e))
    except Exception as e:
        shutil.rmtree(workdir, ignore_errors=True)
        raise HTTPException(500, f"Không xử lý được file: {e}")

    stem = Path(file.filename or "file").stem
    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"ket_qua_{stem}.xlsx",
        headers={"X-Stats": json.dumps(stats, ensure_ascii=True)},
        background=BackgroundTask(shutil.rmtree, workdir, ignore_errors=True),
    )


# ---------------------------------------------------------------------------
# Luồng 3 bước cho giao diện web: tải lên -> phân tích -> kết quả
# ---------------------------------------------------------------------------
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@app.exception_handler(JobError)
def job_error_handler(_request, exc: JobError):
    return JSONResponse({"detail": str(exc)}, status_code=exc.status)


class JobRequest(BaseModel):
    file_id: str = Field(..., min_length=8, max_length=64)
    mode: str = "add-columns"
    sheet: str = Field("", max_length=100)
    column: str = Field("Tên đường", max_length=100)
    order: str = "duong-tu-den"
    use_ai: bool = False
    hide_source: bool = False


@app.post("/api/upload")
def upload(file: UploadFile = File(...)):
    """Bước 1: nhận tệp, trả về danh sách sheet, cột tên đường đoán được và bảng xem trước."""
    name = file.filename or "tep.xlsx"
    if not name.lower().endswith(".xlsx"):
        raise HTTPException(400, "Chỉ nhận file Excel định dạng .xlsx.")

    def write(path):
        limit = MAX_UPLOAD_MB * 1024 * 1024
        size = 0
        with open(path, "wb") as f:
            while chunk := file.file.read(1024 * 1024):
                size += len(chunk)
                if size > limit:
                    raise HTTPException(413, f"File lớn hơn giới hạn {MAX_UPLOAD_MB} MB.")
                f.write(chunk)
        if size == 0:
            raise HTTPException(400, "Tệp rỗng.")
        return size

    sf = STORE.add_file(name, write)
    try:
        info = inspect_workbook(sf.path)
    except Exception:
        STORE.remove_file(sf.id)
        raise
    return {"file_id": sf.id, "filename": sf.filename, "size": sf.size, **info}


@app.get("/api/preview")
def preview(file_id: str, sheet: str = ""):
    """Xem trước một sheet khác của tệp đã tải lên."""
    sf = STORE.get_file(file_id)
    return inspect_workbook(sf.path, sheet or None)


@app.post("/api/jobs")
def create_job(req: JobRequest):
    """Bước 2: bắt đầu phân tích (chạy nền), trả về job_id để hỏi tiến độ."""
    if req.mode not in MODES:
        raise HTTPException(400, f"mode phải là một trong {MODES}.")
    if req.order not in COLUMN_ORDERS:
        raise HTTPException(400, f"order phải là một trong {tuple(COLUMN_ORDERS)}.")
    params = req.model_dump() if hasattr(req, "model_dump") else req.dict()
    params["sheet"] = params["sheet"].strip()
    params["column"] = params["column"].strip()
    params["order"] = COLUMN_ORDERS[req.order]
    job = STORE.add_job(req.file_id, params)
    start_job(STORE, job, get_ai() if req.use_ai else None)
    return {"job_id": job.id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    return STORE.get_job(job_id).public()


@app.get("/api/jobs/{job_id}/download")
def job_download(job_id: str):
    """Bước 3: tải tệp Excel kết quả."""
    job = STORE.get_job(job_id)
    if job.status != "done" or not job.out_path or not os.path.exists(job.out_path):
        raise HTTPException(409, "Kết quả chưa sẵn sàng hoặc đã hết hạn.")
    return FileResponse(job.out_path, media_type=XLSX_MIME, filename=job.download_name)


# SessionMiddleware phải bao ngoài middleware đăng nhập để request.session có sẵn.
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    session_cookie="tach_duong_session",
    max_age=8 * 60 * 60,
    same_site="lax",
    https_only=(APP_ENV == "production"),
)
