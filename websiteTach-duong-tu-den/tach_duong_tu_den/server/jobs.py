"""Tệp tải lên và tác vụ phân tích chạy nền cho giao diện 3 bước.

Luồng: upload -> tạo job -> hỏi tiến độ -> tải kết quả.
Mọi tệp nằm trong thư mục tạm riêng, đặt tên bằng mã ngẫu nhiên (không dùng tên người dùng gửi
lên làm đường dẫn) và bị xoá sau SESSION_TTL_SECONDS giây (mặc định 1 giờ).
"""

import os
import shutil
import tempfile
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

from core.generic import DEFAULT_KEYWORD, find_header, pick_sheet, resolve_column, split_column_in_workbook
from core.processor import export_to_excel, process_file

TTL_SECONDS = int(os.getenv("SESSION_TTL_SECONDS", "3600"))
MAX_STORED_FILES = int(os.getenv("MAX_STORED_FILES", "50"))
MAX_ACTIVE_JOBS = int(os.getenv("MAX_ACTIVE_JOBS", "4"))
PREVIEW_ROWS = 12
PREVIEW_COLS = 10
RESULT_PREVIEW_ROWS = 40
LOG_LIMIT = 200

SOURCE_LABELS = {
    "rule": "Quy tắc",
    "rule-parent": "Kế thừa khu vực",
    "verified-override": "Đã đối chiếu",
    "ai": "AI (Gemini)",
    "unresolved": "Giữ nguyên",
}


class JobError(Exception):
    """Lỗi có thông điệp thân thiện để hiện cho người dùng."""

    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


@dataclass
class StoredFile:
    id: str
    dir: str
    path: str
    filename: str
    size: int
    created: float = field(default_factory=time.time)


@dataclass
class Job:
    id: str
    file_id: str
    params: dict
    dir: str
    status: str = "queued"        # queued | running | done | error
    done: int = 0
    total: int = 0
    message: str = "Đang chờ"
    log: list = field(default_factory=list)
    stats: Optional[dict] = None
    preview: Optional[dict] = None
    warnings: list = field(default_factory=list)
    out_path: Optional[str] = None
    download_name: str = ""
    error: Optional[str] = None
    created: float = field(default_factory=time.time)
    finished: Optional[float] = None

    def add_log(self, text: str):
        self.log.append(text)
        if len(self.log) > LOG_LIMIT:
            del self.log[: len(self.log) - LOG_LIMIT]

    def public(self) -> dict:
        percent = int(100 * self.done / self.total) if self.total else (100 if self.status == "done" else 0)
        return {
            "id": self.id,
            "status": self.status,
            "percent": min(percent, 100),
            "done": self.done,
            "total": self.total,
            "message": self.message,
            "log": list(self.log),
            "stats": self.stats,
            "preview": self.preview,
            "warnings": list(self.warnings),
            "error": self.error,
            "download_name": self.download_name if self.status == "done" else None,
        }


class Store:
    def __init__(self):
        self._lock = threading.Lock()
        self.files: dict = {}
        self.jobs: dict = {}

    # ---- tệp tải lên
    def add_file(self, filename: str, data_writer) -> StoredFile:
        self.cleanup()
        with self._lock:
            if len(self.files) >= MAX_STORED_FILES:
                raise JobError("Máy chủ đang lưu quá nhiều tệp, hãy thử lại sau ít phút.", 429)
        fid = uuid.uuid4().hex
        d = tempfile.mkdtemp(prefix="tach_up_")
        path = os.path.join(d, "input.xlsx")
        try:
            size = data_writer(path)
        except Exception:
            shutil.rmtree(d, ignore_errors=True)
            raise
        sf = StoredFile(id=fid, dir=d, path=path, filename=filename, size=size)
        with self._lock:
            self.files[fid] = sf
        return sf

    def remove_file(self, fid: str):
        with self._lock:
            sf = self.files.pop(fid, None)
        if sf:
            shutil.rmtree(sf.dir, ignore_errors=True)

    def get_file(self, fid: str) -> StoredFile:
        with self._lock:
            sf = self.files.get(fid)
        if not sf:
            raise JobError("Không tìm thấy tệp (có thể đã hết hạn). Hãy tải lên lại.", 404)
        return sf

    # ---- tác vụ
    def add_job(self, file_id: str, params: dict) -> Job:
        self.cleanup()
        self.get_file(file_id)
        with self._lock:
            active = sum(1 for j in self.jobs.values() if j.status in ("queued", "running"))
            if active >= MAX_ACTIVE_JOBS:
                raise JobError("Máy chủ đang bận xử lý các tệp khác, hãy thử lại sau ít phút.", 429)
            job = Job(id=uuid.uuid4().hex, file_id=file_id, params=params, dir=tempfile.mkdtemp(prefix="tach_job_"))
            self.jobs[job.id] = job
        return job

    def get_job(self, jid: str) -> Job:
        with self._lock:
            job = self.jobs.get(jid)
        if not job:
            raise JobError("Không tìm thấy tác vụ (có thể đã hết hạn).", 404)
        return job

    def cleanup(self, now: Optional[float] = None):
        now = now or time.time()
        with self._lock:
            old_files = [k for k, v in self.files.items() if now - v.created > TTL_SECONDS]
            old_jobs = [k for k, v in self.jobs.items()
                        if now - v.created > TTL_SECONDS and v.status in ("done", "error")]
            dirs = [self.files.pop(k).dir for k in old_files] + [self.jobs.pop(k).dir for k in old_jobs]
        for d in dirs:
            shutil.rmtree(d, ignore_errors=True)


STORE = Store()


# ---------------------------------------------------------------------------
# Phân tích tệp vừa tải lên: các sheet, cột tên đường, bảng xem trước
# ---------------------------------------------------------------------------
def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("\n", " ").strip()
    return s if len(s) <= 60 else s[:57] + "..."


def detect(ws) -> dict:
    """Đoán cột tên đường trong sheet. found=False khi không thấy tiêu đề phù hợp."""
    found = find_header(ws, DEFAULT_KEYWORD)
    if not found:
        return {"sheet": ws.title, "found": False, "column": "", "header": "", "header_row": 0}
    col, row = found
    return {
        "sheet": ws.title,
        "found": True,
        "column": get_column_letter(col),
        "header": _cell_text(ws.cell(row, col).value),
        "header_row": row,
    }


def sheet_preview(ws, highlight: str = "") -> dict:
    last_col = min(ws.max_column or 1, PREVIEW_COLS)
    if highlight:
        try:
            last_col = max(last_col, openpyxl.utils.column_index_from_string(highlight))
        except ValueError:
            pass
        last_col = min(last_col, 26)
    rows = []
    for r in range(1, min(ws.max_row or 1, PREVIEW_ROWS) + 1):
        rows.append([_cell_text(ws.cell(r, c).value) for c in range(1, last_col + 1)])
    return {"sheet": ws.title, "columns": [get_column_letter(c) for c in range(1, last_col + 1)],
            "rows": rows, "highlight": highlight, "total_rows": ws.max_row or 0}


def inspect_workbook(path: str, sheet: Optional[str] = None) -> dict:
    try:
        wb = openpyxl.load_workbook(path)
    except Exception:
        raise JobError("Không đọc được tệp. Hãy kiểm tra đây là tệp Excel .xlsx hợp lệ.", 400)
    try:
        sheets = [{"name": ws.title, "rows": ws.max_row or 0, "cols": ws.max_column or 0} for ws in wb.worksheets]
        if sheet:
            if sheet not in wb.sheetnames:
                raise JobError(f"Không có sheet {sheet!r}.", 400)
            ws = wb[sheet]
        else:
            ws = pick_sheet(wb)
        det = detect(ws)
        return {"sheets": sheets, "detected": det, "preview": sheet_preview(ws, det["column"])}
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Chạy phân tích (trong luồng nền)
# ---------------------------------------------------------------------------
def _fmt_money(v) -> str:
    try:
        return f"{float(v):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "" if v is None else str(v)


def run_job(job: Job, sf: StoredFile, ai):
    p = job.params
    mode = p["mode"]
    order = p["order"]
    try:
        job.status = "running"
        job.message = "Đang đọc tệp"
        job.add_log(f"Đã nhận tệp {sf.filename} ({sf.size // 1024} KB).")
        if p.get("use_ai"):
            job.add_log("Có dùng AI cho câu quy tắc không tách được." if ai
                        else "Máy chủ chưa có API key Gemini nên bỏ qua bước AI.")
        else:
            job.add_log("Chỉ dùng quy tắc, không gọi AI.")

        last_logged = [-1]

        def progress(done, total, _msg):
            job.done, job.total = done, total
            job.message = f"Đang tách dòng {done}/{total}"
            step = max(total // 4, 1)
            if done // step != last_logged[0] and (done % step == 0 or done == total):
                last_logged[0] = done // step
                job.add_log(f"Đã duyệt {done}/{total} dòng.")

        out_path = os.path.join(job.dir, "output.xlsx")
        stem = Path(sf.filename).stem or "file"
        if mode == "add-columns":
            job.add_log("Chế độ: thêm 3 cột vào tệp gốc.")
            res = split_column_in_workbook(
                sf.path, out_path, sheet_name=p["sheet"] or None, column_spec=p["column"] or DEFAULT_KEYWORD,
                ai_fallback=ai, column_order=order, hide_source=p["hide_source"], progress_cb=progress)
            job.add_log(f"Sheet «{res.sheet_name}», cột {res.column_letter}: tách {res.processed_rows} dòng, "
                        f"bỏ qua {res.skipped_rows} dòng tiêu đề nhóm.")
            job.stats = {
                "processed": res.processed_rows, "skipped": res.skipped_rows, "result_rows": None,
                "rule": res.count_rule, "rule_parent": res.count_rule_parent,
                "verified_override": res.count_verified_override, "ai": res.count_ai,
                "unresolved": res.count_unresolved, "sheet": res.sheet_name, "column": res.column_letter,
                "placement": res.placement, "first_new_column": res.first_new_column,
            }
            if res.warning:
                job.warnings.append(res.warning)
            job.preview = {
                "columns": ["Dòng", "Tên đường phố (gốc)"] + [f"{c}" for c in order] + ["Cách xử lý"],
                "rows": [[str(r["row"]), r["text"]] + [{"Đường": r["duong"], "Từ": r["tu"], "Đến": r["den"]}[c] for c in order]
                         + [SOURCE_LABELS.get(r["source"], r["source"])] for r in res.rows[:RESULT_PREVIEW_ROWS]],
                "total": res.processed_rows,
            }
            unresolved = res.unresolved_samples
            job.download_name = f"ket_qua_{stem}.xlsx"
        else:
            job.add_log("Chế độ: bảng đối chiếu giá.")
            res = process_file(sf.path, ai_fallback=ai, column_order=order, progress_cb=progress,
                               sheet_name=p["sheet"] or "Bảng giá")
            job.add_log("Đang ghi tệp Excel kết quả.")
            export_to_excel(res.df, out_path)
            job.add_log(f"Từ {res.total_source_rows} dòng gốc tạo ra {len(res.df)} dòng kết quả.")
            job.stats = {
                "processed": res.total_source_rows, "skipped": None, "result_rows": len(res.df),
                "rule": res.count_rule, "rule_parent": res.count_rule_parent,
                "verified_override": res.count_verified_override, "ai": res.count_ai,
                "unresolved": res.count_unresolved, "sheet": p["sheet"] or "Bảng giá", "column": "",
                "placement": "", "first_new_column": "",
            }
            head = res.df.head(RESULT_PREVIEW_ROWS)
            cols = [c for c in head.columns if c != "STT"]
            job.preview = {
                "columns": cols,
                "rows": [[_fmt_money(row[c]) if c == "Đơn giá (đồng)" else _cell_text(row[c]) for c in cols]
                         for _, row in head.iterrows()],
                "total": len(res.df),
            }
            unresolved = res.unresolved_samples
            job.download_name = f"doi_chieu_gia_{stem}.xlsx"

        if unresolved:
            job.add_log(f"{len(unresolved)} câu không có mô tả đoạn, được giữ nguyên.")
        job.stats["unresolved_samples"] = list(unresolved[:20])
        job.out_path = out_path
        job.done = job.total or job.done
        job.message = "Hoàn tất"
        job.add_log("Hoàn tất.")
        job.status = "done"
    except (ValueError, KeyError) as e:
        job.status, job.error = "error", (f"Không tìm thấy sheet hoặc cột cần thiết: {e}" if isinstance(e, KeyError) else str(e))
        job.message = "Có lỗi"
    except Exception as e:  # noqa: BLE001 - báo lỗi thân thiện cho người dùng
        job.status, job.error = "error", f"Không xử lý được tệp: {e}"
        job.message = "Có lỗi"
    finally:
        job.finished = time.time()


def start_job(store: Store, job: Job, ai):
    sf = store.get_file(job.file_id)
    t = threading.Thread(target=run_job, args=(job, sf, ai), daemon=True)
    t.start()
    return t
