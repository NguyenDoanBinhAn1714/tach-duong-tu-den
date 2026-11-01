"""Test API của bản server bằng TestClient (không cần mở cổng thật)."""

import io
import os

import openpyxl
import pytest
from fastapi.testclient import TestClient

from server.app import app

client = TestClient(app)
# Giữ phiên đăng nhập cho các ca kiểm thử API phía dưới.
_login = client.post("/login", data={"username": os.environ["APP_USERNAME"], "password": os.environ["APP_PASSWORD"]})
assert _login.status_code == 200 and str(_login.url).endswith("/"), _login.text[:500]


def _xlsx_bytes():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng giá"
    ws.append(["STT", "Tên đường phố", "Xã cũ", "VT1", None, None])
    ws.append([None, None, None, "Thấp", "Cao", "Bình quân"])
    ws.append([None, "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện", "TT", 100, 120, 110])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_health():
    r = client.get("/api/health")
    assert r.status_code == 200 and r.json()["status"] == "ok"


def test_trang_chu_tra_ve_html():
    r = client.get("/")
    assert r.status_code == 200 and "text/html" in r.headers["content-type"]


def test_split_text():
    r = client.post("/api/split-text", json={"text": "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện"})
    assert r.json() == {"duong": "Quốc lộ 21B", "tu": "cầu Thạch Bích", "den": "ngã ba bệnh viện", "source": "rule"}


def test_split_text_rong_bi_tu_choi():
    assert client.post("/api/split-text", json={"text": "   "}).status_code == 400
    assert client.post("/api/split-text", json={"text": ""}).status_code == 422


@pytest.mark.parametrize("mode, rows", [("add-columns", None), ("price-table", 3)])
def test_split_file_hai_che_do(mode, rows):
    files = {"file": ("a.xlsx", _xlsx_bytes(), "application/octet-stream")}
    r = client.post("/api/split-file", files=files, data={"mode": mode})
    assert r.status_code == 200
    ws = openpyxl.load_workbook(io.BytesIO(r.content)).active
    if mode == "add-columns":
        assert [ws.cell(3, c).value for c in (3, 4, 5)] == ["Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện"]
    else:
        assert ws.max_row - 1 == rows


def test_file_sai_dinh_dang_bi_tu_choi():
    r = client.post("/api/split-file", files={"file": ("a.txt", b"hello", "text/plain")})
    assert r.status_code == 400


def test_sheet_khong_ton_tai_bao_loi_ro_rang():
    files = {"file": ("a.xlsx", _xlsx_bytes(), "application/octet-stream")}
    r = client.post("/api/split-file", files=files, data={"sheet": "Khong co"})
    assert r.status_code == 400 and "Không có sheet" in r.json()["detail"]


def test_tham_so_sai_bi_tu_choi():
    files = {"file": ("a.xlsx", _xlsx_bytes(), "application/octet-stream")}
    assert client.post("/api/split-file", files=files, data={"mode": "abc"}).status_code == 400
    assert client.post("/api/split-file", files=files, data={"order": "abc"}).status_code == 400


def test_file_qua_lon_bi_tu_choi(monkeypatch):
    import server.app as srv
    monkeypatch.setattr(srv, "MAX_UPLOAD_MB", 0)
    files = {"file": ("a.xlsx", _xlsx_bytes(), "application/octet-stream")}
    assert client.post("/api/split-file", files=files).status_code == 413
