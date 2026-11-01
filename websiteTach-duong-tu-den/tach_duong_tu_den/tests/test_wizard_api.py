"""Test luồng 3 bước qua API: upload -> jobs -> kết quả."""

import io
import os
import time

import openpyxl
import pytest
from fastapi.testclient import TestClient

import server.jobs as jobs
from server.app import app

client = TestClient(app)
# Giữ phiên đăng nhập cho các ca kiểm thử API phía dưới.
_login = client.post("/login", data={"username": os.environ["APP_USERNAME"], "password": os.environ["APP_PASSWORD"]})
assert _login.status_code == 200 and str(_login.url).endswith("/"), _login.text[:500]
TEXT = "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện"


def make_xlsx(sheet="Bảng giá"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(["STT", "Tên đường phố", "Xã cũ", "VT1", None, None])
    ws.append([None, None, None, "Thấp", "Cao", "Bình quân"])
    ws.append([None, TEXT, "TT", 100, 120, 110])
    ws.append([None, "Khu đấu giá Sau Trại", "TT", 50, 60, 55])
    other = wb.create_sheet("Ghi chú")
    other.append(["Nội dung khác"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def upload(data=None, name="mau.xlsx"):
    return client.post("/api/upload", files={"file": (name, make_xlsx() if data is None else data, "application/octet-stream")})


def run_job(file_id, **kw):
    body = {"file_id": file_id, **kw}
    r = client.post("/api/jobs", json=body)
    assert r.status_code == 200, r.text
    job_id = r.json()["job_id"]
    for _ in range(100):
        st = client.get(f"/api/jobs/{job_id}").json()
        if st["status"] in ("done", "error"):
            return job_id, st
        time.sleep(0.05)
    pytest.fail("job không kết thúc")


def test_upload_nhan_dien_sheet_va_cot_ten_duong():
    r = upload()
    assert r.status_code == 200
    j = r.json()
    assert j["filename"] == "mau.xlsx" and j["size"] > 0
    assert [s["name"] for s in j["sheets"]] == ["Bảng giá", "Ghi chú"]
    assert j["detected"] == {"sheet": "Bảng giá", "found": True, "column": "B", "header": "Tên đường phố", "header_row": 1}
    assert j["preview"]["highlight"] == "B" and j["preview"]["rows"][2][1] == TEXT


def test_upload_sheet_khong_co_cot_ten_duong_thi_found_false():
    j = client.get("/api/preview", params={"file_id": upload().json()["file_id"], "sheet": "Ghi chú"}).json()
    assert j["detected"]["found"] is False and j["preview"]["sheet"] == "Ghi chú"


@pytest.mark.parametrize("name, data, code", [
    ("a.txt", b"hello", 400),
    ("a.xlsx", b"", 400),
    ("a.xlsx", b"khong phai excel", 400),
])
def test_upload_sai_bi_tu_choi(name, data, code):
    before = len(jobs.STORE.files)
    r = upload(data, name)
    assert r.status_code == code and "detail" in r.json()
    assert len(jobs.STORE.files) == before          # tệp hỏng không bị giữ lại


def test_upload_qua_lon(monkeypatch):
    import server.app as srv
    monkeypatch.setattr(srv, "MAX_UPLOAD_MB", 0)
    assert upload().status_code == 413


def test_preview_file_khong_ton_tai():
    assert client.get("/api/preview", params={"file_id": "x" * 32}).status_code == 404


def test_job_them_cot_chay_het_vong_doi_va_tai_ve():
    fid = upload().json()["file_id"]
    job_id, st = run_job(fid, mode="add-columns", sheet="Bảng giá", column="Tên đường")
    assert st["status"] == "done" and st["percent"] == 100
    assert st["stats"]["processed"] == 2 and st["stats"]["rule"] == 1 and st["stats"]["unresolved"] == 1
    assert st["stats"]["first_new_column"] == "C" and st["stats"]["unresolved_samples"] == ["Khu đấu giá Sau Trại"]
    assert st["preview"]["columns"] == ["Dòng", "Tên đường phố (gốc)", "Đường", "Từ", "Đến", "Cách xử lý"]
    assert st["preview"]["rows"][0][2:5] == ["Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện"]
    assert any("Hoàn tất" in line for line in st["log"])
    d = client.get(f"/api/jobs/{job_id}/download")
    assert d.status_code == 200 and "ket_qua_mau.xlsx" in d.headers["content-disposition"]
    ws = openpyxl.load_workbook(io.BytesIO(d.content))["Bảng giá"]
    assert [ws.cell(1, c).value for c in (3, 4, 5)] == ["Đường (tách)", "Từ (tách)", "Đến (tách)"]
    assert ws.cell(3, 4).value == "cầu Thạch Bích"


def test_job_bang_doi_chieu_gia():
    fid = upload().json()["file_id"]
    job_id, st = run_job(fid, mode="price-table", order="tu-den-duong")
    assert st["status"] == "done" and st["stats"]["result_rows"] == 6   # 2 dòng x 3 mức giá (VT1)
    assert st["preview"]["columns"][:3] == ["Từ", "Đến", "Đường"]
    d = client.get(f"/api/jobs/{job_id}/download")
    assert d.status_code == 200 and "doi_chieu_gia_mau.xlsx" in d.headers["content-disposition"]


def test_job_sheet_khong_ton_tai_bao_loi_than_thien():
    fid = upload().json()["file_id"]
    _, st = run_job(fid, mode="add-columns", sheet="Khong co")
    assert st["status"] == "error" and "Không có sheet" in st["error"] and st["stats"] is None


def test_job_khong_the_tai_khi_chua_xong_hoac_loi():
    fid = upload().json()["file_id"]
    job_id, st = run_job(fid, mode="add-columns", sheet="Khong co")
    assert client.get(f"/api/jobs/{job_id}/download").status_code == 409


@pytest.mark.parametrize("body, code", [
    ({"file_id": "x" * 32}, 404),
    ({"file_id": "x" * 32, "mode": "abc"}, 400),
    ({"file_id": "x" * 32, "order": "abc"}, 400),
    ({"file_id": "ngan"}, 422),
])
def test_tao_job_tham_so_sai(body, code):
    assert client.post("/api/jobs", json=body).status_code == code


def test_job_khong_ton_tai():
    assert client.get("/api/jobs/" + "0" * 32).status_code == 404
    assert client.get("/api/jobs/" + "0" * 32 + "/download").status_code == 404


def test_may_chu_ban_thi_tra_429(monkeypatch):
    fid = upload().json()["file_id"]
    monkeypatch.setattr(jobs, "MAX_ACTIVE_JOBS", 0)
    assert client.post("/api/jobs", json={"file_id": fid}).status_code == 429


def test_tep_va_job_het_han_bi_xoa_khoi_o_dia():
    import os
    fid = upload().json()["file_id"]
    job_id, _ = run_job(fid)
    sf, job = jobs.STORE.files[fid], jobs.STORE.jobs[job_id]
    assert os.path.exists(sf.dir) and os.path.exists(job.dir)
    jobs.STORE.cleanup(now=time.time() + jobs.TTL_SECONDS + 5)
    assert fid not in jobs.STORE.files and job_id not in jobs.STORE.jobs
    assert not os.path.exists(sf.dir) and not os.path.exists(job.dir)
