"""Test giao diện web 3 bước trên trình duyệt thật (Playwright + Chromium).

Tự bỏ qua nếu máy không có Playwright hoặc trình duyệt. Cài: pip install playwright && playwright install chromium
"""

import io
import os
import socket
import subprocess
import sys
import time
import urllib.request

import openpyxl
import pytest

sync_api = pytest.importorskip("playwright.sync_api")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEXT = "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện"


def build_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng giá"
    ws.append(["STT", "Tên đường phố", "Xã cũ", "VT1", None, None])
    ws.append([None, None, None, "Thấp", "Cao", "Bình quân"])
    ws.append(["A", "Xã Thanh Oai", None, None, None, None])
    ws.append([None, TEXT, "TT", 100, 120, 110])
    ws.append([None, "Khu đất dịch vụ Cây Sữa", None, None, None, None])
    ws.append([None, "Trục chính", None, 90, 110, 100])
    ws.append([None, "Khu đấu giá Sau Trại", "TT", 50, 60, 55])
    wb.create_sheet("Ghi chú").append(["Ghi chú chung"])
    wb.save(path)
    return str(path)


@pytest.fixture(scope="module")
def base_url():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    env = dict(os.environ, PORT=str(port), HOST="127.0.0.1", PYTHONPATH=ROOT, APP_ENV="test", APP_USERNAME="test-admin", APP_PASSWORD="TestPassword!234", SESSION_SECRET="test-only-session-secret-please-do-not-use")
    env.pop("GEMINI_API_KEY", None)
    proc = subprocess.Popen([sys.executable, "run_server.py"], cwd=ROOT, env=env,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    url = f"http://127.0.0.1:{port}"
    for _ in range(60):
        try:
            urllib.request.urlopen(url + "/api/health", timeout=1)
            break
        except Exception:
            time.sleep(0.25)
    else:
        proc.terminate()
        pytest.skip("không khởi động được máy chủ")
    yield url
    proc.terminate()


@pytest.fixture(scope="module")
def browser():
    with sync_api.sync_playwright() as p:
        try:
            b = p.chromium.launch()
        except Exception as e:  # noqa: BLE001
            pytest.skip(f"không mở được Chromium: {e}")
        yield b
        b.close()


@pytest.fixture
def page(browser, base_url):
    ctx = browser.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
    pg = ctx.new_page()
    pg.js_errors = []
    pg.on("pageerror", lambda e: pg.js_errors.append(str(e)))
    # Bỏ qua dòng "Failed to load resource" vì trình duyệt luôn ghi nó khi máy chủ cố ý trả 400/404 trong test lỗi.
    pg.on("console", lambda m: pg.js_errors.append(m.text)
          if m.type == "error" and not m.text.startswith("Failed to load resource") else None)
    pg.goto(base_url + "/")
    pg.wait_for_selector("#loginSubmit")
    pg.fill("#username", "test-admin")
    pg.fill("#password", "TestPassword!234")
    pg.click("#loginSubmit")
    pg.wait_for_url(base_url + "/")
    pg.wait_for_function("document.getElementById('maxMb').textContent.length > 0")
    yield pg
    assert pg.js_errors == [], pg.js_errors
    ctx.close()


@pytest.fixture
def xlsx(tmp_path):
    return build_xlsx(tmp_path / "mau.xlsx")


def upload(page, path):
    page.set_input_files("#file", path)
    page.wait_for_selector("#uploaded:not([hidden])")


def step_visible(page):
    return [i for i in (1, 2, 3) if page.is_visible(f"#step{i}")]


def test_buoc_1_tai_len_nhan_dien_sheet_va_cot(page, xlsx):
    assert step_visible(page) == [1]
    assert page.is_disabled("#to2")                                  # chưa có tệp thì chưa đi tiếp được
    upload(page, xlsx)
    assert page.inner_text("#fName") == "mau.xlsx"
    assert "2 sheet" in page.inner_text("#fMeta")
    assert page.eval_on_selector("#sheetTabs [aria-pressed=true]", "e => e.textContent") == "Bảng giá"
    assert "cột B" in page.inner_text("#detectNote") and "Tên đường phố" in page.inner_text("#detectNote")
    assert page.is_enabled("#to2")


def test_doi_sheet_khong_co_cot_ten_duong_hien_canh_bao(page, xlsx):
    upload(page, xlsx)
    page.click("#sheetTabs >> text=Ghi chú")
    page.wait_for_function("document.getElementById('detectNote').textContent.includes('Chưa tìm thấy')")
    assert page.eval_on_selector("#sheetTabs [aria-pressed=true]", "e => e.textContent") == "Ghi chú"


def test_luong_day_du_them_3_cot_va_tai_ket_qua(page, xlsx, tmp_path):
    upload(page, xlsx)
    page.click("#to2")
    assert step_visible(page) == [2]
    assert page.input_value("#column") == "B"
    assert "mau.xlsx" in page.inner_text("#summary")
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")
    assert step_visible(page) == [3]
    stats = page.inner_text("#stats")
    assert "Dòng đã tách" in stats and "Quy tắc" in stats
    assert "ngay sau cột B" in page.inner_text("#whereNote")
    assert "Quốc lộ 21B" in page.inner_text("#resultTable")
    assert page.is_visible("#unresolvedBox") and "Khu đấu giá Sau Trại" in page.text_content("#unresolvedBox")

    with page.expect_download() as dl:
        page.click("#download")
    out = tmp_path / dl.value.suggested_filename
    dl.value.save_as(out)
    ws = openpyxl.load_workbook(out)["Bảng giá"]
    assert [ws.cell(1, c).value for c in (3, 4, 5)] == ["Đường (tách)", "Từ (tách)", "Đến (tách)"]
    assert [ws.cell(4, c).value for c in (3, 4, 5)] == ["Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện"]
    assert ws.cell(6, 3).value == "Khu đất dịch vụ Cây Sữa - Trục chính"


def test_che_do_bang_doi_chieu_gia_va_doi_thu_tu_cot(page, xlsx, tmp_path):
    upload(page, xlsx)
    page.click("#to2")
    page.check("input[name=mode][value=price-table]")
    assert page.is_disabled("#column") and page.is_disabled("#hideSource")
    page.select_option("#order", "tu-den-duong")
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")
    assert page.inner_text("#resultTable thead").split()[:3] == ["Từ", "Đến", "Đường"]
    with page.expect_download() as dl:
        page.click("#download")
    assert dl.value.suggested_filename == "doi_chieu_gia_mau.xlsx"
    out = tmp_path / "r.xlsx"
    dl.value.save_as(out)
    assert openpyxl.load_workbook(out).active.max_row - 1 == 9      # 3 dòng dữ liệu x 3 mức giá


def test_dang_chay_hien_tien_do_va_khoa_form(page, xlsx):
    def slow(route):
        time.sleep(0.6)
        route.continue_()
    page.route("**/api/jobs/*", lambda r: slow(r) if r.request.method == "GET" and not r.request.url.endswith("/download") else r.continue_())
    upload(page, xlsx)
    page.click("#to2")
    page.click("#run")
    page.wait_for_selector("#runPanel:not([hidden])")
    assert page.is_disabled("#run") and page.is_disabled("#back1")
    assert page.get_attribute("#bar", "role") == "progressbar"
    page.wait_for_selector("#step3:not([hidden])")


def test_loi_phan_tich_hien_thong_bao_va_cho_thu_lai(page, xlsx):
    upload(page, xlsx)
    page.click("#to2")
    page.fill("#column", "Cột không tồn tại")
    page.click("#run")
    page.wait_for_selector("#runError:not([hidden])")
    assert "Không tìm thấy cột" in page.inner_text("#runError")
    assert page.is_enabled("#run") and page.inner_text("#runLabel") == "Thử lại"
    page.fill("#column", "B")
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")


def test_tep_sai_dinh_dang_va_tep_hong_bi_tu_choi(page, tmp_path):
    bad = tmp_path / "a.txt"
    bad.write_text("xin chào")
    page.set_input_files("#file", str(bad))
    assert "Chỉ nhận tệp Excel" in page.inner_text("#uploadError")
    broken = tmp_path / "hong.xlsx"
    broken.write_text("khong phai excel")
    page.set_input_files("#file", str(broken))
    page.wait_for_function("document.getElementById('uploadError').textContent.includes('Không đọc được')")
    assert page.is_disabled("#to2") and page.is_hidden("#uploaded")


def test_dieu_huong_buoc_va_tro_lai(page, xlsx):
    upload(page, xlsx)
    assert page.is_disabled("#stepper li[data-step='2'] button") is False
    assert page.is_disabled("#stepper li[data-step='3'] button")        # chưa có kết quả
    page.click("#stepper li[data-step='2'] button")
    assert step_visible(page) == [2]
    assert page.get_attribute("#stepper li[data-step='2'] button", "aria-current") == "step"
    assert page.evaluate("document.activeElement.id") == "t2"           # focus về tiêu đề bước
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")
    page.click("#stepper li[data-step='1'] button")
    assert step_visible(page) == [1]
    assert page.is_enabled("#stepper li[data-step='3'] button")         # kết quả vẫn xem lại được
    page.click("#stepper li[data-step='3'] button")
    assert step_visible(page) == [3]
    page.click("#rerun")
    assert step_visible(page) == [2]
    page.click("#back1")
    page.click("#to2")
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")
    page.click("#restart")
    assert step_visible(page) == [1] and page.is_hidden("#uploaded") and page.is_disabled("#to2")


def test_thu_nhanh_mot_cau(page, xlsx):
    upload(page, xlsx)
    page.click("#to2")
    page.click(".quick summary")
    page.click("#quickBtn")
    page.wait_for_selector("#quickRes:not([hidden])")
    assert [page.inner_text(i) for i in ("#qD", "#qT", "#qE")][:2] == ["Quốc lộ 21B", "cầu Thạch Bích"]
    page.fill("#quickText", "Đường Láng - Hòa Lạc, đoạn gần cầu vượt")
    page.click("#quickBtn")
    page.wait_for_function("document.getElementById('quickNote').textContent.includes('giữ nguyên câu')")


def test_noi_dung_nguoi_dung_khong_bi_chen_thanh_html(page, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng giá"
    ws.append(["STT", "Tên đường phố", "Xã cũ", "VT1", None, None])
    ws.append([None, '<img src=x onerror="window.__xss=1">Đường A: Đoạn từ <b>X</b> đến Y', "TT", 1, 2, 3])
    p = tmp_path / "xss.xlsx"
    wb.save(p)
    upload(page, str(p))
    page.click("#to2")
    page.click("#run")
    page.wait_for_selector("#step3:not([hidden])")
    time.sleep(0.3)
    assert page.evaluate("window.__xss") is None
    assert page.locator("#resultTable img").count() == 0
    assert "<img" in page.inner_text("#resultTable")


def test_dien_thoai_khong_tran_ngang_va_chay_het_luong(browser, base_url, xlsx):
    ctx = browser.new_context(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True)
    pg = ctx.new_page()
    pg.goto(base_url + "/")
    pg.wait_for_selector("#loginSubmit")
    pg.fill("#username", "test-admin")
    pg.fill("#password", "TestPassword!234")
    pg.click("#loginSubmit")
    pg.wait_for_url(base_url + "/")
    for _ in range(3):
        pg.set_input_files("#file", xlsx) if pg.is_visible("#step1") else None
        if pg.is_visible("#step1"):
            pg.wait_for_selector("#uploaded:not([hidden])")
            pg.click("#to2")
        elif pg.is_visible("#step2"):
            pg.click("#run")
            pg.wait_for_selector("#step3:not([hidden])")
        assert pg.evaluate("document.documentElement.scrollWidth <= window.innerWidth + 1"), "bị tràn ngang"
    labels = pg.eval_on_selector_all("#stepper .step-label", "els => els.filter(e => e.offsetParent !== null).map(e => e.textContent)")
    assert labels == ["Kết quả"]                                      # điện thoại chỉ hiện tên bước hiện tại
    ctx.close()
