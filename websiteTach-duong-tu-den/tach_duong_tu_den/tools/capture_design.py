"""Chụp ảnh thiết kế 3 màn hình từ chính giao diện web (dùng làm ảnh trong tài liệu và báo cáo).

Chạy:  python tools/capture_design.py --input "BGĐ_ Update Dự thảo Bảng giá_ Thanh Oai.xlsx" --out docs/design
Cần: pip install playwright pillow && playwright install chromium
"""

import argparse
import json
import os
import socket
import subprocess
import sys
import time
import urllib.request

from PIL import Image, ImageDraw, ImageFont
from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def start_server():
    port = free_port()
    env = dict(os.environ, PORT=str(port), PYTHONPATH=ROOT)
    env.pop("GEMINI_API_KEY", None)
    proc = subprocess.Popen([sys.executable, "run_server.py"], cwd=ROOT, env=env,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    url = f"http://127.0.0.1:{port}"
    for _ in range(60):
        try:
            urllib.request.urlopen(url + "/api/health", timeout=1)
            return proc, url
        except Exception:
            time.sleep(0.25)
    proc.terminate()
    raise SystemExit("Không khởi động được máy chủ")


def contact_sheet(paths, titles, out):
    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 26)
    imgs = [Image.open(p).convert("RGB") for p in paths]
    w = 700
    imgs = [im.resize((w, int(im.height * w / im.width)), Image.LANCZOS) for im in imgs]
    h = max(im.height for im in imgs)
    pad, top = 30, 70
    sheet = Image.new("RGB", (len(imgs) * w + (len(imgs) + 1) * pad, h + top + pad), "#E9EEF5")
    d = ImageDraw.Draw(sheet)
    for i, (im, t) in enumerate(zip(imgs, titles)):
        x = pad + i * (w + pad)
        d.text((x, 20), t, fill="#0B2545", font=font)
        sheet.paste(im, (x, top))
        d.rectangle([x, top, x + w - 1, top + im.height - 1], outline="#B8C6D6")
    sheet.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Tệp Excel bảng giá dùng để chụp")
    ap.add_argument("--out", default=os.path.join(ROOT, "docs", "design"))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    o = lambda n: os.path.join(a.out, n)

    proc, url = start_server()
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(viewport={"width": 1440, "height": 900})
            pg.goto(url + "/")
            pg.wait_for_function("document.getElementById('maxMb').textContent.length > 0")
            pg.screenshot(path=o("01a_tai_du_lieu_rong.png"))

            bad = os.path.join(a.out, "_tmp.txt")
            open(bad, "w").write("x")
            pg.set_input_files("#file", bad)
            pg.screenshot(path=o("04_trang_thai_loi.png"))
            os.remove(bad)

            pg.set_input_files("#file", a.input)
            pg.wait_for_selector("#uploaded:not([hidden])")
            pg.screenshot(path=o("01_tai_du_lieu.png"), full_page=True)
            pg.screenshot(path=o("_v1.png"))
            pg.click("#to2")
            pg.wait_for_timeout(800)  # chờ trang cuộn mượt xong
            pg.screenshot(path=o("02_chay_phan_tich.png"), full_page=True)
            pg.screenshot(path=o("_v2.png"))

            # trạng thái đang chạy: giả lập một lần trả lời "đang chạy" để chụp đúng giao diện giữa chừng
            running = {"id": "demo", "status": "running", "percent": 62, "done": 79, "total": 127,
                       "message": "Đang tách dòng 79/127", "stats": None, "preview": None, "warnings": [],
                       "error": None, "download_name": None,
                       "log": ["Đã nhận tệp.", "Chế độ: thêm 3 cột vào tệp gốc.", "Đã duyệt 32/127 dòng.", "Đã duyệt 64/127 dòng."]}
            pg.route("**/api/jobs/demo", lambda r: r.fulfill(status=200, content_type="application/json", body=json.dumps(running)))
            pg.route("**/api/jobs", lambda r: r.fulfill(status=200, content_type="application/json", body='{"job_id":"demo"}')
                     if r.request.method == "POST" else r.continue_())
            pg.click("#run")
            pg.wait_for_function("document.getElementById('runPct').textContent === '62%'")
            pg.screenshot(path=o("02b_dang_chay.png"), full_page=True)
            pg.unroute("**/api/jobs/demo")
            pg.unroute("**/api/jobs")

            pg.reload()
            pg.set_input_files("#file", a.input)
            pg.wait_for_selector("#uploaded:not([hidden])")
            pg.click("#to2")
            pg.click("#run")
            pg.wait_for_selector("#step3:not([hidden])")
            pg.wait_for_timeout(800)
            pg.screenshot(path=o("03_ket_qua.png"), full_page=True)
            pg.screenshot(path=o("_v3.png"))

            m = b.new_context(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True,
                              device_scale_factor=2).new_page()
            m.goto(url + "/")
            m.set_input_files("#file", a.input)
            m.wait_for_selector("#uploaded:not([hidden])")
            m.screenshot(path=o("05_dien_thoai_1.png"))
            m.click("#to2")
            m.screenshot(path=o("05_dien_thoai_2.png"))
            m.click("#run")
            m.wait_for_selector("#step3:not([hidden])")
            m.screenshot(path=o("05_dien_thoai_3.png"))
            b.close()
        contact_sheet([o("_v1.png"), o("_v2.png"), o("_v3.png")],
                      ["Bước 1. Tải dữ liệu", "Bước 2. Chạy phân tích", "Bước 3. Kết quả"], o("00_tong_quan.png"))
        for n in ("_v1.png", "_v2.png", "_v3.png"):
            os.remove(o(n))
    finally:
        proc.terminate()
    print("Đã lưu ảnh vào", a.out)


if __name__ == "__main__":
    main()
