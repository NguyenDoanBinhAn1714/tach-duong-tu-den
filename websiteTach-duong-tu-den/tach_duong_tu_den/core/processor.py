"""
processor.py
------------
Đọc file Excel "Bảng giá đất" gốc, áp dụng splitter.py cho từng dòng,
rồi "unpivot" bảng giá (4 vị trí x 3 mức giá) thành dạng bảng dài,
trả về pandas DataFrame đúng 8 cột: STT, Đường, Từ, Đến, Vị trí,
Mức giá, Đơn giá (đồng), Dòng nguồn.
"""

from dataclasses import dataclass, field

import openpyxl
import pandas as pd

from .splitter import GENERIC_LABELS, is_area_heading, split_row

SHEET_NAME = "Bảng giá"
MUC_GIA_LABELS = ["Giá thấp nhất", "Giá cao nhất", "Giá trung bình"]
VITRI_ORDER = ["VT1", "VT2", "VT3", "VT4"]

# Cột trong sheet gốc: B = Tên đường phố, D..F = VT1, G..I = VT2, J..L = VT3, M..O = VT4
PRICE_COLUMNS = {
    "VT1": (4, 5, 6),
    "VT2": (7, 8, 9),
    "VT3": (10, 11, 12),
    "VT4": (13, 14, 15),
}


@dataclass
class ProcessResult:
    df: pd.DataFrame
    total_source_rows: int = 0
    count_rule: int = 0
    count_rule_parent: int = 0
    count_verified_override: int = 0
    count_ai: int = 0
    count_unresolved: int = 0
    unresolved_samples: list = field(default_factory=list)


def read_source_rows(input_path: str, sheet_name: str = SHEET_NAME):
    """Đọc sheet gốc, trả về danh sách dict {row, text, prices} cho các dòng có giá thật."""
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb[sheet_name]

    rows = []
    last_area_name = None
    for r in range(3, ws.max_row + 1):
        raw = ws.cell(r, 2).value
        text = str(raw).strip() if raw else None

        if is_area_heading(text):
            last_area_name = text
            continue

        if text and text.lower() not in GENERIC_LABELS and not is_area_heading(text):
            # Không ghi đè khu vực khi đây là tuyến đường bình thường.
            pass

        if not text:
            continue

        prices = {}
        for vt, (c1, c2, c3) in PRICE_COLUMNS.items():
            prices[vt] = (ws.cell(r, c1).value, ws.cell(r, c2).value, ws.cell(r, c3).value)

        if all(v is None for vals in prices.values() for v in vals):
            continue  # dòng tiêu đề nhóm, không có giá -> bỏ qua

        rows.append({"row": r, "text": text, "last_area_name": last_area_name, "prices": prices})

    return rows


def process_file(input_path: str, ai_fallback=None, progress_cb=None,
                  column_order=("Đường", "Từ", "Đến"), sheet_name: str = SHEET_NAME) -> ProcessResult:
    """
    Xử lý toàn bộ file gốc, trả về ProcessResult (bao gồm DataFrame kết quả
    và thống kê minh bạch về số dòng xử lý theo từng phương pháp).

    progress_cb: callback(done:int, total:int, message:str) tuỳ chọn, dùng
                 để cập nhật thanh tiến trình trên giao diện.
    column_order: thứ tự 3 cột địa chỉ trong file xuất ra, ví dụ
                  ("Đường", "Từ", "Đến") hoặc ("Từ", "Đến", "Đường").
    """
    source_rows = read_source_rows(input_path, sheet_name)
    total = len(source_rows)

    output_rows = []
    stt = 1
    result = ProcessResult(df=None, total_source_rows=total)

    for i, item in enumerate(source_rows, start=1):
        duong, tu, den, source = split_row(
            item["text"], item["last_area_name"], ai_fallback=ai_fallback
        )

        if source == "rule":
            result.count_rule += 1
        elif source == "rule-parent":
            result.count_rule_parent += 1
        elif source == "verified-override":
            result.count_verified_override += 1
        elif source == "ai":
            result.count_ai += 1
        else:
            result.count_unresolved += 1
            result.unresolved_samples.append(item["text"])

        address_values = {"Đường": duong, "Từ": tu, "Đến": den}
        for vt in VITRI_ORDER:
            gia_thap, gia_cao, gia_bq = item["prices"][vt]
            for label, value in zip(MUC_GIA_LABELS, (gia_thap, gia_cao, gia_bq)):
                if value is None:
                    continue
                row = {"STT": stt}
                for col in column_order:
                    row[col] = address_values[col]
                row.update({
                    "Vị trí": vt,
                    "Mức giá": label,
                    "Đơn giá (đồng)": value,
                    "Dòng nguồn": item["row"],
                })
                output_rows.append(row)
                stt += 1

        if progress_cb:
            progress_cb(i, total, f"Đang xử lý dòng gốc {item['row']} ({i}/{total})")

    result.df = pd.DataFrame(output_rows)
    return result


def export_to_excel(df: pd.DataFrame, output_path: str, sheet_name: str = "Đối chiếu giá"):
    df.to_excel(output_path, index=False, sheet_name=sheet_name)
