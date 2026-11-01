"""Test việc chèn 3 cột cạnh cột gốc: ô gộp, công thức, độ rộng, bộ lọc, phương án dự phòng."""

import openpyxl
import pytest
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from core.generic import shift_formula, split_column_in_workbook

TEXT = "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện"


@pytest.fixture
def book(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng giá"
    ws.append(["STT", "Tên đường phố", "Xã cũ", "Giá thấp", "Giá cao", "Tỷ lệ", "Ghi chú", None])
    ws.append([None] * 8)
    ws.append([None, TEXT, "TT", 100, 120, "=E3/D3", "=A3&B3", None])
    ws.append([None, "Đường Đìa Muỗi", "TT", 50, 60, "=E4/D4", None, None])
    ws.merge_cells("B1:B2")        # tiêu đề cột gốc gộp 2 dòng
    ws.merge_cells("G1:H1")        # ô gộp nằm bên phải cột gốc
    ws.column_dimensions["D"].width = 17
    ws.auto_filter.ref = "A2:G4"
    ws.freeze_panes = "C3"
    path = tmp_path / "in.xlsx"
    wb.save(path)
    return str(path)


def run(book, tmp_path, **kw):
    out = str(tmp_path / "out.xlsx")
    res = split_column_in_workbook(book, out, **kw)
    return res, openpyxl.load_workbook(out)["Bảng giá"]


def test_3_cot_nam_ngay_sau_cot_goc_va_cot_cu_dich_sang_phai(book, tmp_path):
    res, ws = run(book, tmp_path)
    assert res.placement == "after-source" and res.first_new_column == "C"
    assert [ws.cell(1, c).value for c in (2, 3, 4, 5, 6)] == [
        "Tên đường phố", "Đường (tách)", "Từ (tách)", "Đến (tách)", "Xã cũ"]
    assert [ws.cell(3, c).value for c in (3, 4, 5)] == ["Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện"]
    assert ws.cell(3, 2).value == TEXT                 # cột gốc còn nguyên
    assert ws.cell(3, 7).value == 100                  # Giá thấp: D -> G


def test_cong_thuc_duoc_sua_tham_chieu(book, tmp_path):
    _, ws = run(book, tmp_path)
    assert ws.cell(3, 9).value == "=H3/G3"             # =E3/D3 -> =H3/G3, ô nằm ở cột I
    assert ws.cell(3, 10).value == "=A3&B3"            # tham chiếu bên trái cột chèn giữ nguyên


def test_o_gop_do_rong_bo_loc_duoc_dich(book, tmp_path):
    _, ws = run(book, tmp_path)
    merged = {str(m) for m in ws.merged_cells.ranges}
    assert {"B1:B2", "C1:C2", "D1:D2", "E1:E2", "J1:K1"} <= merged
    assert ws.column_dimensions["G"].width == 17       # cột D cũ
    assert ws.auto_filter.ref == "A2:J4"
    assert ws.freeze_panes == "C3"


def test_an_cot_goc(book, tmp_path):
    _, ws = run(book, tmp_path, hide_source=True)
    assert ws.column_dimensions["B"].hidden is True
    assert ws.cell(3, 2).value == TEXT


def test_mo_file_ra_la_thay_sheet_vua_xu_ly(book, tmp_path):
    out = str(tmp_path / "out.xlsx")
    split_column_in_workbook(book, out)
    wb = openpyxl.load_workbook(out)
    assert wb.active.title == "Bảng giá"


def test_co_dinh_vung_cot_gom_ca_cot_chen_neu_chen_trong_vung_co_dinh(book, tmp_path):
    wb = openpyxl.load_workbook(book)
    wb["Bảng giá"].freeze_panes = "D3"                 # cố định A..C, chèn sau B nằm trong vùng đó
    wb.save(book)
    _, ws = run(book, tmp_path)
    assert ws.freeze_panes == "G3"


def test_phuong_an_cuoi_bang_khi_chon_end(book, tmp_path):
    res, ws = run(book, tmp_path, placement="end")
    assert res.placement == "end" and res.first_new_column == "I"
    assert [ws.cell(1, c).value for c in (9, 10, 11)] == ["Đường (tách)", "Từ (tách)", "Đến (tách)"]
    assert ws.cell(3, 4).value == 100                  # cột cũ không bị dịch


def test_tu_dong_ve_cuoi_bang_khi_co_dinh_dang_co_dieu_kien(book, tmp_path):
    wb = openpyxl.load_workbook(book)
    wb["Bảng giá"].conditional_formatting.add(
        "D3:D4", CellIsRule(operator="greaterThan", formula=["10"], fill=PatternFill("solid", start_color="FFFF00")))
    wb.save(book)
    res, ws = run(book, tmp_path)
    assert res.placement == "end" and "định dạng có điều kiện" in res.warning
    assert ws.cell(1, 9).value == "Đường (tách)"


def test_tu_dong_ve_cuoi_bang_khi_sheet_khac_tro_vao(book, tmp_path):
    wb = openpyxl.load_workbook(book)
    other = wb.create_sheet("Tong hop")
    other["A1"] = "='Bảng giá'!D3"
    wb.save(book)
    res, _ = run(book, tmp_path)
    assert res.placement == "end" and "Tong hop" in res.warning


@pytest.mark.parametrize("formula, expected", [
    ("=E4/D4", "=H4/G4"),
    ("=A4+B4", "=A4+B4"),
    ("=SUM(B4:E4)", "=SUM(B4:H4)"),
    ("=SUM($D$2:$D$9)", "=SUM($G$2:$G$9)"),
    ("='Bảng giá'!E4+Khac!E4", "='Bảng giá'!H4+Khac!E4"),
    ("=IF(E4>0,\"E4\",D4)", "=IF(H4>0,\"E4\",G4)"),
])
def test_shift_formula(formula, expected):
    assert shift_formula(formula, "Bảng giá", after_col=2, n=3) == expected
