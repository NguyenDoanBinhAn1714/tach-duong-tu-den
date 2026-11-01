"""Test chạy cả pipeline trên một file Excel nhỏ dựng tại chỗ (không cần file thật)."""

import openpyxl
import pytest

from core.generic import resolve_column, split_column_in_workbook
from core.processor import process_file


@pytest.fixture
def sample_xlsx(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bảng giá"
    ws.append(["STT", "Tên đường phố", "Xã", "VT1", None, None])
    ws.append([None, None, None, "Thấp", "Cao", "Bình quân"])
    ws.append(["A", "Xã Thanh Oai", None, None, None, None])               # dòng tiêu đề nhóm
    ws.append([None, "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện", "TT", 100, 120, 110])
    ws.append([None, "Khu đất dịch vụ Cây Sữa", None, None, None, None])  # tiêu đề khu vực
    ws.append([None, "Trục chính", None, 90, 110, 100])
    ws.append([None, "Đường làn trong", None, 70, 85, 77.5])
    path = tmp_path / "mau.xlsx"
    wb.save(path)
    return str(path)


def test_resolve_column_theo_tieu_de_va_theo_chu_cai(sample_xlsx):
    ws = openpyxl.load_workbook(sample_xlsx)["Bảng giá"]
    assert resolve_column(ws, "Tên đường") == (2, 1)
    assert resolve_column(ws, "b") == (2, 1)


def test_resolve_column_bao_loi_khi_khong_thay(sample_xlsx):
    ws = openpyxl.load_workbook(sample_xlsx)["Bảng giá"]
    with pytest.raises(ValueError):
        resolve_column(ws, "Cột không tồn tại")


def test_them_3_cot_vao_file_goc(sample_xlsx, tmp_path):
    out = str(tmp_path / "ra.xlsx")
    res = split_column_in_workbook(sample_xlsx, out)
    assert (res.processed_rows, res.skipped_rows) == (3, 2)
    ws = openpyxl.load_workbook(out)["Bảng giá"]
    assert [ws.cell(1, c).value for c in (3, 4, 5)] == ["Đường (tách)", "Từ (tách)", "Đến (tách)"]
    assert [ws.cell(4, c).value for c in (3, 4, 5)] == ["Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện"]
    assert ws.cell(6, 3).value == "Khu đất dịch vụ Cây Sữa - Trục chính"
    assert ws.cell(7, 3).value == "Khu đất dịch vụ Cây Sữa - Đường làn trong"
    assert ws.cell(3, 3).value is None                                      # dòng tiêu đề nhóm để trống
    assert ws.cell(4, 2).value.startswith("Quốc lộ 21B: Đoạn từ")           # dữ liệu gốc còn nguyên


def test_khu_vuc_chung_khong_tao_ban_ghi_doc_lap_va_duong_con_ke_thua_khu_vuc(sample_xlsx, tmp_path):
    wb = openpyxl.load_workbook(sample_xlsx)
    ws = wb["Bảng giá"]
    ws.insert_rows(8)
    ws["B8"] = "Khu Đấu giá Cửa Chùa, thôn Cao Mật Hạ"
    ws["C8"] = "TT"
    ws["D8"] = 80
    ws["E8"] = 100
    ws["F8"] = 90
    ws["B9"] = "Trục chính"
    ws["D9"] = 90
    ws["E9"] = 110
    ws["F9"] = 100
    ws["B10"] = "Đường làn trong"
    ws["D10"] = 70
    ws["E10"] = 85
    ws["F10"] = 77.5
    wb.save(sample_xlsx)

    out = str(tmp_path / "ra_khu_vuc.xlsx")
    res = split_column_in_workbook(sample_xlsx, out)
    out_ws = openpyxl.load_workbook(out)["Bảng giá"]

    assert out_ws["C8"].value is None
    assert out_ws["C9"].value == "Khu Đấu giá Cửa Chùa, thôn Cao Mật Hạ - Trục chính"
    assert out_ws["C10"].value == "Khu Đấu giá Cửa Chùa, thôn Cao Mật Hạ - Đường làn trong"
    assert out_ws["D9"].value is None and out_ws["E9"].value is None
    assert res.skipped_rows >= 3


def test_doi_thu_tu_cot(sample_xlsx, tmp_path):
    out = str(tmp_path / "ra.xlsx")
    split_column_in_workbook(sample_xlsx, out, column_order=("Từ", "Đến", "Đường"))
    ws = openpyxl.load_workbook(out)["Bảng giá"]
    assert [ws.cell(1, c).value for c in (3, 4, 5)] == ["Từ (tách)", "Đến (tách)", "Đường (tách)"]


def test_bang_doi_chieu_gia_moi_muc_gia_mot_dong(sample_xlsx):
    res = process_file(sample_xlsx)
    assert res.total_source_rows == 3
    assert len(res.df) == 9                                                 # 3 dòng x 3 mức giá (VT1)
    assert list(res.df.columns) == ["STT", "Đường", "Từ", "Đến", "Vị trí", "Mức giá", "Đơn giá (đồng)", "Dòng nguồn"]
