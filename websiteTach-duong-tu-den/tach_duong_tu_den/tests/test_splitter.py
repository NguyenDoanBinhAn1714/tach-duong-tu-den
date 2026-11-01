"""Test cho core.splitter. Các ca dưới đây lấy từ bảng giá Thanh Oai thật,
đáp án là giá trị trong file mẫu đối chiếu do bên giao việc cung cấp."""

import pytest

from core.splitter import is_area_heading, split_row

CASES = [
    (
        'Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba giao cắt với đường bệnh viện Thanh Oai',
        ('Quốc lộ 21B', 'cầu Thạch Bích', 'ngã ba giao cắt với đường bệnh viện Thanh Oai'),
    ),
    (
        'Đường vào thôn Kim Lâm: Từ giáp Quốc lộ 21B qua nhà Văn hóa đi sông Hòa Bình (đến hết đường)',
        ('Đường vào thôn Kim Lâm', 'giáp Quốc lộ 21B qua nhà Văn hóa đi sông Hòa Bình (đến hết đường)', ''),
    ),
    (
        'Đường nối đoạn từ giáp Quốc lộ 21B đến giáp trường THCS Thanh Mai (xã Thanh Oai) (Đường trục xã Thanh Mai cũ)',
        ('Đường nối', 'giáp Quốc lộ 21B', 'giáp trường THCS Thanh Mai (xã Thanh Oai) (Đường trục xã Thanh Mai cũ)'),
    ),
    (
        'Đường nối đoạn từ đê sông Đáy đến ngã ba thôn Hoạch An (xã Thanh Oai): Đường trục xã Kim An cũ',
        ('Đường nối', 'đê sông Đáy', 'ngã ba thôn Hoạch An (xã Thanh Oai): Đường trục xã Kim An cũ'),
    ),
    (
        'Đường nối đoạn từ giáp đường cienco 5 đến đường dẫn qua đình làng Văn Quán (Đường Đỗ Động - Thanh Văn: Đoạn giáp đường cienco 5 đến hết đường Kim Bài - Đỗ Động)',
        ('Đường nối đoạn từ giáp đường cienco 5 đến đường dẫn qua đình làng Văn Quán (Đường Đỗ Động - Thanh Văn', 'giáp đường cienco 5', 'hết đường Kim Bài - Đỗ Động'),
    ),
    (
        'Quốc lộ 21B (đoạn từ giáp phường Phú Lương đến cầu Thạch Bích)',
        ('Quốc lộ 21B', 'giáp phường Phú Lương', 'cầu Thạch Bích'),
    ),
    (
        'Đường Bích Hòa - Cự Khê (từ QL 21B đến giáp dự án đường vành đai 4)',
        ('Đường Bích Hòa - Cự Khê', 'QL 21B', 'giáp dự án đường vành đai 4'),
    ),
    (
        'Đường nối đoạn từ giáp Quốc lộ 21B đến giáp tỉnh lộ 427 (Qua địa phận các xã: Dân Hoà, Tam Hưng) (Đường Vác - Thanh Văn: Dân Hòa; Tân Ước, Thanh Văn, Thanh Thùy)',
        ('Đường nối', 'giáp Quốc lộ 21B', 'giáp tỉnh lộ 427 (Qua địa phận các xã: Dân Hoà, Tam Hưng) (Đường Vác - Thanh Văn: Dân Hòa; Tân Ước, Thanh Văn, Thanh Thùy)'),
    ),
    (
        'Đường Đìa Muỗi (đoạn qua xã Thanh Oai)',
        ('Đường Đìa Muỗi (đoạn qua xã Thanh Oai)', '', ''),
    ),
    (
        'Khu đấu giá Sau Trại',
        ('Khu đấu giá Sau Trại', '', ''),
    ),
    (
        'Thanh Thùy cũ',
        ('Thanh Thùy cũ', '', ''),
    ),
]


@pytest.mark.parametrize("text, expected", CASES)
def test_cau_thuc_te(text, expected):
    assert split_row(text)[:3] == expected


def test_nhan_dien_ten_khu_vuc_chung():
    assert is_area_heading("Khu Đấu giá Cửa Chùa, thôn Cao Mật Hạ")
    assert is_area_heading("Xã Thanh Oai")
    assert not is_area_heading("Quốc lộ 21B")


def test_nhan_phu_ke_thua_ten_khu_vuc():
    duong, tu, den, nguon = split_row("Trục chính", "Khu đất dịch vụ Cây Sữa")
    assert (duong, tu, den) == ("Khu đất dịch vụ Cây Sữa - Trục chính", "", "")
    assert nguon == "rule-parent"


def test_nhan_phu_khong_co_khu_vuc_thi_giu_nguyen():
    assert split_row("Trục chính")[:3] == ("Trục chính", "", "")


def test_cau_hai_cum_tu_den_dung_dap_an_da_doi_chieu():
    text, expected = CASES[4]
    assert split_row(text)[3] == "verified-override"


def test_khoang_trang_thua_khong_lam_hong_dap_an_doi_chieu():
    text, expected = CASES[4]
    assert split_row("  " + text.replace(" ", "  ") + " ")[:3] == expected


def test_cau_khong_khop_quy_tac_duoc_dua_cho_ai_va_giu_nguyen_neu_ai_tra_none():
    called = []
    split_row("Khu đấu giá Sau Trại", ai_fallback=lambda t: called.append(t))
    assert called == ["Khu đấu giá Sau Trại"]
    assert split_row("Khu đấu giá Sau Trại", ai_fallback=lambda t: None)[:4] == (
        "Khu đấu giá Sau Trại", "", "", "unresolved")


def test_ai_duoc_dung_khi_quy_tac_khong_tach_duoc():
    r = split_row("Khu đấu giá Sau Trại", ai_fallback=lambda t: ("A", "B", "C"))
    assert r == ("A", "B", "C", "ai")


def test_toi_duoc_xem_nhu_den():
    assert split_row("Quốc lộ 21B: Đoạn từ cầu Thạch Bích tới ngã ba bệnh viện")[:3] == (
        "Quốc lộ 21B", "cầu Thạch Bích", "ngã ba bệnh viện")
    assert split_row("Đường Bích Hòa (từ QL 21B tới giáp vành đai 4)")[:3] == (
        "Đường Bích Hòa", "QL 21B", "giáp vành đai 4")
