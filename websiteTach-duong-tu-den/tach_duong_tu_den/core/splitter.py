"""Tách một câu "Tên đường phố" thành 3 phần: Đường / Từ / Đến.

Cách làm, theo thứ tự ưu tiên:
  1. Nhãn phụ như "Trục chính" lấy tên khu vực của dòng ngay phía trên.
  2. Câu nằm trong danh sách đã đối chiếu tay (VERIFIED_OVERRIDES).
  3. Quy tắc regex: dấu ":" đứng trước thì cắt theo dấu hai chấm, ngược lại
     tìm cụm "từ X đến Y" / "giáp X đến Y" trong cả câu.
  4. Không khớp quy tắc nào thì nhờ AI (nếu có), không có AI thì giữ nguyên.
"""

import re
import unicodedata

# Nhãn vị trí phụ nằm dưới một khu đất / khu chợ, bản thân không phải tên đường.
GENERIC_LABELS = {"trục chính", "đường làn trong"}
AREA_HEADING_RE = re.compile(
    r"^(?:khu(?:\s+đấu\s+giá)?(?:\s+vực)?|xã|thôn|ấp|tổ\s+dân\s+phố)\b",
    re.IGNORECASE,
)


def is_area_heading(text: str) -> bool:
    """Nhận diện dòng chỉ là tiêu đề nhóm/khu vực, không phải tuyến đường độc lập."""
    return bool(AREA_HEADING_RE.match(normalize_text(text)))


def normalize_text(text: str) -> str:
    """Chuẩn hoá để so khớp: gộp khoảng trắng, đưa về dạng NFC."""
    return " ".join(unicodedata.normalize("NFC", str(text)).split())


# Câu có hai cụm "từ...đến" lồng nhau, regex không phân biệt được cụm nào là
# cụm chính. Đáp án đã đối chiếu với file mẫu của bên giao việc.
# Khoá là nội dung câu gốc (không phụ thuộc số dòng nên dùng được cho file khác).
_AMBIGUOUS_TEXT = (
    "Đường nối đoạn từ giáp đường cienco 5 đến đường dẫn qua đình làng Văn Quán "
    "(Đường Đỗ Động - Thanh Văn: Đoạn giáp đường cienco 5 đến hết đường Kim Bài - Đỗ Động)"
)
VERIFIED_OVERRIDES = {
    normalize_text(_AMBIGUOUS_TEXT): (
        "Đường nối đoạn từ giáp đường cienco 5 đến đường dẫn qua đình làng "
        "Văn Quán (Đường Đỗ Động - Thanh Văn",
        "giáp đường cienco 5",
        "hết đường Kim Bài - Đỗ Động",
    ),
}


def _strip_paren(prefix: str, value: str) -> str:
    """Bỏ ")" thừa ở cuối value khi prefix còn "(" chưa đóng."""
    if value.endswith(")") and prefix.count("(") > prefix.count(")"):
        return value[:-1].strip()
    return value


def _clean_duong(s: str) -> str:
    s = s.strip()
    s = re.sub(r"[:\s]+$", "", s)
    s = re.sub(r"(?i)\s*đoạn\s*$", "", s).strip()
    s = re.sub(r"\(\s*$", "", s).strip()
    return s


def _find_tu_den(s: str):
    """Tìm "từ X đến Y" hoặc "giáp X đến Y" trong cả câu."""
    m = re.search(r"^(.*?)\s*(?:đoạn\s+)?từ\s+(.+?)\s+(?:đến|tới)\s+(.+)$", s, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2).strip(), m.group(3).strip()
    m = re.search(r"^(.*?)\s*(?:đoạn\s+)?(giáp\s+.+?)\s+(?:đến|tới)\s+(.+)$", s, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2).strip(), m.group(3).strip()
    return None


def split_row(text: str, last_area_name=None, ai_fallback=None):
    """Trả về (duong, tu, den, nguon).

    nguon là một trong: "rule", "rule-parent", "verified-override", "ai", "unresolved".
    ai_fallback: hàm nhận câu gốc, trả về (duong, tu, den) hoặc None.
    """
    text = text.strip()

    if text.lower() in GENERIC_LABELS:
        if not last_area_name:
            return text, "", "", "unresolved"
        return f"{last_area_name} - {text}", "", "", "rule-parent"

    key = normalize_text(text)
    if key in VERIFIED_OVERRIDES:
        d, t, e = VERIFIED_OVERRIDES[key]
        return d, t, e, "verified-override"

    pos_colon = text.find(":")
    m_conn = re.search(r"(?:từ|giáp)\s", text, re.IGNORECASE)
    pos_conn = m_conn.start() if m_conn else 10 ** 9
    if pos_colon == -1:
        pos_colon = 10 ** 9

    if pos_colon < pos_conn:
        duong, mota = text.split(":", 1)
        duong, mota = duong.strip(), mota.strip()
        mota_clean = re.sub(r"^(đoạn\s+)", "", mota, flags=re.IGNORECASE).strip()
        if not mota_clean:
            return duong, "", "", "rule"
        mm = re.search(r"^(?:từ\s+)?(.+?)\s+(?:đến|tới)\s+(.+)$", mota_clean, re.IGNORECASE)
        if mm:
            tu = mm.group(1).strip()
            den = _strip_paren(duong, mm.group(2).strip())
            return duong, tu, den, "rule"
        mm2 = re.match(r"^từ\s+(.+)$", mota_clean, re.IGNORECASE)
        if mm2:
            return duong, mm2.group(1).strip(), "", "rule"
        return duong, mota_clean, "", "rule"

    r = _find_tu_den(text)
    if r:
        duong_raw, tu, den = r
        den = _strip_paren(duong_raw, den)
        return _clean_duong(duong_raw), tu, den, "rule"

    # Câu không có mô tả đoạn: giữ nguyên là hợp lệ, chỉ nhờ AI khi có cấu hình.
    if ai_fallback is not None:
        result = ai_fallback(text)
        if result is not None:
            d, t, e = result
            return d, t, e, "ai"

    return text, "", "", "unresolved"
