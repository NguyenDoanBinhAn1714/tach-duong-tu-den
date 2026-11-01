"""Chế độ dùng cho file Excel bất kỳ: tách một cột tên đường rồi chèn 3 cột
Đường / Từ / Đến ngay bên phải cột đó, giữ nguyên dữ liệu gốc.

Khi sheet có những thứ mà việc chèn cột có thể làm hỏng (định dạng có điều kiện,
kiểm tra dữ liệu, bảng Excel, công thức từ sheet khác trỏ vào...), tool không
chèn mà ghi 3 cột vào cuối bảng và báo lý do trong kết quả.
"""

import copy
import re
import unicodedata
from dataclasses import dataclass, field

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.dimensions import ColumnDimension

from .splitter import GENERIC_LABELS, is_area_heading, split_row

DEFAULT_KEYWORD = "Tên đường"
MAX_HEADER_SCAN_ROWS = 10
NEW_HEADER_SUFFIX = " (tách)"
MAX_PREVIEW_ROWS = 500
NEW_COLUMN_WIDTHS = {"Đường": 30, "Từ": 30, "Đến": 36}
HEADER_FILL = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
BODY_FILL = PatternFill("solid", start_color="F3F8FD", end_color="F3F8FD")


@dataclass
class GenericResult:
    output_path: str
    sheet_name: str
    column_letter: str
    header_row: int
    placement: str = "after-source"  # "after-source" hoặc "end"
    first_new_column: str = ""
    warning: str = ""
    processed_rows: int = 0
    skipped_rows: int = 0
    count_rule: int = 0
    count_rule_parent: int = 0
    count_verified_override: int = 0
    count_ai: int = 0
    count_unresolved: int = 0
    unresolved_samples: list = field(default_factory=list)
    rows: list = field(default_factory=list)  # tối đa MAX_PREVIEW_ROWS dòng để xem trước kết quả

    def as_dict(self):
        return {
            "sheet": self.sheet_name,
            "column": self.column_letter,
            "header_row": self.header_row,
            "placement": self.placement,
            "first_new_column": self.first_new_column,
            "warning": self.warning,
            "processed_rows": self.processed_rows,
            "skipped_rows": self.skipped_rows,
            "rule": self.count_rule,
            "rule_parent": self.count_rule_parent,
            "verified_override": self.count_verified_override,
            "ai": self.count_ai,
            "unresolved": self.count_unresolved,
        }


def _norm(s) -> str:
    return unicodedata.normalize("NFC", str(s)).strip().lower()


def list_sheets(path: str):
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def find_header(ws, keyword: str = DEFAULT_KEYWORD):
    """Tìm ô tiêu đề chứa keyword trong vài dòng đầu. Trả về (cột, dòng) hoặc None."""
    kw = _norm(keyword)
    for r in range(1, min(ws.max_row, MAX_HEADER_SCAN_ROWS) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and kw in _norm(v):
                return c, r
    return None


def resolve_column(ws, spec: str):
    """spec là chữ cái cột ("B") hoặc từ khoá trong tiêu đề ("Tên đường")."""
    spec = (spec or "").strip()
    if re.fullmatch(r"[A-Za-z]{1,3}", spec):
        col = column_index_from_string(spec.upper())
        header_row = 1
        for r in range(1, min(ws.max_row, MAX_HEADER_SCAN_ROWS) + 1):
            if isinstance(ws.cell(r, col).value, str) and ws.cell(r, col).value.strip():
                header_row = r
                break
        return col, header_row

    found = find_header(ws, spec or DEFAULT_KEYWORD)
    if not found:
        raise ValueError(
            f'Không tìm thấy cột có tiêu đề chứa "{spec or DEFAULT_KEYWORD}" '
            f"trong {MAX_HEADER_SCAN_ROWS} dòng đầu của sheet {ws.title!r}."
        )
    return found


def pick_sheet(wb, sheet_name: str = None, keyword: str = DEFAULT_KEYWORD):
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Không có sheet {sheet_name!r}. Các sheet hiện có: {wb.sheetnames}")
        return wb[sheet_name]
    # Nhiều sheet cùng có cột tên đường thì lấy sheet dài nhất (bảng chính).
    candidates = [ws for ws in wb.worksheets if find_header(ws, keyword)]
    if candidates:
        return max(candidates, key=lambda w: w.max_row)
    return wb.worksheets[0]


def _is_number_like(v) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    return isinstance(v, str) and v.startswith("=")  # ô công thức


def _looks_like_index(text: str) -> bool:
    return bool(re.fullmatch(r"\(?\d+\)?\.?", text))


# ----------------------------------------------------------------------------
# Chèn cột: dịch chuyển công thức, ô gộp, độ rộng cột, bộ lọc, ô cố định
# ----------------------------------------------------------------------------
_CELL_PART = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$")


def _shift_ref_part(part: str, after_col: int, n: int) -> str:
    m = _CELL_PART.match(part)
    if not m:
        return part
    d1, letters, d2, row = m.groups()
    col = column_index_from_string(letters.upper())
    if col > after_col:
        col += n
    return f"{d1}{get_column_letter(col)}{d2}{row}"


def _shift_range_token(value: str, sheet_title: str, after_col: int, n: int) -> str:
    prefix, ref = "", value
    if "!" in value:
        sheet, ref = value.rsplit("!", 1)
        if sheet.strip("'").replace("''", "'") != sheet_title:
            return value  # trỏ sang sheet khác thì giữ nguyên
        prefix = sheet + "!"
    return prefix + ":".join(_shift_ref_part(p, after_col, n) for p in ref.split(":"))


def shift_formula(formula: str, sheet_title: str, after_col: int, n: int) -> str:
    """Sửa các tham chiếu ô trong công thức khi chèn n cột sau cột after_col."""
    tok = Tokenizer(formula)
    parts = []
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            parts.append(_shift_range_token(t.value, sheet_title, after_col, n))
        else:
            parts.append(t.value)
    return "=" + "".join(parts)


def _insert_blockers(wb, ws) -> str:
    """Trả về lý do không nên chèn cột (chuỗi rỗng nếu chèn được an toàn)."""
    if len(wb.defined_names):
        return "workbook có tên vùng (defined names)"
    if ws.conditional_formatting and len(list(ws.conditional_formatting)):
        return "sheet có định dạng có điều kiện"
    if ws.data_validations and ws.data_validations.dataValidation:
        return "sheet có kiểm tra dữ liệu (data validation)"
    if ws._hyperlinks or ws.tables or ws._images or ws._charts:
        return "sheet có liên kết, bảng Excel, ảnh hoặc biểu đồ"
    if ws.auto_filter.filterColumn:
        return "bộ lọc đang áp dụng điều kiện"
    for row in ws.iter_rows():
        for c in row:
            if not isinstance(c.value, str) and c.value is not None and type(c.value).__name__ in (
                "ArrayFormula", "DataTableFormula"
            ):
                return "sheet có công thức mảng"
    needle = ws.title.lower()
    for other in wb.worksheets:
        if other is ws:
            continue
        for row in other.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and needle in c.value.lower():
                    return f"sheet {other.title!r} có công thức trỏ vào sheet này"
    return ""


def _shift_column_dimensions(ws, after_col: int, n: int):
    per_col = {}
    for key, dim in list(ws.column_dimensions.items()):
        lo = dim.min or column_index_from_string(key)
        hi = dim.max or lo
        for c in range(lo, hi + 1):
            per_col[c] = dim
    ws.column_dimensions.clear()
    for c, dim in sorted(per_col.items()):
        nc = c + n if c > after_col else c
        letter = get_column_letter(nc)
        new = ColumnDimension(
            ws, index=letter, width=dim.width, bestFit=dim.bestFit, hidden=dim.hidden,
            outlineLevel=dim.outlineLevel, collapsed=dim.collapsed, min=nc, max=nc,
        )
        new._style = copy.copy(dim._style)
        ws.column_dimensions[letter] = new


def _insert_columns(ws, after_col: int, n: int):
    """Chèn n cột trống ngay sau cột after_col, kéo theo mọi thứ nằm bên phải."""
    # 1. Ô gộp: gỡ ra, chèn xong gộp lại ở vị trí mới.
    merges = [(r.min_row, r.min_col, r.max_row, r.max_col) for r in list(ws.merged_cells.ranges)]
    for r in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(r))

    # 2. Công thức cùng sheet: sửa tham chiếu trước khi dịch ô.
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = shift_formula(c.value, ws.title, after_col, n)

    # 3. Dịch ô (giá trị + định dạng) sang phải.
    ws.insert_cols(after_col + 1, n)

    # 4. Ô gộp, độ rộng cột, bộ lọc, ô cố định.
    for r1, c1, r2, c2 in merges:
        nc1 = c1 + n if c1 > after_col else c1
        nc2 = c2 + n if c2 > after_col else c2
        ws.merge_cells(start_row=r1, start_column=nc1, end_row=r2, end_column=nc2)
    _shift_column_dimensions(ws, after_col, n)

    if ws.auto_filter.ref:
        rng = CellRange(ws.auto_filter.ref)
        lo = rng.min_col + n if rng.min_col > after_col else rng.min_col
        hi = rng.max_col + n if rng.max_col > after_col else rng.max_col
        ws.auto_filter.ref = f"{get_column_letter(lo)}{rng.min_row}:{get_column_letter(hi)}{rng.max_row}"

    if ws.freeze_panes:
        fc = ws[ws.freeze_panes]
        # Cột bị cố định là A..(fc.column-1). Chèn nằm trong vùng đó thì mở rộng vùng cố định.
        if after_col < fc.column - 1:
            ws.freeze_panes = f"{get_column_letter(fc.column + n)}{fc.row}"


def _header_merge_rows(ws, header_row: int, col: int):
    """Nếu ô tiêu đề cột gốc gộp theo chiều dọc (vd B1:B2) thì trả về (dòng đầu, dòng cuối)."""
    for r in ws.merged_cells.ranges:
        if r.min_col == r.max_col == col and r.min_row <= header_row <= r.max_row and r.max_row > r.min_row:
            return r.min_row, r.max_row
    return None


def split_column_in_workbook(
    input_path: str,
    output_path: str,
    sheet_name: str = None,
    column_spec: str = DEFAULT_KEYWORD,
    ai_fallback=None,
    column_order=("Đường", "Từ", "Đến"),
    only_data_rows: bool = True,
    placement: str = "after-source",
    hide_source: bool = False,
    progress_cb=None,
) -> GenericResult:
    """Đọc file, tách cột tên đường, ghi 3 cột mới rồi lưu ra output_path.

    placement: "after-source" chèn 3 cột ngay bên phải cột gốc (mặc định),
               "end" ghi vào cuối bảng.
    hide_source: ẩn cột gốc sau khi tách (dữ liệu vẫn còn, bỏ ẩn là thấy lại).
    only_data_rows: chỉ tách những dòng có số liệu ở các cột khác (bỏ qua
               dòng tiêu đề nhóm như tên xã, tên khu vực).
    """
    wb = openpyxl.load_workbook(input_path)
    default_kw = column_spec if not re.fullmatch(r"[A-Za-z]{1,3}", column_spec or "") else DEFAULT_KEYWORD
    ws = pick_sheet(wb, sheet_name, default_kw)
    col, header_row = resolve_column(ws, column_spec)

    result = GenericResult(
        output_path=output_path, sheet_name=ws.title,
        column_letter=get_column_letter(col), header_row=header_row,
    )

    # ---- Bước 1: tách xong toàn bộ trước, chưa đụng vào cấu trúc sheet ----
    max_col = ws.max_column
    skip_cols = {1, col}
    outputs = {}
    last_area_name = None
    total_rows = ws.max_row - header_row
    for i, r in enumerate(range(header_row + 1, ws.max_row + 1), start=1):
        raw = ws.cell(r, col).value
        text = str(raw).strip() if raw is not None else ""
        if progress_cb:
            progress_cb(i, total_rows, f"Dòng {r}")
        if not text or _looks_like_index(text):
            continue

        if is_area_heading(text):
            last_area_name = text
            result.skipped_rows += 1
            continue

        # Các tên nhóm bắt đầu bằng khu/xã/thôn được ghi nhớ nhưng không tạo tuyến riêng.
        # Nhãn Trục chính/Đường làn trong kế thừa tên nhóm gần nhất.

        if only_data_rows and not any(
            _is_number_like(ws.cell(r, c).value) for c in range(1, max_col + 1) if c not in skip_cols
        ):
            result.skipped_rows += 1
            continue

        duong, tu, den, source = split_row(text, last_area_name, ai_fallback=ai_fallback)
        outputs[r] = {"Đường": duong, "Từ": tu, "Đến": den}
        result.processed_rows += 1
        if len(result.rows) < MAX_PREVIEW_ROWS:
            result.rows.append({"row": r, "text": text, "duong": duong, "tu": tu, "den": den, "source": source})
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
            result.unresolved_samples.append(text)

    # ---- Bước 2: tạo chỗ cho 3 cột mới ----
    n = len(column_order)
    header_span = _header_merge_rows(ws, header_row, col)
    if placement == "after-source":
        blocker = _insert_blockers(wb, ws)
        if blocker:
            placement = "end"
            result.warning = f"Không chèn cột cạnh cột gốc được ({blocker}); đã ghi 3 cột vào cuối bảng."
    if placement == "after-source":
        _insert_columns(ws, col, n)
        first_new = col + 1
    else:
        first_new = max_col + 1
    result.placement = placement
    result.first_new_column = get_column_letter(first_new)
    new_cols = {name: first_new + i for i, name in enumerate(column_order)}

    # ---- Bước 3: ghi tiêu đề + dữ liệu, định dạng theo cột gốc ----
    last_row = ws.max_row
    for name, c in new_cols.items():
        for r in range(1, last_row + 1):
            src = ws.cell(r, col)
            dst = ws.cell(r, c)
            if getattr(src, "has_style", False):
                dst._style = copy.copy(src._style)

        head = ws.cell(header_row, c, name + NEW_HEADER_SUFFIX)
        head.fill = HEADER_FILL
        head.font = Font(name=head.font.name, size=head.font.size, bold=True)
        head.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if header_span:
            ws.merge_cells(start_row=header_span[0], start_column=c, end_row=header_span[1], end_column=c)

        for r, values in outputs.items():
            cell = ws.cell(r, c)
            cell.value = values[name] or None
            cell.fill = BODY_FILL

        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = NEW_COLUMN_WIDTHS.get(name, 30)

    if hide_source:
        ws.column_dimensions[get_column_letter(col)].hidden = True

    # Mở file ra là thấy ngay sheet vừa xử lý (và chỉ sheet đó được chọn).
    wb.active = wb.worksheets.index(ws)
    for s in wb.worksheets:
        s.sheet_view.tabSelected = s is ws

    wb.save(output_path)
    return result
