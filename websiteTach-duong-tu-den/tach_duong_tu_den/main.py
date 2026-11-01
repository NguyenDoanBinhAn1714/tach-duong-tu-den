"""Bản máy tính (desktop) của công cụ tách Tên đường phố -> Đường / Từ / Đến.

Chạy: python main.py
Bản chạy trên server (trình duyệt): xem run_server.py
"""

import os
import shutil
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from core.ai_fallback import GeminiFallback
from core.generic import list_sheets, split_column_in_workbook
from core.processor import export_to_excel, process_file
from core.splitter import split_row

try:
    from config import GEMINI_API_KEY
except ImportError:
    GEMINI_API_KEY = ""

NAVY_DARK = "#0B2545"
NAVY_MED = "#13315C"
OCEAN_BLUE = "#1E6091"
OCEAN_BLUE_HOVER = "#2E86AB"
TEXT_LIGHT = "#E8F1F2"
TEXT_MUTED = "#9FB3C8"
LOG_BG = "#EAF4FC"
LOG_FG = "#0B2545"

COLUMN_ORDER_OPTIONS = {
    "Đường → Từ → Đến": ("Đường", "Từ", "Đến"),
    "Từ → Đến → Đường": ("Từ", "Đến", "Đường"),
}
MODE_OPTIONS = {
    "Thêm 3 cột vào file gốc (mọi file Excel)": "add-columns",
    "Bảng đối chiếu giá (đúng file mẫu Thanh Oai)": "price-table",
}
SOURCE_LABELS = {
    "rule": "quy tắc", "rule-parent": "kế thừa khu vực", "verified-override": "đáp án đã đối chiếu",
    "ai": "AI (Gemini)", "unresolved": "không có mô tả đoạn, giữ nguyên",
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tách Đường / Từ / Đến - Bảng giá đất Thanh Oai")
        self.geometry("900x840")
        self.configure(bg=NAVY_DARK)
        self.minsize(820, 720)

        self.input_path = tk.StringVar()
        self.api_key = tk.StringVar(value=GEMINI_API_KEY)
        self.use_ai = tk.BooleanVar(value=False)
        self.hide_source = tk.BooleanVar(value=False)
        self.mode_label = tk.StringVar(value=list(MODE_OPTIONS.keys())[0])
        self.column_order_label = tk.StringVar(value=list(COLUMN_ORDER_OPTIONS.keys())[0])
        self.sheet_name = tk.StringVar()
        self.column_spec = tk.StringVar(value="Tên đường")
        self.demo_text = tk.StringVar()

        self.result_df = None
        self.result_tmp_path = None
        self.result_kind = None

        self._build_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ style
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=NAVY_DARK)
        style.configure("Card.TFrame", background=NAVY_MED)
        style.configure("TLabel", background=NAVY_DARK, foreground=TEXT_LIGHT, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=NAVY_DARK, foreground=TEXT_MUTED, font=("Segoe UI", 9))
        style.configure("Card.TLabel", background=NAVY_MED, foreground=TEXT_LIGHT, font=("Segoe UI", 10))
        style.configure("CardMuted.TLabel", background=NAVY_MED, foreground=TEXT_MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=NAVY_DARK, foreground=TEXT_LIGHT, font=("Segoe UI", 16, "bold"))
        style.configure("Section.TLabel", background=NAVY_DARK, foreground=TEXT_LIGHT, font=("Segoe UI", 12, "bold"))
        style.configure("Stat.TLabel", background=NAVY_MED, foreground=TEXT_LIGHT, font=("Segoe UI", 10, "bold"))
        style.configure("Result.TLabel", background=NAVY_MED, foreground="#7FD1AE", font=("Segoe UI", 10, "bold"))
        style.configure("Card.TCheckbutton", background=NAVY_MED, foreground=TEXT_LIGHT, font=("Segoe UI", 10))
        style.map("Card.TCheckbutton", background=[("active", NAVY_MED)])
        style.configure(
            "Ocean.TButton", background=OCEAN_BLUE, foreground="white",
            font=("Segoe UI", 10, "bold"), padding=8, borderwidth=0,
        )
        style.map("Ocean.TButton", background=[("active", OCEAN_BLUE_HOVER)])
        style.configure(
            "Ocean.Horizontal.TProgressbar", troughcolor=NAVY_MED, background=OCEAN_BLUE,
            bordercolor=NAVY_DARK, lightcolor=OCEAN_BLUE, darkcolor=OCEAN_BLUE,
        )
        style.configure("TCombobox", fieldbackground="white", background="white")

    # --------------------------------------------------------------- ui
    def _build_ui(self):
        header = ttk.Frame(self, style="TFrame")
        header.pack(fill="x", padx=16, pady=8)
        ttk.Label(header, text="Tách cột Tên đường phố → Đường / Từ / Đến", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Quy tắc regex xử lý phần lớn dữ liệu, AI (Gemini) chỉ dùng cho câu quá phức tạp và chỉ khi bạn bật.",
            style="Muted.TLabel",
        ).pack(anchor="w")

        self._build_demo_panel()
        self._build_file_panel()
        self._build_action_panel()
        self._build_log_panel()
        self._mode_changed()

    def _entry(self, parent, var, **kw):
        return tk.Entry(parent, textvariable=var, font=("Segoe UI", 10), bg="white", fg=NAVY_DARK,
                        relief="flat", disabledbackground="#C9D6E3", disabledforeground="#6B7C93", **kw)

    def _build_demo_panel(self):
        ttk.Label(self, text="Thử nhanh một câu", style="Section.TLabel").pack(anchor="w", padx=16, pady=(4, 0))

        card = ttk.Frame(self, style="Card.TFrame")
        card.pack(fill="x", padx=16, pady=(4, 8))
        card.columnconfigure(0, weight=1)

        entry = self._entry(card, self.demo_text)
        entry.grid(row=0, column=0, sticky="we", padx=(12, 8), pady=12, ipady=4)
        entry.insert(0, "Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba giao cắt với đường bệnh viện Thanh Oai")
        ttk.Button(card, text="Tách thử", style="Ocean.TButton", command=self._demo_clicked).grid(
            row=0, column=1, padx=(0, 12), pady=12
        )

        self.demo_result_var = tk.StringVar(value='Nhập một câu ở trên rồi bấm "Tách thử" để xem kết quả.')
        ttk.Label(card, textvariable=self.demo_result_var, style="Result.TLabel", wraplength=820).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 12)
        )

    def _build_file_panel(self):
        ttk.Label(self, text="Xử lý cả file", style="Section.TLabel").pack(anchor="w", padx=16, pady=(4, 0))

        card = ttk.Frame(self, style="Card.TFrame")
        card.pack(fill="x", padx=16, pady=(4, 8))
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        def label(text, row, col, muted=False):
            ttk.Label(card, text=text, style="CardMuted.TLabel" if muted else "Card.TLabel").grid(
                row=row, column=col, sticky="w", padx=12, pady=(10, 3)
            )

        label("File Excel gốc:", 0, 0)
        path_frame = ttk.Frame(card, style="Card.TFrame")
        path_frame.grid(row=1, column=0, columnspan=2, sticky="we", padx=12)
        self._entry(path_frame, self.input_path).pack(side="left", fill="x", expand=True, ipady=4)
        ttk.Button(path_frame, text="Chọn file...", style="Ocean.TButton", command=self._browse_input).pack(
            side="left", padx=(8, 0)
        )

        label("Chế độ xử lý:", 2, 0)
        label("Thứ tự cột địa chỉ khi xuất file:", 2, 1)
        self.mode_combo = ttk.Combobox(
            card, textvariable=self.mode_label, values=list(MODE_OPTIONS.keys()),
            state="readonly", font=("Segoe UI", 10),
        )
        self.mode_combo.grid(row=3, column=0, sticky="we", padx=12)
        self.mode_combo.bind("<<ComboboxSelected>>", lambda _e: self._mode_changed())
        ttk.Combobox(
            card, textvariable=self.column_order_label, values=list(COLUMN_ORDER_OPTIONS.keys()),
            state="readonly", font=("Segoe UI", 10),
        ).grid(row=3, column=1, sticky="we", padx=12)

        label("Sheet (để trống = tự chọn):", 4, 0)
        label("Cột chứa tên đường (tiêu đề hoặc chữ cái cột):", 4, 1)
        self.sheet_combo = ttk.Combobox(card, textvariable=self.sheet_name, values=[], font=("Segoe UI", 10))
        self.sheet_combo.grid(row=5, column=0, sticky="we", padx=12)
        self.column_entry = self._entry(card, self.column_spec)
        self.column_entry.grid(row=5, column=1, sticky="we", padx=12, ipady=3)

        label("API key Gemini (tuỳ chọn):", 6, 0)
        key_frame = ttk.Frame(card, style="Card.TFrame")
        key_frame.grid(row=7, column=0, columnspan=2, sticky="we", padx=12)
        self.entry_key = self._entry(key_frame, self.api_key, show="•")
        self.entry_key.pack(side="left", fill="x", expand=True, ipady=4)
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(key_frame, text="Hiện", variable=self.show_key_var, style="Card.TCheckbutton",
                        command=self._toggle_key_visibility).pack(side="left", padx=(8, 0))

        ai_row = ttk.Frame(card, style="Card.TFrame")
        ai_row.grid(row=8, column=0, columnspan=2, sticky="we", padx=12, pady=(8, 4))
        ttk.Checkbutton(ai_row, text="Dùng AI (Gemini) cho câu quy tắc không tách được",
                        variable=self.use_ai, style="Card.TCheckbutton").pack(side="left")
        hint = "key đã nạp sẵn từ config.py" if GEMINI_API_KEY else "chưa có key — cần dán key mới dùng được AI"
        ttk.Label(ai_row, text=f"({hint})", style="CardMuted.TLabel").pack(side="left", padx=(8, 0))

        src_row = ttk.Frame(card, style="Card.TFrame")
        src_row.grid(row=9, column=0, columnspan=2, sticky="we", padx=12, pady=(0, 12))
        self.hide_source_check = ttk.Checkbutton(
            src_row, text="Ẩn cột Tên đường phố gốc sau khi tách (dữ liệu vẫn còn, bỏ ẩn là thấy lại)",
            variable=self.hide_source, style="Card.TCheckbutton")
        self.hide_source_check.pack(side="left")

    def _build_action_panel(self):
        action_frame = ttk.Frame(self, style="TFrame")
        action_frame.pack(fill="x", padx=16, pady=8)
        self.btn_run = ttk.Button(action_frame, text="▶  Bắt đầu xử lý", style="Ocean.TButton", command=self._run_clicked)
        self.btn_run.pack(side="left")
        self.btn_save = ttk.Button(
            action_frame, text="💾  Lưu file kết quả...", style="Ocean.TButton",
            command=self._save_clicked, state="disabled",
        )
        self.btn_save.pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(self, style="Ocean.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", padx=16, pady=(0, 8))

        stat_card = ttk.Frame(self, style="Card.TFrame")
        stat_card.pack(fill="x", padx=16, pady=(0, 8))
        self.stat_label = ttk.Label(stat_card, text="Chưa xử lý — chọn file và bấm Bắt đầu.", style="Stat.TLabel",
                                    wraplength=840)
        self.stat_label.pack(anchor="w", padx=12, pady=10)

    def _build_log_panel(self):
        ttk.Label(self, text="Nhật ký xử lý:", style="TLabel").pack(anchor="w", padx=16)
        self.log_box = scrolledtext.ScrolledText(
            self, height=8, bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9), relief="flat"
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(4, 16))
        self.log_box.configure(state="disabled")

    # ----------------------------------------------------------- hành vi
    def _mode_changed(self):
        price = MODE_OPTIONS[self.mode_label.get()] == "price-table"
        self.column_entry.configure(state="disabled" if price else "normal")
        self.hide_source_check.configure(state="disabled" if price else "normal")

    def _toggle_key_visibility(self):
        self.entry_key.configure(show="" if self.show_key_var.get() else "•")

    def _make_ai(self):
        """Trả về bộ gọi Gemini nếu người dùng bật AI và có key, ngược lại None."""
        if not self.use_ai.get():
            return None
        key = self.api_key.get().strip()
        if not key:
            self.log("Đã bật AI nhưng chưa có API key, bỏ qua bước gọi AI.")
            return None
        ai = GeminiFallback(key, log=self.log)
        return ai if ai.available() else None

    def _demo_clicked(self):
        text = self.demo_text.get().strip()
        if not text:
            self.demo_result_var.set("Bạn chưa nhập câu nào.")
            return

        def worker():
            ai = self._make_ai()
            duong, tu, den, source = split_row(text, last_area_name=None, ai_fallback=ai)
            msg = (f"Đường: {duong or '—'}   |   Từ: {tu or '—'}   |   Đến: {den or '—'}   "
                   f"(xử lý bằng: {SOURCE_LABELS.get(source, source)})")
            if source == "unresolved":
                if ai is None:
                    msg += ('\nQuy tắc không thấy cụm "từ … đến …" nên giữ nguyên câu. '
                            'Tick "Dùng AI" và dán API key Gemini để AI thử tách.')
                else:
                    msg += "\nAI cũng không trả được kết quả (xem nhật ký), câu được giữ nguyên."
            self.demo_result_var.set(msg)

        threading.Thread(target=worker, daemon=True).start()

    def _browse_input(self):
        path = filedialog.askopenfilename(title="Chọn file Excel gốc", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.input_path.set(path)
        try:
            sheets = list_sheets(path)
        except Exception as e:
            self.log(f"Không đọc được danh sách sheet: {e}")
            return
        self.sheet_combo.configure(values=sheets)
        self.sheet_name.set("Bảng giá" if "Bảng giá" in sheets else "")

    def log(self, message: str):
        def _append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.after(0, _append)

    def _set_progress(self, done, total, message):
        def _update():
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
        self.after(0, _update)
        if done % 25 == 0 or done == total:
            self.log(f"  {message} ({done}/{total})")

    def _discard_tmp(self):
        if self.result_tmp_path and os.path.exists(self.result_tmp_path):
            try:
                os.remove(self.result_tmp_path)
            except OSError:
                pass
        self.result_tmp_path = None

    def _run_clicked(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("Thiếu file", "Bạn chưa chọn file Excel gốc.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("Không tìm thấy file", f"Không tìm thấy: {path}")
            return

        self.btn_run.configure(state="disabled")
        self.btn_save.configure(state="disabled")
        self._discard_tmp()
        self.result_df = None
        self.result_kind = None
        self.stat_label.configure(text="Đang xử lý...")
        self.progress["value"] = 0
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        threading.Thread(target=self._process_worker, args=(path,), daemon=True).start()

    def _process_worker(self, path):
        ai = self._make_ai()
        self.log("Có dùng AI cho câu quy tắc không tách được." if ai else "Chỉ dùng quy tắc (regex), không gọi AI.")

        mode = MODE_OPTIONS[self.mode_label.get()]
        order = COLUMN_ORDER_OPTIONS[self.column_order_label.get()]
        sheet = self.sheet_name.get().strip()

        try:
            if mode == "add-columns":
                fd, tmp = tempfile.mkstemp(suffix=".xlsx")
                os.close(fd)
                res = split_column_in_workbook(
                    path, tmp, sheet_name=sheet or None,
                    column_spec=self.column_spec.get().strip() or "Tên đường",
                    ai_fallback=ai, column_order=order, hide_source=self.hide_source.get(),
                    progress_cb=self._set_progress,
                )
                self.result_tmp_path = tmp
                self.result_kind = mode
                head = (f"Sheet «{res.sheet_name}», cột {res.column_letter}: tách {res.processed_rows} dòng "
                        f"(bỏ qua {res.skipped_rows} dòng tiêu đề nhóm)")
                counts = (res.count_rule, res.count_rule_parent, res.count_verified_override,
                          res.count_ai, res.count_unresolved)
                unresolved = res.unresolved_samples
                if res.placement == "after-source":
                    tail = (f"Ba cột mới (tô xanh nhạt) nằm ngay sau cột {res.column_letter}, "
                            f"từ cột {res.first_new_column}. Dữ liệu gốc giữ nguyên, các cột bên phải dịch sang phải.")
                else:
                    tail = res.warning or "Ba cột mới được ghi vào cuối bảng."
            else:
                res = process_file(path, ai_fallback=ai, progress_cb=self._set_progress,
                                   column_order=order, sheet_name=sheet or "Bảng giá")
                self.result_df = res.df
                self.result_kind = mode
                head = (f"Tổng dòng dữ liệu gốc: {res.total_source_rows}  →  "
                        f"{len(res.df)} dòng kết quả (mỗi mức giá một dòng)")
                counts = (res.count_rule, res.count_rule_parent, res.count_verified_override,
                          res.count_ai, res.count_unresolved)
                unresolved = res.unresolved_samples
                tail = ""
        except Exception as e:
            self.log(f"Lỗi khi xử lý file: {e}")
            self.after(0, lambda: messagebox.showerror("Lỗi", str(e)))
            self.after(0, lambda: self.btn_run.configure(state="normal"))
            self.after(0, lambda: self.stat_label.configure(text="Xử lý thất bại, xem nhật ký bên dưới."))
            return

        rule, parent, override, ai_n, unres = counts
        summary = (f"{head}\nQuy tắc: {rule}  |  Kế thừa khu vực: {parent}  |  Đã đối chiếu sẵn: {override}  |  "
                   f"AI: {ai_n}  |  Giữ nguyên (không có mô tả đoạn): {unres}")
        self.log("—" * 70)
        self.log(summary)
        if tail:
            self.log(tail)
        if unresolved:
            self.log("Các câu không có mô tả đoạn (giữ nguyên trong cột Đường):")
            for s in unresolved:
                self.log(f"   - {s}")

        self.after(0, lambda: self.stat_label.configure(text=summary))
        self.after(0, lambda: self.btn_run.configure(state="normal"))
        self.after(0, lambda: self.btn_save.configure(state="normal"))
        self.log("Xử lý xong. Bấm 'Lưu file kết quả...' để xuất Excel.")

    def _save_clicked(self):
        if self.result_kind is None:
            return
        if self.result_kind == "add-columns":
            initial = f"ket_qua_{Path(self.input_path.get()).stem}.xlsx"
        else:
            initial = "Doi_chieu_gia_Thanh_Oai_ket_qua.xlsx"
        path = filedialog.asksaveasfilename(
            title="Lưu file kết quả", defaultextension=".xlsx", initialfile=initial,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            if self.result_kind == "add-columns":
                shutil.copyfile(self.result_tmp_path, path)
            else:
                export_to_excel(self.result_df, path)
        except Exception as e:
            messagebox.showerror("Lỗi khi lưu file", str(e))
            return
        self.log(f"Đã lưu file kết quả tại: {path}")
        messagebox.showinfo("Hoàn tất", f"Đã lưu file kết quả:\n{path}")

    def _on_close(self):
        self._discard_tmp()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
