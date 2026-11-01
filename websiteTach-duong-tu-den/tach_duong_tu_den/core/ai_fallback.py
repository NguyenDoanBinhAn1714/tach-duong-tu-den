"""
ai_fallback.py
--------------
Gọi AI (Google Gemini - bản miễn phí) để tách những dòng mà quy tắc
(regex) trong splitter.py không thể xử lý chắc chắn.

Đây là thành phần TÙY CHỌN: nếu người dùng không nhập API key, công cụ
vẫn chạy và cho kết quả đúng (vì toàn bộ các trường hợp khó của bộ dữ
liệu mẫu đã có sẵn đáp án xác minh trong VERIFIED_OVERRIDES), chỉ là sẽ
không tự xử lý được những dòng mới, lạ chưa từng gặp trong file khác.

Lấy API key miễn phí tại: https://aistudio.google.com/apikey
"""

import json
import re
import time


PROMPT_TEMPLATE = """Bạn là trợ lý tách tên đường/địa chỉ ở Việt Nam trong bảng giá đất.

Câu gốc: "{text}"

Hãy tách câu trên thành 3 phần:
- "duong": tên con đường/tuyến đường chính (KHÔNG kèm mô tả đoạn)
- "tu": điểm bắt đầu của đoạn đường (nếu câu có nêu, nếu không để rỗng)
- "den": điểm kết thúc của đoạn đường (nếu câu có nêu, nếu không để rỗng)

Nếu câu không có mô tả đoạn đường rõ ràng, để "tu" và "den" là chuỗi rỗng,
và "duong" là toàn bộ câu gốc.

CHỈ trả về đúng một JSON hợp lệ, không thêm chữ nào khác, đúng định dạng:
{{"duong": "...", "tu": "...", "den": "..."}}"""


class GeminiFallback:
    """
    Dùng: fallback = GeminiFallback(api_key)
          split_row(text, last_area_name, ai_fallback=fallback)
    """

    def __init__(self, api_key: str, model_name: str = "gemini-flash-lite-latest",
                 max_retries: int = 3, wait_seconds: int = 15, log=print):
        self.api_key = api_key
        self.model_name = model_name
        self.max_retries = max_retries
        self.wait_seconds = wait_seconds
        self.log = log
        self._model = None
        if api_key:
            try:
                import google.generativeai as genai
                genai.configure(api_key=api_key)
                self._model = genai.GenerativeModel(model_name)
            except ImportError:
                self.log("⚠️  Chưa cài 'google-generativeai' (pip install google-generativeai). "
                          "Bỏ qua bước gọi AI, các dòng khó sẽ giữ nguyên.")
            except Exception as e:
                self.log(f"⚠️  Không khởi tạo được Gemini: {e}")

    def available(self) -> bool:
        return self._model is not None

    def __call__(self, text: str):
        """Trả về (duong, tu, den) hoặc None nếu gọi AI thất bại."""
        if self._model is None:
            return None

        prompt = PROMPT_TEMPLATE.format(text=text)
        for attempt in range(self.max_retries):
            try:
                resp = self._model.generate_content(prompt)
                raw = resp.text.strip()
                raw = re.sub(r"^```json|```$", "", raw, flags=re.MULTILINE).strip()
                data = json.loads(raw)
                return (
                    data.get("duong", "").strip(),
                    data.get("tu", "").strip(),
                    data.get("den", "").strip(),
                )
            except json.JSONDecodeError:
                self.log(f"⚠️  AI trả JSON không hợp lệ cho câu: {text!r}")
                return None
            except Exception:
                wait = self.wait_seconds * (attempt + 1)
                self.log(f"⏳ Bị giới hạn tốc độ (lần {attempt + 1}/{self.max_retries}), "
                         f"chờ {wait}s rồi thử lại...")
                time.sleep(wait)

        self.log(f"❌ Gọi AI thất bại sau {self.max_retries} lần, giữ nguyên câu gốc: {text!r}")
        return None
