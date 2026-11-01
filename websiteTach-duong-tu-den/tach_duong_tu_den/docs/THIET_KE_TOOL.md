# Thiết kế tool "Tách Đường / Từ / Đến"

## 1. Bài toán

Bảng giá đất ghi tên đường và đoạn đường chung trong một ô, ví dụ:

> Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba giao cắt với đường bệnh viện Thanh Oai

Khó tra cứu và đối chiếu theo từng đoạn. Tool tách ô này thành ba phần:

| Đường | Từ | Đến |
|---|---|---|
| Quốc lộ 21B | cầu Thạch Bích | ngã ba giao cắt với đường bệnh viện Thanh Oai |

## 2. Đầu vào / đầu ra

| Chế độ | Đầu vào | Đầu ra |
|---|---|---|
| `add-columns` | File .xlsx bất kỳ có cột tên đường | Chính file đó, chèn 3 cột "Đường / Từ / Đến (tách)" ngay bên phải cột tên đường, tô xanh nhạt để dễ thấy. Dữ liệu gốc giữ nguyên, các cột bên phải dịch sang phải (công thức, ô gộp, độ rộng cột, bộ lọc dịch theo). Có tuỳ chọn ẩn cột gốc |
| `price-table` | File bảng giá Thanh Oai (sheet "Bảng giá", 4 vị trí VT1–VT4) | Bảng 8 cột, mỗi mức giá một dòng, đúng định dạng file mẫu (1029 dòng với file hiện có) |

## 3. Kiến trúc

Thuật toán nằm ở `core/`, hai bản chạy (máy tính và server) chỉ là hai "vỏ" khác nhau gọi cùng lõi đó.

```
   Bản máy tính (PC-based)            Bản server (server-based)
   ┌─────────────────────┐            ┌──────────────────────────┐
   │ main.py (Tkinter)   │            │ server/app.py (FastAPI)  │
   │ chạy trên máy mình  │            │ trang web + REST API     │
   └──────────┬──────────┘            └────────────┬─────────────┘
              └──────────────┬─────────────────────┘
                             ▼
   core/splitter.py     tách một câu -> (Đường, Từ, Đến)
   core/processor.py    đọc bảng giá Thanh Oai, tách theo vị trí giá
   core/generic.py      tách cột bất kỳ, ghi thêm 3 cột vào file
   core/ai_fallback.py  gọi Gemini cho câu quy tắc không tách được (tuỳ chọn)
```

## 4. Thuật toán tách một câu (`core/splitter.py`)

Thứ tự ưu tiên:

1. Nhãn phụ "Trục chính" / "Đường làn trong" lấy tên khu vực của dòng phía trên.
2. Câu nằm trong danh sách đã đối chiếu tay (hiện có 1 câu, xuất hiện 2 lần trong file Thanh Oai) dùng đáp án có sẵn.
3. Quy tắc regex. Nếu dấu `:` đứng trước cụm "từ/giáp" thì cắt theo dấu hai chấm, ngược lại tìm "từ X đến Y" hoặc "giáp X đến Y" trong cả câu.
4. Không khớp quy tắc nào thì đưa cho AI (nếu bật), không có AI thì giữ nguyên câu gốc.

Kết quả trên file Thanh Oai (127 dòng dữ liệu): 70 dòng do regex tách, 51 dòng nhãn phụ kế thừa tên khu vực, 2 dòng dùng đáp án đối chiếu sẵn, 4 dòng không có mô tả đoạn (ví dụ "Khu đấu giá Sau Trại") được giữ nguyên. Không dòng nào phải gọi AI, và kết quả khớp file mẫu 1029/1029 dòng.

## 5. PC-based và server-based

**PC-based**: chương trình cài và chạy ngay trên máy người dùng.
**Server-based**: chương trình chạy trên một máy chủ, người dùng chỉ mở trình duyệt, tải file lên và nhận kết quả.

| | PC-based | Server-based |
|---|---|---|
| Cách dùng | Mở ứng dụng trên máy mình | Mở địa chỉ web trên trình duyệt |
| Cài đặt | Mỗi máy cài một lần (hoặc gửi file .exe) | Cài một lần trên server, người dùng không cài gì |
| Cập nhật | Phải gửi bản mới cho từng người | Sửa một nơi, mọi người dùng bản mới |
| Dữ liệu | File Excel không rời khỏi máy | File được tải lên server, cần bảo mật đường truyền và máy chủ |
| Nhiều người cùng dùng | Mỗi người một bản riêng | Dùng chung một hệ thống |
| Tích hợp hệ thống khác | Khó | Gọi qua API (`/api/split-file`, `/api/split-text`) |
| Cần gì để chạy | Python hoặc file .exe | Máy chủ luôn bật, mở cổng mạng |
| Phù hợp khi | Một vài người dùng, dữ liệu nhạy cảm | Cả nhóm/công ty dùng chung, hoặc cần gọi từ hệ thống khác |

Dự án có cả hai trên cùng một lõi nên chọn kiểu nào cũng không phải viết lại thuật toán.

## 6. Cách chạy

**Bản máy tính**

```bash
pip install -r requirements.txt
python main.py
```

Đóng gói thành một file .exe cho Windows: chạy `build_exe.bat` (cần cài Python và pyinstaller trên máy build).

**Bản server**

```bash
pip install -r requirements-server.txt
python run_server.py          # mở http://127.0.0.1:8000
```

- Cho cả mạng nội bộ truy cập: `HOST=0.0.0.0 python run_server.py` (PowerShell: `$env:HOST="0.0.0.0"`).
- Trang `/docs` là giao diện thử API tự sinh của FastAPI.
- Biến môi trường: `HOST`, `PORT`, `MAX_UPLOAD_MB` (mặc định 20), `GEMINI_API_KEY` (key nằm trên server, người dùng không cần nhập).
- Docker: `docker build -t tach-duong .` rồi `docker run -p 8000:8000 -e GEMINI_API_KEY=... tach-duong`.

## 7. API

| Phương thức | Đường dẫn | Mô tả |
|---|---|---|
| GET | `/api/health` | Kiểm tra server, cho biết AI đã cấu hình chưa |
| POST | `/api/split-text` | JSON `{"text": "...", "use_ai": false}` trả về `{duong, tu, den, source}` |
| POST | `/api/split-file` | Form: `file` (.xlsx), `mode` (`add-columns` / `price-table`), `sheet`, `column`, `order` (`duong-tu-den` / `tu-den-duong`), `use_ai`. Trả về file .xlsx; thống kê nằm ở header `X-Stats` |

Luồng 3 bước của giao diện web dùng thêm các điểm API sau (có thể gọi trực tiếp):

| Phương thức | Đường dẫn | Mô tả |
|---|---|---|
| POST | `/api/upload` | Bước 1: nhận tệp, trả `file_id`, danh sách sheet, cột tên đường đoán được và bảng xem trước |
| GET | `/api/preview?file_id=&sheet=` | Xem trước một sheet khác của tệp đã tải |
| POST | `/api/jobs` | Bước 2: JSON `{file_id, mode, sheet, column, order, use_ai, hide_source}`, chạy nền, trả `job_id` |
| GET | `/api/jobs/{job_id}` | Tiến độ, nhật ký, thống kê và bảng xem trước kết quả |
| GET | `/api/jobs/{job_id}/download` | Bước 3: tải tệp Excel kết quả |

Tệp tải lên và kết quả nằm trong thư mục tạm, tự xoá sau `SESSION_TTL_SECONDS` giây (mặc định 3600). Giới hạn đồng thời: `MAX_ACTIVE_JOBS` (4), `MAX_STORED_FILES` (50).

Ví dụ:

```bash
curl -F "file=@bang_gia.xlsx" -F "mode=add-columns" -o ket_qua.xlsx http://127.0.0.1:8000/api/split-file
```

## 8. Giới hạn cần biết

- Bản server chưa có đăng nhập. Chỉ nên chạy trong mạng nội bộ, hoặc đặt sau reverse proxy có HTTPS và xác thực trước khi mở ra internet.
- Chế độ `add-columns` lưu lại bằng openpyxl: giữ dữ liệu, công thức, ô gộp; hình ảnh và biểu đồ trong file (nếu có) có thể mất.
- Nếu sheet có định dạng có điều kiện, kiểm tra dữ liệu, bảng Excel, ảnh/biểu đồ, công thức mảng, hoặc có sheet khác có công thức trỏ vào, tool không chèn cột mà ghi 3 cột vào cuối bảng và nói rõ lý do (trường `warning`), để không làm hỏng file.
- Quy tắc viết theo cách diễn đạt của bảng giá Thanh Oai ("từ…đến", "giáp…đến"). File có cách viết khác có thể có câu không tách được, lúc đó bật AI hoặc sửa tay.
- Gemini bản miễn phí hay bị giới hạn tốc độ nên AI chỉ là phương án dự phòng, mặc định tắt.
- Chưa chạy thử trong môi trường soạn code: `build_exe.bat` (cần máy Windows) và `Dockerfile` (cần Docker). Phần còn lại đã chạy và test thật.

## 9. Kiểm thử

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

Gồm test cho thuật toán tách (ca lấy từ file thật), pipeline hai chế độ và các endpoint API.
