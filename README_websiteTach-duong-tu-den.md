# websiteTach-duong-tu-den

Website hỗ trợ đọc và chuẩn hóa dữ liệu bảng giá đất từ Excel, tự động tách mô tả tên đường phố thành ba trường **Đường – Từ – Đến**, xem trước kết quả và tải lại file Excel sau xử lý.

Dự án sử dụng **Python, FastAPI, pandas và openpyxl**. Giao diện web được xây dựng bằng HTML, CSS và JavaScript; phần xử lý có thể dùng quy tắc/biểu thức chính quy và tùy chọn gọi Google Gemini cho những câu khó.

> **Lưu ý về phiên bản hiện tại:** dự án có chức năng đăng ký và đăng nhập. Tuy nhiên, đây vẫn là bản ứng dụng thử nghiệm: file và tác vụ đang được lưu trong bộ nhớ dùng chung của máy chủ, chưa có phân quyền sở hữu dữ liệu theo từng tài khoản và kho tài khoản JSON cần được đặt trên vùng lưu trữ bền vững khi triển khai. Không dùng bản hiện tại để xử lý dữ liệu nhạy cảm hoặc mở công khai cho nhiều khách hàng trước khi hoàn thiện các điểm này.

---

## 1. Chức năng chính

### Tài khoản và đăng nhập

- Có trang **Đăng ký** và **Đăng nhập** ngay trên website.
- Người dùng đăng ký bằng tên đăng nhập, mật khẩu và nhập lại mật khẩu.
- Tên đăng nhập phải có từ 3 đến 100 ký tự; mật khẩu đăng ký từ 8 đến 256 ký tự.
- Từ chối tên đăng nhập đã tồn tại hoặc mật khẩu nhập lại không khớp.
- Mật khẩu tài khoản đăng ký được băm bằng PBKDF2-HMAC-SHA256 trước khi ghi vào `data/users.json`; không lưu mật khẩu dạng văn bản thuần.
- Phiên đăng nhập dùng cookie có chữ ký, thời hạn 8 giờ; đăng xuất sẽ xóa phiên.
- Có giới hạn số lần đăng nhập sai theo địa chỉ IP. Các trang và API chính yêu cầu đăng nhập.
- Có thêm một tài khoản quản trị cấu hình bằng biến môi trường `APP_USERNAME` và `APP_PASSWORD`.

### Xử lý file Excel

- Nhận file Excel định dạng `.xlsx`; mặc định giới hạn dung lượng **20 MB**.
- Kiểm tra file rỗng, phần mở rộng và khả năng đọc workbook.
- Đọc danh sách sheet, tự nhận diện sheet/cột có khả năng chứa tên đường phố.
- Hiển thị dữ liệu mẫu để người dùng xác nhận sheet và cột trước khi xử lý.
- Có hai chế độ xử lý:
  1. **Thêm cột:** thêm ba cột `Đường`, `Từ`, `Đến` vào workbook, mặc định đặt cạnh cột tên đường; có tùy chọn ẩn cột gốc.
  2. **Bảng đối chiếu giá:** đọc cấu trúc bảng giá đất và chuyển dữ liệu giá thành bảng kết quả dạng dài.
- Cho phép chọn thứ tự các cột đầu ra: `Đường – Từ – Đến` hoặc `Từ – Đến – Đường`.
- Chạy xử lý theo tác vụ nền; giao diện hiển thị trạng thái, tiến trình, nhật ký và thống kê.
- Xem trước kết quả, ghi nhận các dòng chưa tách chắc chắn và tải file Excel đầu ra.
- File tạm và tác vụ hoàn tất/lỗi được dọn theo thời hạn cấu hình.

### Thuật toán tách Đường – Từ – Đến

Thuật toán hiện tại áp dụng theo thứ tự:

1. Nhận diện nhãn phụ như `Trục chính`, `Đường làn trong`; nếu có tên khu vực hợp lệ ở trước thì ghép thành tên đường theo ngữ cảnh.
2. Kiểm tra các trường hợp đặc biệt đã được đối chiếu thủ công.
3. Áp dụng quy tắc chuỗi/regex để nhận diện các dạng `từ ... đến ...`, `giáp ... đến ...` hoặc mô tả có dấu hai chấm.
4. Nếu người dùng bật AI và đã cấu hình `GEMINI_API_KEY`, hệ thống có thể gọi Gemini để hỗ trợ câu khó.
5. Nếu chưa xác định được, giữ nội dung gốc và đánh dấu chưa giải quyết; không tự suy đoán điểm đầu hoặc điểm cuối.

Với những dòng chỉ là tiêu đề nhóm/khu vực hoặc dòng không có dữ liệu cần xử lý, ứng dụng có thể bỏ qua tùy chế độ xử lý. Kết quả cần được người dùng rà soát trước khi sử dụng chính thức.

---

## 2. Công nghệ sử dụng

| Thành phần | Công nghệ | Vai trò |
|---|---|---|
| Ngôn ngữ | Python 3.12 | Xây dựng ứng dụng và thuật toán xử lý |
| Web/API | FastAPI, Uvicorn | Máy chủ web và API |
| Giao diện | HTML, CSS, JavaScript | Đăng nhập, tải file, xem trước, tiến trình và kết quả |
| Xử lý Excel | pandas, openpyxl | Đọc, phân tích và ghi workbook Excel |
| Tách nội dung | Python, regular expressions | Nhận diện tên đường và cụm từ điểm đầu/điểm cuối |
| AI tùy chọn | Google Gemini API | Hỗ trợ xử lý một số mô tả khó khi được bật |
| Kiểm thử | pytest, FastAPI TestClient, Playwright | Kiểm thử logic, API và luồng giao diện |
| Đóng gói/triển khai | Docker, Render Blueprint | Đóng gói và cấu hình dịch vụ web |

---

## 3. Cấu trúc thư mục

```text
websiteTach-duong-tu-den/
├── core/
│   ├── ai_fallback.py       # Tích hợp Gemini tùy chọn
│   ├── generic.py           # Tách cột trong workbook Excel
│   ├── processor.py         # Xử lý bảng giá đất dạng đối chiếu
│   └── splitter.py          # Quy tắc tách Đường – Từ – Đến
├── server/
│   ├── app.py               # FastAPI, đăng ký, đăng nhập và API
│   ├── jobs.py              # Quản lý file tạm, job và tiến trình
│   └── static/
│       └── index.html       # Giao diện web
├── tests/                   # Kiểm thử tự động
├── docs/                    # Tài liệu thiết kế và quy trình
├── tools/                   # Công cụ hỗ trợ dự án
├── data/                    # Tạo khi chạy; lưu tài khoản đăng ký
├── main.py                  # Phiên bản ứng dụng máy tính Tkinter
├── run_server.py            # Điểm khởi động web server
├── run_server.bat           # Chạy nhanh trên Windows
├── requirements.txt         # Thư viện xử lý dữ liệu
├── requirements-server.txt  # Thư viện chạy website
├── requirements-dev.txt     # Thư viện kiểm thử
├── .env.example             # Mẫu cấu hình môi trường
├── .gitignore               # Loại trừ file riêng tư/cache khỏi Git
├── Dockerfile               # Đóng gói Docker
├── render.yaml              # Cấu hình Render Blueprint
└── README.md
```

`data/users.json` được tạo khi có người đăng ký tài khoản đầu tiên. Tệp này là dữ liệu runtime và đã được `.gitignore` loại khỏi Git.

---

## 4. Cài đặt và chạy trên Windows

### Yêu cầu

- Windows 10/11.
- Python 3.12 khuyến nghị.
- PowerShell hoặc terminal trong VS Code.
- Git nếu muốn quản lý mã nguồn trên GitHub.

### Bước 1 – Mở thư mục dự án

Mở PowerShell tại thư mục gốc `websiteTach-duong-tu-den` – nơi có file `run_server.py`.

### Bước 2 – Tạo môi trường Python

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

Nếu PowerShell chặn kích hoạt môi trường, có thể chạy lệnh sau trong cửa sổ PowerShell hiện tại rồi kích hoạt lại:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### Bước 3 – Cài thư viện

```powershell
pip install -r requirements-server.txt
```

### Bước 4 – Tạo file cấu hình

```powershell
Copy-Item .env.example .env
notepad .env
```

Các biến thường dùng:

| Biến | Ý nghĩa | Giá trị gợi ý khi chạy local |
|---|---|---|
| `APP_ENV` | Môi trường chạy | `development` |
| `APP_USERNAME` | Tài khoản quản trị | `admin` |
| `APP_PASSWORD` | Mật khẩu quản trị | Tự đặt; có thể để trống khi chạy local để ứng dụng sinh mật khẩu tạm |
| `SESSION_SECRET` | Khóa ký cookie phiên | Tạo chuỗi ngẫu nhiên; local có thể để trống để ứng dụng tự sinh |
| `MAX_UPLOAD_MB` | Dung lượng upload tối đa | `20` |
| `SESSION_TTL_SECONDS` | Thời hạn lưu file/job tạm | `3600` (1 giờ) |
| `REGISTRATION_ENABLED` | Bật/tắt đăng ký tài khoản | `true` |
| `USER_STORE_PATH` | Đường dẫn lưu tài khoản đăng ký | `data/users.json` |
| `GEMINI_API_KEY` | API key Gemini, không bắt buộc | Để trống nếu không dùng AI |

Tạo khóa phiên ngẫu nhiên bằng lệnh:

```powershell
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

**Lưu ý:** `SESSION_TTL_SECONDS` là thời gian giữ file và tác vụ tạm, không phải thời hạn cookie đăng nhập. Cookie phiên đăng nhập hiện có thời hạn 8 giờ.

Nếu để trống `APP_PASSWORD` khi chạy local, chương trình tạo mật khẩu tạm và in ra terminal. Hãy đọc mật khẩu đó để đăng nhập lần đầu. Không sử dụng cấu hình/mật khẩu tạm cho website công khai.

### Bước 5 – Khởi động website

```powershell
python run_server.py
```

Mặc định máy chủ chạy tại `127.0.0.1:8000`. Mở trình duyệt:

**http://127.0.0.1:8000**

Lần đầu sử dụng, chọn **Đăng ký ngay**, nhập tên đăng nhập và mật khẩu (tối thiểu 8 ký tự), nhập lại mật khẩu rồi tạo tài khoản. Sau khi đăng ký thành công, hệ thống tự đăng nhập và chuyển đến giao diện xử lý Excel.

Nếu cổng 8000 đang được sử dụng, có thể đổi cổng trong PowerShell:

```powershell
$env:PORT="8001"
python run_server.py
```

Sau đó mở `http://127.0.0.1:8001`.

---

## 5. Cách sử dụng website

1. **Đăng ký/đăng nhập:** mở website, tạo tài khoản mới hoặc đăng nhập bằng tài khoản đã có.
2. **Tải dữ liệu:** chọn file `.xlsx`; hệ thống kiểm tra file, liệt kê sheet, nhận diện cột và hiển thị dữ liệu mẫu.
3. **Chọn phương án xử lý:** chọn chế độ thêm ba cột hoặc tạo bảng đối chiếu; xác nhận sheet/cột và thứ tự cột kết quả.
4. **Chạy phân tích:** bắt đầu tác vụ và theo dõi phần trăm hoàn thành, nhật ký và thống kê.
5. **Kiểm tra kết quả:** xem trước dữ liệu đã tách; rà soát các dòng chưa xác định hoặc cảnh báo.
6. **Tải file:** tải file Excel kết quả về máy và mở bằng Excel để kiểm tra.
7. **Đăng xuất:** chọn đăng xuất khi kết thúc phiên làm việc.

### Cấu trúc dữ liệu ở chế độ Bảng đối chiếu giá

Chế độ này hiện được viết theo cấu trúc bảng giá mẫu trong mã nguồn:

- Tên sheet mặc định: `Bảng giá`.
- Cột B: tên đường phố.
- Các nhóm cột D–O: dữ liệu giá của VT1, VT2, VT3, VT4; mỗi vị trí gồm giá thấp nhất, giá cao nhất và giá trung bình.

Nếu file Excel thực tế khác cấu trúc trên, cần kiểm tra và điều chỉnh `core/processor.py` trước khi xử lý hàng loạt.

---

## 6. Cấu hình AI Gemini (không bắt buộc)

Chức năng tách bằng quy tắc vẫn hoạt động khi không có Gemini API key. Nếu cần bật AI cho các dòng khó:

1. Tạo API key tại Google AI Studio.
2. Mở file `.env`.
3. Gán key cho biến `GEMINI_API_KEY`:

```env
GEMINI_API_KEY=YOUR_GEMINI_API_KEY
```

4. Khởi động lại server.
5. Bật tùy chọn AI trong giao diện khi chạy phân tích.

Không đăng key lên GitHub, không gửi key trong ảnh chụp màn hình và không đưa key vào tài liệu công khai. Nếu key đã bị commit hoặc chia sẻ công khai, cần thu hồi key cũ và tạo key mới.

---

## 7. Chạy kiểm thử

Cài thư viện phát triển:

```powershell
pip install -r requirements-dev.txt
```

Chạy toàn bộ kiểm thử:

```powershell
python -m pytest -q
```

Chạy riêng nhóm kiểm thử giao diện đầu-cuối:

```powershell
pip install playwright
playwright install chromium
python -m pytest -q tests/e2e
```

### Kết quả chạy trên gói mã nguồn được dùng để viết README

Lần chạy `python -m pytest -q` ngày **01/11/2026** cho kết quả:

- 71 bài kiểm thử đạt.
- 2 bài kiểm thử API chưa đạt trong `tests/test_wizard_api.py` (kỳ vọng số dòng xử lý/kết quả chưa khớp với dữ liệu fixture và logic hiện tại).
- 11 bài kiểm thử giao diện E2E bị bỏ qua trong môi trường hiện tại.

Vì vậy, trước khi xem đây là bản phát hành hoàn chỉnh, cần xử lý hai kiểm thử API đang lỗi và chạy lại toàn bộ bộ kiểm thử. Kết quả trên có thể thay đổi sau khi mã nguồn được cập nhật.

---

## 8. Triển khai website bằng Render

Dự án có sẵn `Dockerfile` và `render.yaml` để cấu hình triển khai.

### Chuẩn bị GitHub

1. Tạo repository tên `websiteTach-duong-tu-den`.
2. Đưa **các file bên trong thư mục dự án** lên repository; các file như `Dockerfile`, `render.yaml`, `run_server.py`, `requirements-server.txt`, `core/` và `server/` phải nằm trực tiếp ở thư mục gốc repository.
3. Kiểm tra trước khi commit rằng không có `.env`, `data/users.json`, API key, file Excel thật hoặc mật khẩu.
4. Commit và push mã nguồn lên GitHub.

Ví dụ lệnh Git (thực hiện tại thư mục dự án):

```powershell
git init
git add .
git status
git commit -m "Initial project website tach duong tu den"
git branch -M main
git remote add origin https://github.com/<GITHUB_USERNAME>/websiteTach-duong-tu-den.git
git push -u origin main
```

Thay `<GITHUB_USERNAME>` bằng tên tài khoản GitHub của người sở hữu repository. Không chạy `git add -f` với `.env` hoặc `data/users.json`.

### Tạo Web Service trên Render

1. Đăng nhập Render và kết nối GitHub.
2. Chọn tạo Blueprint từ repository có `render.yaml`, hoặc tạo Web Service và chọn runtime Docker.
3. Cấu hình các biến môi trường bí mật:
   - `APP_ENV=production`
   - `APP_USERNAME`: tên tài khoản quản trị
   - `APP_PASSWORD`: mật khẩu quản trị ít nhất 12 ký tự
   - `SESSION_SECRET`: chuỗi ngẫu nhiên ít nhất 32 ký tự
   - `REGISTRATION_ENABLED=true` nếu cho phép người dùng tự đăng ký
   - `USER_STORE_PATH=/app/data/users.json` nếu dùng vùng lưu trữ gắn tại `/app/data`
4. Đảm bảo Health Check Path là `/api/health`.
5. Deploy; sau khi hoàn tất, mở URL HTTPS Render cấp và thử đăng ký, đăng nhập, tải file Excel, xử lý và tải kết quả.

### Lưu ý bắt buộc về lưu tài khoản khi chạy trên máy chủ

Tài khoản đăng ký hiện được ghi vào file JSON (`data/users.json`), không phải PostgreSQL. `render.yaml` hiện chưa khai báo ổ đĩa lưu trữ bền vững. Nếu triển khai trên môi trường có filesystem tạm thời, file tài khoản có thể mất khi service được tạo lại hoặc triển khai lại.

Để dùng ổn định, cần cấu hình persistent disk và đặt `USER_STORE_PATH` trỏ vào thư mục được gắn ổ, hoặc nâng cấp phần lưu tài khoản sang cơ sở dữ liệu như PostgreSQL. Không được đưa `data/users.json` chứa tài khoản người dùng lên GitHub.

### Giới hạn nhiều tài khoản cần biết

Đăng ký/đăng nhập đã có, nhưng file upload và job đang được quản lý bởi một `STORE` trong bộ nhớ chung của tiến trình. Mã nguồn hiện chưa gắn từng file/job với `username` sở hữu và chưa kiểm tra quyền sở hữu khi xem trước/tải kết quả. Do đó, **không coi đăng ký tài khoản hiện tại là cơ chế cách ly dữ liệu giữa nhiều người dùng**. Nếu mở website công khai cho nhiều người dùng hoặc xử lý dữ liệu riêng tư, cần bổ sung phân quyền theo chủ sở hữu, lưu trữ bền vững, kiểm thử bảo mật và chính sách xóa dữ liệu trước.

---

## 9. Bảo mật và lưu ý dữ liệu

- Không commit `.env`, `data/users.json`, API key hoặc mật khẩu.
- Giữ repository ở chế độ Private trong quá trình phát triển nếu chưa muốn công khai mã nguồn.
- Khi chạy production, đặt `APP_ENV=production`, `APP_PASSWORD` tối thiểu 12 ký tự và `SESSION_SECRET` tối thiểu 32 ký tự. Ứng dụng sẽ dừng khởi động nếu thiếu các giá trị bắt buộc.
- Cookie phiên được cấu hình `HttpOnly` bởi SessionMiddleware, `SameSite=Lax` và chỉ gửi qua HTTPS trong production.
- File upload chỉ được giữ tạm theo cơ chế TTL; không coi máy chủ hiện tại là nơi lưu trữ hồ sơ Excel lâu dài.
- Không tải lên dữ liệu cá nhân/nhạy cảm khi chưa có quy trình bảo vệ, phân quyền, sao lưu và xóa dữ liệu phù hợp.
- Sau khi xử lý, người dùng cần kiểm tra thủ công các dòng được đánh dấu chưa xác định trước khi dùng kết quả làm dữ liệu chính thức.

---

## 10. Mời giảng viên/người hướng dẫn xem mã nguồn

Trong GitHub repository:

1. Mở **Settings**.
2. Chọn **Collaborators** (trong phần Access).
3. Nhấn **Add people**.
4. Nhập email `loctesla@gmail.com` hoặc username GitHub của người hướng dẫn.
5. Chọn đúng tài khoản và gửi lời mời.
6. Người được mời cần chấp nhận lời mời thì mới truy cập được repository.

---

## 11. Hướng phát triển tiếp theo

- Sửa hai kiểm thử API đang chưa đạt và bổ sung test hồi quy cho các cấu trúc Excel thực tế.
- Chuyển kho tài khoản từ JSON sang PostgreSQL hoặc kho dữ liệu bền vững.
- Bổ sung quyền sở hữu file/job theo tài khoản, bảo đảm người dùng chỉ xem/tải dữ liệu của mình.
- Bổ sung xác minh email, đổi và khôi phục mật khẩu, quản trị tài khoản và nhật ký hoạt động.
- Hoàn thiện chính sách lưu trữ, xóa file và sao lưu dữ liệu.
- Mở rộng bộ quy tắc tách dựa trên các file Excel đã được người dùng xác nhận.

---

## 12. Tài liệu trong dự án

- `docs/QUY_TRINH_THUC_HIEN.md`: mô tả quy trình thực hiện dự án.
- `docs/QUY_TRINH_THUC_HIEN.docx`: bản Word của quy trình.
- `docs/THIET_KE_GIAO_DIEN.md`: mô tả thiết kế giao diện.
- `docs/THIET_KE_TOOL.md`: mô tả thiết kế công cụ.
- `docs/design/`: hình ảnh mô phỏng các màn hình.

---

**Tên dự án:** websiteTach-duong-tu-den  
**Mục đích:** Hỗ trợ chuẩn hóa thông tin đường phố từ bảng giá đất Excel thành các trường Đường – Từ – Đến.
