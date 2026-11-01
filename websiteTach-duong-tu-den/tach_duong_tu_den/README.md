# Tách Đường / Từ / Đến

Ứng dụng web hỗ trợ đọc file Excel bảng giá đất, nhận diện sheet/cột chứa mô tả đường phố, tách thành các trường **Đường – Từ – Đến**, xem trước kết quả và tải file Excel đã xử lý. Dự án sử dụng FastAPI và giao diện HTML/CSS/JavaScript; thuật toán xử lý được dùng chung với ứng dụng máy tính.

## Chức năng chính

- Đăng nhập bằng tài khoản/mật khẩu cấu hình qua biến môi trường; cookie phiên có chữ ký, tự hết hạn sau 8 giờ.
- Chỉ nhận file `.xlsx`, mặc định tối đa 20 MB; từ chối file rỗng/hỏng hoặc tham số không hợp lệ.
- Tự nhận diện sheet và cột có tên đường; cho xem trước trước khi xử lý.
- Hai chế độ: thêm ba cột Đường/Từ/Đến vào file gốc hoặc chuyển bảng giá sang bảng đối chiếu.
- Chạy tác vụ nền, xem tiến trình/nhật ký, xem thống kê và tải file kết quả. Tệp đầu vào/kết quả tạm tự dọn sau khoảng thời gian TTL cấu hình.
- Có thể bật Gemini tùy chọn cho câu khó bằng `GEMINI_API_KEY`; không có key thì quy tắc vẫn chạy và những câu chưa tách được sẽ được giữ lại để kiểm tra.

## 1. Chạy trên Windows

Mở PowerShell, vào thư mục chứa dự án và chạy:

```powershell
cd tach_duong_tu_den
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements-server.txt
Copy-Item .env.example .env
notepad .env
```

Trong `.env`, đặt `APP_USERNAME`, `APP_PASSWORD` (tối thiểu 12 ký tự khi triển khai production) và `SESSION_SECRET` (chuỗi ngẫu nhiên dài ít nhất 32 ký tự). Có thể tạo khóa phiên bằng lệnh:

```powershell
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

Sau đó chạy:

```powershell
python run_server.py
```

Mở **http://127.0.0.1:8000**. Nếu không đặt `APP_PASSWORD`, ứng dụng local sinh mật khẩu tạm và in vào terminal. Đừng dùng mật khẩu tạm đó cho website công khai.

## 2. Triển khai website lên Internet bằng Render

Dự án có `Dockerfile` và `render.yaml`. Sau khi đẩy mã nguồn lên GitHub, đăng nhập Render và tạo **New → Blueprint** (hoặc New → Web Service, chọn repository và runtime Docker). Render dùng `Dockerfile` để build/chạy web service. Một web service phải lắng nghe trên `0.0.0.0` và cổng do nền tảng cung cấp. Xem hướng dẫn chính thức: [Render Web Services](https://render.com/docs/web-services), [Docker](https://render.com/docs/docker).

Khi cấu hình biến môi trường trong Render, đặt các giá trị bí mật sau; không đưa mật khẩu thật vào Git:

- `APP_ENV=production`
- `APP_USERNAME`: tên tài khoản gửi cho người kiểm tra
- `APP_PASSWORD`: mật khẩu dài tối thiểu 12 ký tự
- `SESSION_SECRET`: chuỗi ngẫu nhiên dài ít nhất 32 ký tự (Render Blueprint có thể tự tạo)
- `MAX_UPLOAD_MB=20`
- `SESSION_TTL_SECONDS=3600`

Nếu tạo service thủ công thay vì dùng Blueprint, đặt Health Check Path là `/api/health`. Sau khi trạng thái Deploy thành công, dùng URL `.onrender.com` do Render cấp, mở URL để thử đăng nhập và tải một file Excel mẫu. Render có thể áp dụng giới hạn tài nguyên/thời gian ngủ tùy gói; kiểm tra dashboard của service trước khi bàn giao.

## 3. Đẩy mã nguồn lên GitHub

Nếu chưa có repository, tạo một repository trống trên GitHub (không chọn thêm README để tránh xung đột), sau đó ở PowerShell tại thư mục dự án:

```powershell
git init
git add .
git status
git commit -m "Hoan thien website tach duong va dang nhap"
git branch -M main
git remote add origin https://github.com/<TEN_GITHUB>/<TEN_REPOSITORY>.git
git push -u origin main
```

Thay hai phần trong URL bằng repo thật của mình. Kiểm tra `git status` và xác nhận `.env`, `config.py`, API key, mật khẩu không được đưa lên Git. Nếu `config.py` từng được commit ở một phiên bản cũ, hãy gỡ khỏi Git bằng `git rm --cached config.py`, commit thay đổi, đồng thời thu hồi/tạo lại API key đã lộ. Việc xóa khỏi commit mới không tự xóa bí mật khỏi lịch sử Git.

### Mời chú vào repository

Vào GitHub → mở repository → **Settings → Collaborators → Add people**, nhập `loctesla@gmail.com` (hoặc username GitHub nếu email chưa được GitHub nhận diện), chọn đúng người rồi gửi lời mời. Chú cần chấp nhận lời mời thì mới truy cập được. Tham khảo [hướng dẫn chính thức của GitHub](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/repository-access-and-collaboration/inviting-collaborators-to-a-personal-repository).

## 4. Chạy kiểm thử

```powershell
pip install -r requirements-dev.txt
python -m pytest -q
```

Muốn chạy test giao diện trên trình duyệt thật:

```powershell
pip install playwright
playwright install chromium
python -m pytest -q tests/e2e
```

Các kiểm thử bao gồm thuật toán tách, kiểm tra file Excel, API, tiến trình job, tải kết quả, xác thực đăng nhập và phiên đăng xuất.

## Cấu trúc mã nguồn

```text
main.py                 Ứng dụng máy tính Tkinter
run_server.py           Khởi động web server
server/app.py           FastAPI, đăng nhập, API và xác thực
server/jobs.py          Lưu tệp tạm, quản lý tiến trình, dọn dữ liệu hết hạn
server/static/index.html Giao diện web ba bước
core/splitter.py        Quy tắc tách một mô tả đường
core/generic.py         Thêm ba cột vào workbook Excel
core/processor.py       Xử lý bảng giá đối chiếu
core/ai_fallback.py     AI Gemini tùy chọn
tests/                  Kiểm thử tự động
docs/                   Tài liệu thiết kế và hướng dẫn
Dockerfile              Đóng gói để triển khai
render.yaml             Cấu hình Render Blueprint
.env.example            Mẫu biến môi trường không chứa bí mật
```

## Lưu ý trước khi dùng với dữ liệu thật

Website xử lý file đã tải lên trên máy chủ; không nên tải dữ liệu mật nếu chưa thống nhất chính sách lưu trữ và quyền truy cập. Phiên bản này dùng xác thực một tài khoản qua biến môi trường, chưa có quản lý nhiều người dùng, đổi mật khẩu trên web hay cơ sở dữ liệu tài khoản. Tệp và job hiện lưu tạm trên ổ đĩa máy chủ và có TTL; không coi đây là kho lưu trữ lâu dài.
