# BẢN MÔ TẢ CHI TIẾT QUÁ TRÌNH THỰC HIỆN DỰ ÁN

## 1. Tên dự án
**Xây dựng website hỗ trợ tách thông tin Đường – Từ – Đến từ file Excel bảng giá đất.**

## 2. Mục tiêu và bài toán
Trong bảng giá đất, một ô có thể ghi gộp tên đường và mô tả đoạn đường, ví dụ: “Quốc lộ 21B: Đoạn từ cầu Thạch Bích đến ngã ba bệnh viện”. Khi cần tra cứu hoặc đối chiếu, dữ liệu gộp gây khó khăn cho lọc, thống kê và chuẩn hóa. Dự án xây dựng công cụ nhận file Excel và tách mô tả đó thành ba trường: **Đường**, **Từ**, **Đến**. Kết quả được giữ trong file Excel để người dùng tiếp tục kiểm tra và sử dụng.

Mục tiêu triển khai là người dùng có thể mở website bằng trình duyệt, đăng nhập, tải file Excel, chọn chế độ xử lý, xem kết quả và tải file đầu ra mà không cần cài phần mềm xử lý trên máy của mình.

## 3. Phạm vi chức năng
1. Đăng nhập bằng tài khoản được quản trị viên cấu hình.
2. Tải file `.xlsx`; kiểm tra phần mở rộng, dung lượng, file rỗng/hỏng.
3. Liệt kê sheet, tự đoán sheet và cột có nội dung tên đường, hiển thị dữ liệu mẫu.
4. Chọn chế độ xử lý, sheet, cột và thứ tự ba cột kết quả.
5. Chạy phân tích theo tác vụ nền; hiển thị tiến trình, nhật ký và thống kê.
6. Xem trước kết quả, cảnh báo dòng chưa tách được và tải file Excel đầu ra.
7. Đăng xuất; tự hết hạn phiên; tự dọn file tạm theo thời gian TTL.

## 4. Công nghệ sử dụng
| Thành phần | Công nghệ | Vai trò |
|---|---|---|
| Giao diện | HTML, CSS, JavaScript | Biểu mẫu tải file, các bước xử lý, bảng xem trước, tiến trình và tải file |
| Web/API | Python, FastAPI, Uvicorn | Tiếp nhận HTTP, xác thực, kiểm tra dữ liệu đầu vào và cung cấp API |
| Xử lý bảng | pandas, openpyxl | Đọc/ghi workbook Excel, lấy dữ liệu và giữ cấu trúc bảng cần thiết |
| Bộ quy tắc | Python, regex | Nhận diện mẫu “từ … đến …”, tên đường và điểm đầu/cuối |
| AI tùy chọn | Gemini API | Thử hỗ trợ câu khó khi cấu hình API key; không phải điều kiện bắt buộc |
| Kiểm thử | pytest, FastAPI TestClient, Playwright tùy chọn | Kiểm tra logic, API và luồng người dùng |
| Đóng gói/triển khai | Docker, Render Blueprint | Đóng gói ứng dụng và định nghĩa cấu hình triển khai web service |

## 5. Các bước thực hiện chi tiết

### Bước 1. Khảo sát yêu cầu và dữ liệu đầu vào
- Xác định người dùng chính là người cần chuẩn hóa bảng giá đất từ file Excel.
- Khảo sát các dạng mô tả: tên đường kèm “đoạn từ… đến…”, “giáp… đến…”, dòng ghi “Trục chính/Đường làn trong”, và trường hợp chỉ có tên khu vực mà không mô tả điểm đầu/cuối.
- Chốt đầu vào là file `.xlsx`; mặc định giới hạn dung lượng 20 MB.
- Xác định đầu ra: ba trường Đường/Từ/Đến, thống kê các dòng đã xử lý và danh sách ví dụ chưa tách được để người dùng kiểm tra.

### Bước 2. Phân rã hệ thống và thiết kế kiến trúc
- Tách phần giao diện và phần xử lý dữ liệu thành các module độc lập.
- `server/static/index.html` đảm nhiệm giao diện web.
- `server/app.py` cung cấp trang web, đăng nhập và các API.
- `server/jobs.py` quản lý file tải lên, job phân tích, tiến trình và thời hạn lưu tạm.
- `core/splitter.py` chứa quy tắc tách một dòng; `core/generic.py` ghi thêm các cột vào workbook; `core/processor.py` xử lý định dạng bảng giá đối chiếu.
- Dùng chung lõi xử lý giúp bản web và bản máy tính tránh việc duy trì hai thuật toán khác nhau.

### Bước 3. Xây dựng thuật toán tách mô tả
Thuật toán ưu tiên các bước theo thứ tự sau:
1. Xử lý dòng phụ thuộc ngữ cảnh, như “Trục chính” hoặc “Đường làn trong”, bằng cách kế thừa tên khu vực của dòng liên quan phía trên khi dữ liệu phù hợp.
2. Áp dụng các trường hợp đã được kiểm tra và đối chiếu thủ công.
3. Dùng quy tắc chuỗi/biểu thức chính quy để nhận diện tên đường và cụm “từ X đến Y” hoặc “giáp X đến Y”.
4. Nếu người dùng bật AI và máy chủ có Gemini API key, gửi câu khó tới mô-đun AI dự phòng.
5. Nếu vẫn không chắc chắn, giữ lại nội dung gốc/đánh dấu chưa giải quyết thay vì tự tạo ra điểm đầu/cuối không có trong dữ liệu.

Cách xử lý bảo thủ giúp người dùng phân biệt kết quả theo quy tắc, kết quả AI và dữ liệu cần rà soát.

### Bước 4. Xây dựng chức năng đọc và kiểm tra Excel
- Nhận file trong request tải lên, kiểm tra đuôi `.xlsx` và giới hạn dung lượng khi ghi theo từng khối.
- Đọc workbook để xác định danh sách sheet, hàng tiêu đề và cột có khả năng chứa tên đường.
- Hiển thị bảng xem trước có giới hạn số hàng/cột nhằm giúp người dùng xác nhận lựa chọn trước khi phân tích.
- Không sử dụng tên file người dùng gửi lên như một đường dẫn tùy ý trên máy chủ; ứng dụng dùng thư mục tạm và mã định danh ngẫu nhiên cho các tệp/job.

### Bước 5. Xây dựng giao diện web ba bước
**Bước 1 – Tải dữ liệu:** chọn hoặc kéo-thả file Excel, xem tên file, dung lượng, sheet và cột được nhận diện. Nếu không tìm được cột phù hợp, người dùng có thể chọn sheet/cột thủ công.

**Bước 2 – Chạy phân tích:** chọn chế độ thêm ba cột hay tạo bảng đối chiếu, điều chỉnh thứ tự cột, chọn có ẩn cột gốc hay không và bật/tắt AI nếu có. Khi chạy, giao diện khóa những trường không nên thay đổi và hỏi máy chủ về tiến trình job.

**Bước 3 – Kết quả:** hiển thị số dòng xử lý, số dòng theo từng phương pháp, số dòng chưa tách, bảng xem trước và nút tải file Excel. Cho phép quay lại hoặc chạy lại khi cần.

### Bước 6. Thêm đăng nhập và bảo vệ API
- Cấu hình `APP_USERNAME`, `APP_PASSWORD` và `SESSION_SECRET` bằng biến môi trường/tệp `.env` local.
- Tạo trang `/login`; xác thực tên đăng nhập và mật khẩu bằng phép so sánh an toàn, không lưu mật khẩu trong mã nguồn.
- Sau khi đăng nhập, máy chủ cấp cookie phiên có chữ ký; cookie có thời hạn 8 giờ và dùng `SameSite=Lax`. Trong production, cookie được đánh dấu chỉ gửi qua HTTPS.
- Trang chính, API tải file, API xem trước, API tạo job và API tải kết quả đều yêu cầu phiên đăng nhập. API chưa xác thực trả lỗi 401; trang web chuyển về trang login.
- Giới hạn số lần đăng nhập sai theo địa chỉ IP, trả lỗi khi vượt ngưỡng và bổ sung header bảo mật cơ bản.
- Ẩn tài liệu API tương tác trong môi trường production. Endpoint `/api/health` vẫn được mở cho hệ thống kiểm tra trạng thái máy chủ.

Đây là xác thực một tài khoản đơn giản phù hợp bản demo; nếu mở rộng thành sản phẩm nhiều khách hàng, cần thêm quản lý người dùng, đổi/reset mật khẩu, CSRF toàn diện, ghi nhật ký bảo mật, và chính sách dữ liệu rõ ràng.

### Bước 7. Quản lý tác vụ và dữ liệu tạm
- Khi người dùng tải file lên, máy chủ trả về `file_id` cùng thông tin workbook.
- Khi bắt đầu chạy, máy chủ tạo `job_id` và xử lý trong nền để giao diện vẫn có thể hỏi tiến trình.
- Giao diện định kỳ hỏi trạng thái job, phần trăm hoàn thành, nhật ký và thống kê.
- Khi thành công, API tải file kết quả theo `job_id`; file đầu vào/job đã hết TTL được dọn để tránh giữ dữ liệu vô thời hạn. TTL, số file tối đa và số tác vụ đồng thời được cấu hình qua biến môi trường.

### Bước 8. Kiểm thử
Các nhóm kiểm thử chính:
- Quy tắc tách: đường có điểm đầu/cuối, dòng kế thừa tên khu vực, câu không khớp và kết quả cần AI.
- Excel: nhận diện sheet/cột, workbook nhiều sheet, định dạng sai, file rỗng/hỏng và file quá dung lượng.
- API: tham số hợp lệ/không hợp lệ, tạo job, xem tiến trình, tải kết quả và tình huống job không tồn tại/hết hạn.
- Bảo mật: người chưa đăng nhập bị từ chối; đăng nhập đúng được cấp phiên; sai mật khẩu không cấp phiên; đăng xuất làm mất quyền truy cập.
- Giao diện: luồng tải file → chạy → xem kết quả → tải tệp, thông báo lỗi và bố cục điện thoại (khi môi trường Playwright/Chromium có sẵn).

Kết quả kiểm thử tại thời điểm đóng gói bản này: **69 test đạt; 11 test E2E bị bỏ qua do môi trường kiểm thử không có Playwright/Chromium**. Có thể chạy lại E2E bằng `pip install playwright`, `playwright install chromium` rồi `python -m pytest -q tests/e2e`.

### Bước 9. Cấu hình môi trường và bảo vệ bí mật
- Tạo `.env.example` để hướng dẫn biến môi trường, nhưng bỏ qua `.env` thật trong Git.
- API key Gemini chỉ đọc từ `GEMINI_API_KEY`; không lưu khóa thật trong mã nguồn.
- Nếu một API key đã từng được đưa vào file chia sẻ hoặc commit Git, cần thu hồi/tạo lại key đó trong nhà cung cấp; chỉ xóa khỏi file hiện tại chưa đủ để xóa khỏi lịch sử Git.
- Production phải có `APP_PASSWORD` dài tối thiểu 12 ký tự và `SESSION_SECRET` dài tối thiểu 32 ký tự; ứng dụng dừng khởi động nếu thiếu để tránh vô tình chạy công khai với cấu hình không an toàn.

### Bước 10. Đóng gói và triển khai website
- Dockerfile cài thư viện từ `requirements-server.txt`, sao chép module cần chạy và khởi động `run_server.py`.
- `render.yaml` mô tả web service, health check và các biến môi trường cần nhập.
- Đẩy mã nguồn sạch lên GitHub, kết nối repository với Render, sau đó triển khai bằng Blueprint hoặc tạo Web Service chạy Docker.
- Nhập tài khoản/mật khẩu và khóa phiên trong phần Environment của Render; không đưa các giá trị đó vào Git.
- Khi deploy thành công, mở URL HTTPS được nền tảng cấp, đăng nhập, tải file mẫu, chạy cả hai chế độ, kiểm tra file tải về và thử đăng xuất.
- Chỉ gửi URL, tài khoản và mật khẩu cho người kiểm tra qua kênh riêng; không ghi mật khẩu vào repository hoặc tài liệu công khai.

### Bước 11. Quản lý mã nguồn và mời người kiểm tra
- Kiểm tra `git status`, rà soát không có `.env`, API key hoặc thông tin đăng nhập trong staged files.
- Commit các thay đổi với mô tả ngắn gọn; đẩy branch `main` lên repository.
- Trong GitHub, mở **Settings → Collaborators → Add people**, tìm `loctesla@gmail.com` hoặc username GitHub của người đó và gửi lời mời. Người được mời cần chấp nhận lời mời mới truy cập được.

## 6. Kiến trúc tổng thể

```text
Người dùng (trình duyệt)
        | HTTPS / cookie phiên
        v
Trang đăng nhập + giao diện HTML/CSS/JS
        | HTTP API + kiểm tra phiên
        v
FastAPI: upload / preview / jobs / download
        |
        +--> server/jobs.py: file tạm, trạng thái job, TTL
        |
        +--> core/splitter.py: tách mô tả đường
        +--> core/generic.py / processor.py: Excel in/out
        +--> Gemini API (tùy chọn)
        |
        v
File Excel kết quả -> trình duyệt tải xuống
```

## 7. Quy trình sử dụng website sau khi triển khai
1. Mở địa chỉ website HTTPS và đăng nhập.
2. Ở bước “Tải dữ liệu”, chọn file `.xlsx`, kiểm tra sheet/cột và bảng mẫu.
3. Sang bước “Chạy phân tích”, chọn chế độ và tùy chọn cần thiết.
4. Theo dõi tiến trình; nếu có lỗi, đọc thông báo, sửa lựa chọn rồi chạy lại.
5. Kiểm tra thống kê, ví dụ các dòng chưa giải quyết và bảng xem trước.
6. Bấm tải file kết quả; mở file bằng Excel để kiểm tra định dạng và một số dòng mẫu.
7. Bấm “Đăng xuất” khi kết thúc.

## 8. Kết quả đạt được và giới hạn
**Đạt được:** có giao diện web, API xử lý Excel, hai chế độ đầu ra, xem trước dữ liệu, tác vụ nền, đăng nhập, logout, giới hạn đăng nhập sai, cấu hình Docker/Render và kiểm thử tự động cho các chức năng chính.

**Giới hạn cần biết:** hệ thống hiện dùng một tài khoản chung; chưa có phân quyền nhiều người dùng hoặc lịch sử xử lý bền vững; dữ liệu tạm lưu trên máy chủ trong khoảng TTL; các dòng mô tả không theo mẫu có thể phải kiểm tra thủ công. Nếu dùng cho dữ liệu quan trọng hoặc quy mô lớn, cần thêm giám sát, lưu trữ phù hợp, sao lưu, kiểm soát dung lượng, HTTPS và quy trình quản trị người dùng.

## 9. Danh sách nghiệm thu
- [ ] Trang chủ chuyển về trang đăng nhập khi chưa xác thực.
- [ ] Đăng nhập đúng mở được giao diện, sai mật khẩu không mở được giao diện.
- [ ] Người chưa đăng nhập không gọi được API upload/job/download.
- [ ] Tải file Excel mẫu thành công; hệ thống nhận đúng sheet và cột.
- [ ] Chạy được chế độ thêm cột và chế độ bảng giá đối chiếu.
- [ ] Tải được file kết quả và kiểm tra dữ liệu bằng Excel.
- [ ] Đăng xuất xong, vào trang chủ lại phải đăng nhập.
- [ ] Biến bí mật nằm ngoài Git; API key cũ đã được thu hồi nếu từng bị lộ.
- [ ] Website triển khai có URL HTTPS và Health Check hoạt động.
- [ ] Gửi lời mời GitHub cho `loctesla@gmail.com` và chờ người nhận chấp nhận.
