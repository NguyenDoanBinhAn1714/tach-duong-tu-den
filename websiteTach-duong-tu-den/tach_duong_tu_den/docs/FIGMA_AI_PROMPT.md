# Prompt dán vào Figma AI (Figma Make / First Draft)

Mình không truy cập được Figma, nên bạn dán đoạn dưới đây vào Figma AI để nó dựng bản thiết kế trong tài khoản của bạn.
Ảnh tham chiếu để so sánh nằm ở `docs/design/` (01 đến 03 là máy tính, 04 là điện thoại).

```
Design a clean, professional web app called "Tách Đường · Từ · Đến" (Vietnamese UI). It splits the
"Tên đường phố" column of a land-price Excel sheet into three columns: Đường, Từ, Đến.
Desktop frame 1440x900 plus a mobile frame 390x844. Light UI, ocean-blue brand.

Global
- Background #F4F7FB. Cards white, radius 14, 1px border #D7E0EA, soft shadow (0 1 2 and 0 8 24, navy 8%).
- Top bar navy #0B2545, height 72: small logo square (ocean #2E86AB) + title "Tách Đường · Từ · Đến" (white, 20 semibold) +
  subtitle "Tách cột Tên đường phố trong bảng giá đất" (#9FB3C8, 14).
- Font: Segoe UI / Inter. Sizes 12, 14, 16, 20, 28. Body color #26384D, muted #5B6B80. Spacing scale 4/8/12/16/24/32/48.
- Primary button #1E6091 (hover #2E86AB), white text, radius 8, height 44. Secondary button white with #1E6091 border.
- A 3-step stepper under the top bar: "1 Tải dữ liệu", "2 Chạy phân tích", "3 Kết quả". Current step: filled ocean circle;
  done: circle with check; upcoming: gray outline. Thin connector line between steps.

Screen 1 - Tải dữ liệu gốc lên (Upload)
- Title "Tải dữ liệu gốc lên", helper text "Chọn tệp Excel .xlsx chứa bảng giá đất".
- Large dropzone, dashed border #1E6091, bg #F0F6FC, cloud-upload icon, text "Kéo thả tệp .xlsx vào đây hoặc bấm để chọn".
- State after upload: file chip "BGĐ_ Update Dự thảo Bảng giá_ Thanh Oai.xlsx · 41 KB · 3 sheet" with "Đổi tệp" link;
  pill tabs for sheets (selected = filled ocean): "Thanh Oai sau sáp nhập", "Giáp ranh + Ven trục", "Bảng giá";
  green info alert "Đã nhận diện cột B «Tên đường phố» ở sheet «Bảng giá»";
  preview table of first 12 rows with column letters A..J as header (#DEEAF6) and column B highlighted (#F0F6FC);
  primary button "Tiếp tục" with arrow, right aligned.

Screen 2 - Chạy phân tích (Analyze)
- Summary strip: "Tệp: BGĐ_ Update ... .xlsx · Sheet: Bảng giá".
- Two selectable option cards: "Thêm 3 cột vào tệp gốc" (selected) and "Bảng đối chiếu giá", each with icon and 1-line description.
- Fields: "Cột chứa tên đường" input (value "B"), "Thứ tự cột" select ("Đường → Từ → Đến"), checkbox "Ẩn cột gốc sau khi tách",
  checkbox "Dùng AI (Gemini) cho câu quy tắc không tách được" (disabled variant with hint "Máy chủ chưa có API key").
- Buttons: secondary "Quay lại", primary "Bắt đầu phân tích".
- Running variant: progress bar 62% with label "Đang tách dòng 79/127", monospace log box below, form dimmed.
- Error variant: red alert with message and button "Thử lại".

Screen 3 - Kết quả đầu ra (Results)
- Green success alert "Đã tách xong 127 dòng".
- Six stat cards: 127 "Dòng đã tách", 70 "Quy tắc", 51 "Kế thừa khu vực", 2 "Đã đối chiếu", 0 "AI (Gemini)", 4 "Giữ nguyên".
- Blue info line "Mở tệp Excel: ba cột tô xanh nhạt nằm ngay sau cột B".
- Preview table (40 rows) columns: Dòng | Tên đường phố (gốc) | Đường | Từ | Đến | Cách xử lý (small colored badge).
- Primary button with download icon "Tải tệp Excel kết quả", secondary "Chạy lại với cài đặt khác", text button "Xử lý tệp khác".

Mobile: single column, stepper shows only numbered dots + current label, stat cards 2 per row, tables scroll horizontally.
Provide components: Stepper, Card, Dropzone, FileChip, SheetTab, DataTable, OptionCard, Field, Checkbox, Button, ProgressBar, LogBox,
StatCard, Badge, Alert (info / success / warning / error).
```
