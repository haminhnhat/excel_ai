# Quy Trình Upload Excel Cho Người Dùng Không Kỹ Thuật

Tài liệu này mô tả luồng setup một file Excel tài chính mới mà không cần sửa YAML, không cần chạy lệnh Python, và không cần hiểu cấu trúc thư mục kỹ thuật.

## Mục tiêu

Người dùng nội bộ có thể:

```text
Upload Excel
→ hệ thống phân tích workbook
→ review các ô input/output được đề xuất trên trình duyệt
→ validate lựa chọn
→ tạo profile
→ chạy scenario
```

## Cách mở app

Lần đầu:

1. Double-click `Setup_First_Time.bat`.
2. Chờ đến khi thấy `Setup complete`.

Các lần sau:

1. Double-click `Start_App.bat`.
2. Trình duyệt sẽ mở:

```text
http://127.0.0.1:8000
```

Giữ cửa sổ màu đen mở trong lúc dùng app.

## Setup file Excel mới

1. Trong app, bấm tab `Setup new Excel file`.
2. Chọn file Excel `.xlsx` hoặc `.xlsm`.
3. Nhập tên profile dễ nhớ, ví dụ:

```text
project_a
```

4. Bấm `Analyze workbook`.
5. Chờ hệ thống phân tích workbook.
6. Review bảng mapping.

Ý nghĩa các lựa chọn:

- `Approve`: dùng ô này trong profile.
- `Review`: chưa quyết định, cần xem lại.
- `Reject`: không dùng ô này.

Nếu hệ thống gợi ý sai:

- Điền tên sheet đúng vào `Correct Sheet`.
- Điền ô hoặc range đúng vào `Correct Cell`.

Ví dụ:

```text
Correct Sheet: Assumptions
Correct Cell: C12
```

7. Bấm `Validate selection`.
8. Nếu validation báo thành công, bấm `Create profile`.
9. Quay lại tab `Run scenario`.

Profile và workbook vừa tạo sẽ được chọn sẵn để chạy scenario.

## Chạy scenario

Trong tab `Run scenario`, nhập câu lệnh như:

```text
Cho tôi NPV và IRR hiện tại
```

```text
Tăng lãi vay lên 8%
```

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

Bấm `Run scenario`, sau đó tải file kết quả ở phần download.

## Người review cần kiểm tra gì?

Input:

- Là ô mà app được phép sửa.
- Thường là ô nhập liệu thủ công trong Excel.
- Không nên là ô công thức.

Output:

- Là ô app chỉ đọc kết quả.
- Có thể là ô công thức, ví dụ NPV, IRR, lợi nhuận.

Correct Sheet / Correct Cell:

- Dùng khi hệ thống đoán sai vị trí.
- Người review có thể sửa trực tiếp trên UI.

## API liên quan

Các endpoint chính của workflow:

```text
GET  /api/profiles
POST /api/onboarding/analyze
POST /api/onboarding/validate-selection
POST /api/onboarding/create-profile
```

## Quy tắc an toàn quan trọng

Input không được trỏ vào ô công thức. Nếu một input được chọn là ô công thức, validation sẽ fail.

Output có thể là ô công thức vì hệ thống chỉ đọc kết quả, không ghi vào output.
