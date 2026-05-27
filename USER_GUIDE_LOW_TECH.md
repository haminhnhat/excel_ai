# Hướng Dẫn Sử Dụng Excel AI Controller

Tài liệu này dành cho người dùng nội bộ không chuyên kỹ thuật.

## 1. Chuẩn bị

Máy tính cần có:

- Microsoft Excel.
- Python đã được cài sẵn nếu đây là lần setup đầu tiên.
- Thư mục app `Excel AI Controller`.

Nếu chưa có Python, nhờ IT/admin cài Python 3.11 hoặc mới hơn từ:

```text
https://www.python.org/downloads/
```

Khi cài Python, nhớ chọn `Add python.exe to PATH`.

## 2. Mở app lần đầu

1. Mở thư mục app.
2. Double-click `Setup_First_Time.bat`.
3. Chờ đến khi thấy dòng:

```text
Setup complete.
```

4. Double-click `Start_App.bat`.
5. Trình duyệt sẽ tự mở app.

Nếu trình duyệt không tự mở, vào địa chỉ:

```text
http://127.0.0.1:8000
```

Giữ cửa sổ màu đen mở trong lúc dùng app. Khi đóng cửa sổ đó, app sẽ dừng.

## 3. Chạy scenario

1. Chọn tab `Run scenario`.
2. Chọn profile phù hợp.
3. Nhập câu lệnh.
4. Bấm `Run scenario`.
5. Xem kết quả trên màn hình.
6. Bấm download để tải file Excel scenario.

Ví dụ câu lệnh:

```text
Cho tôi NPV và IRR hiện tại
```

```text
Tăng lãi vay lên 8%
```

```text
Giảm giá bán 5%
```

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

## 4. Setup file Excel mới

Dùng bước này khi có workbook tài chính mới chưa có profile.

1. Chọn tab `Setup new Excel file`.
2. Bấm chọn file Excel.
3. Nhập tên profile ngắn, ví dụ:

```text
hoa_yen
```

4. Bấm `Analyze workbook`.
5. App sẽ hiển thị danh sách các ô có thể là input/output.
6. Review từng dòng:

- `Approve`: ô đúng, cho phép dùng.
- `Review`: chưa chắc, cần xem lại.
- `Reject`: ô sai, không dùng.

Nếu ô gợi ý sai:

- Nhập sheet đúng vào `Correct Sheet`.
- Nhập ô đúng vào `Correct Cell`.

7. Bấm `Validate selection`.
8. Nếu thành công, bấm `Create profile`.
9. Quay lại tab `Run scenario` để dùng profile mới.

## 5. Cần chú ý khi review mapping

Input là ô app được phép sửa. Input nên là ô nhập liệu thủ công, không phải ô công thức.

Output là ô app chỉ đọc kết quả. Output có thể là ô công thức, ví dụ NPV, IRR, doanh thu, lợi nhuận.

Nếu không chắc một ô có đúng không, để `Review` hoặc hỏi người phụ trách model.

## 6. File nằm ở đâu?

Các file upload nằm trong:

```text
uploads/
```

File scenario được tạo ra nằm trong:

```text
outputs/
```

Log thao tác nằm trong:

```text
logs/audit_log.csv
```

Profile nằm trong:

```text
config/profiles/<ten_profile>/model_map.yaml
```

## 7. Lỗi thường gặp

Nếu app không mở:

- Kiểm tra đã chạy `Start_App.bat` chưa.
- Mở thủ công `http://127.0.0.1:8000`.

Nếu scenario chạy lỗi:

- Kiểm tra Microsoft Excel có mở bình thường không.
- Đóng file Excel đang bị mở ở chế độ protected/read-only nếu cần.
- Kiểm tra profile có đúng với workbook đang dùng không.

Nếu validation báo input là ô công thức:

- Chọn lại ô input thật trong `Correct Cell`.
- Không dùng ô công thức làm input.

Nếu không hiểu lỗi:

- Chụp màn hình lỗi.
- Gửi kèm file Excel, tên profile, và câu lệnh đã nhập cho admin/kỹ thuật.
