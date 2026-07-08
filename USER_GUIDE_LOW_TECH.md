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

Nếu admin/IT đã chuẩn bị sẵn bản offline, user không cần tự tải thư viện Python từ internet. Chỉ cần chạy file `.bat` theo hướng dẫn bên dưới.

## 2. Mở app lần đầu

1. Mở thư mục app.
2. Double-click `Setup_First_Time.bat`.
3. Chờ đến khi thấy dòng:

```text
Setup complete.
```

4. Double-click `Excel AI Controller.lnk` neu co shortcut, hoac `Excel AI Controller.bat`.
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

- Kiểm tra đã chạy `Excel AI Controller.lnk` hoặc `Excel AI Controller.bat` chưa.
- Mở thủ công `http://127.0.0.1:8000`.

Nếu scenario chạy lỗi:

- Kiểm tra Microsoft Excel có mở bình thường không.
- Đóng file Excel đang bị mở ở chế độ protected/read-only nếu cần.
- Kiểm tra profile có đúng với workbook đang dùng không.

Nếu scenario chạy lâu:

- Workbook lớn có thể mất vài phút.
- Xem đồng hồ ở góc `Scenario result`.
- Nếu quá 5 phút, báo admin/kỹ thuật kiểm tra `logs/scenario_runtime.log`.

Nếu validation báo input là ô công thức:

- Chọn lại ô input thật trong `Correct Cell`.
- Không dùng ô công thức làm input.

Nếu không hiểu lỗi:

- Chụp màn hình lỗi.
- Gửi kèm file Excel, tên profile, và câu lệnh đã nhập cho admin/kỹ thuật.

## 8. Ghi chú cho admin/IT

Để giảm lỗi khi setup trên nhiều máy, admin/IT nên chuẩn bị bản offline trước.

Trên một máy có internet, chạy:

```text
Build_Offline_Package.bat
```

Sau khi chạy xong, thư mục sau sẽ có các package cần thiết:

```text
vendor/wheels/
```

Copy toàn bộ thư mục project sang máy user. User chỉ cần chạy:

```text
Setup_First_Time.bat
```

Nếu setup thấy package trong `vendor/wheels`, app sẽ cài offline và không cần tải package từ internet.

## 9. Ghi chú nâng độ chính xác

Khi user chạy scenario, app tự ghi log kỹ thuật vào:

```text
logs/accuracy_events.jsonl
```

Nếu user báo app hiểu sai câu lệnh, admin/kỹ thuật nên lấy:

- Câu lệnh user đã nhập.
- Tên profile.
- File Excel đang dùng.
- Dòng log tương ứng trong `logs/accuracy_events.jsonl`.

Sau đó có thể bổ sung alias hoặc test case để lần sau app hiểu đúng hơn.

## 10. Thay số liệu từ file đã chỉnh vào file gốc

Dùng phần này khi bạn có:

- Một file gốc cần giữ format và cấu trúc.
- Một file đã được chỉnh số liệu, ví dụ file xuất ra sau khi chạy scenario hoặc câu lệnh.

Cách dùng:

1. Mở tab `Run scenario`.
2. Tìm phần `Thay số liệu vào file gốc`.
3. Chọn `File gốc cần giữ format`.
4. Chọn `File đã chỉnh số liệu`.
5. Nếu chỉ muốn thay một sheet, nhập tên sheet. Nếu để trống, app tự tìm các sheet trùng tên.
6. Bấm `Preview changes` để xem số ô sẽ được thay.
7. Nếu preview đúng, bấm `Export safe copy`.
8. Tải file Excel mới ở link download.

App không ghi đè file gốc. App luôn tạo một bản copy mới trong thư mục `outputs/`.

File xuất là bản copy đầy đủ của file gốc. Các sheet, format, công thức và nội dung không nằm trong danh sách thay đổi sẽ được giữ từ file gốc.

Mặc định app sẽ bỏ qua ô công thức để tránh làm hỏng mô hình. Nếu muốn tạo một bản snapshot số liệu cố định, có thể chọn `Cho phép thay ô công thức bằng value`. Khi bật lựa chọn này, các ô công thức được thay sẽ trở thành số cố định trong file xuất mới.

## 11. Chuẩn hóa format workbook

Dùng phần này khi file khách hàng có cùng loại dữ liệu nhưng thứ tự cột hoặc tên cột không giống template chuẩn.

Ví dụ template chuẩn là:

```text
a | b | c
```

File khách hàng là:

```text
a | c | b
```

App sẽ đọc header, dự đoán cột nào tương ứng với cột chuẩn, rồi xuất file mới theo thứ tự của template.

Cách dùng:

1. Mở tab `Run scenario`.
2. Tìm phần `Chuẩn hóa format workbook`.
3. Chọn `File format rác / khách hàng`.
4. Chọn `File template chuẩn`.
5. Nếu chỉ muốn chuẩn hóa một sheet, nhập tên sheet. Nếu để trống, app tự match sheet gần giống.
6. Bấm `Preview mapping`.
7. Kiểm tra các cột có confidence thấp hoặc action là `Review`.
8. Nếu mapping ổn, bấm `Export normalized file`.

App không ghi đè file khách hàng hoặc template chuẩn. File mới được tạo trong thư mục `outputs/`.

Nếu tiêu đề không giống 100%, app sẽ dùng normalize và fuzzy matching để dự đoán. Với confidence thấp, admin/kỹ thuật nên kiểm tra lại trước khi dùng file output.
