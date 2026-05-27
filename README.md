# Excel AI Controller

Ứng dụng này giúp người dùng nội bộ chạy mô hình tài chính Excel bằng câu lệnh tự nhiên, ví dụ:

```text
Cho tôi NPV và IRR hiện tại
```

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

Thiết kế an toàn của hệ thống:

```text
Câu lệnh người dùng
→ bộ parse
→ action plan dạng JSON
→ validator kiểm tra quyền sửa ô
→ Excel controller
→ Excel tính lại mô hình
→ trả kết quả và file scenario
```

AI không được tự ý sửa ô bất kỳ trong workbook. Hệ thống chỉ được sửa các ô input đã được duyệt trong `model_map.yaml`.

## Dành cho người dùng không kỹ thuật

Nếu bạn chỉ cần sử dụng app:

1. Mở thư mục project.
2. Double-click `Setup_First_Time.bat` trong lần đầu tiên.
3. Sau khi setup xong, double-click `Start_App.bat`.
4. Trình duyệt sẽ mở app tại:

```text
http://127.0.0.1:8000
```

Giữ cửa sổ màu đen đang chạy app mở trong lúc sử dụng.

Hướng dẫn chi tiết hơn nằm trong:

```text
USER_GUIDE_LOW_TECH.md
```

## Luồng sử dụng chính

### Chạy scenario với file đã có profile

1. Mở app.
2. Chọn tab `Run scenario`.
3. Chọn profile phù hợp.
4. Nhập câu lệnh.
5. Bấm `Run scenario`.
6. Tải file Excel kết quả ở phần download.

Ví dụ câu lệnh:

```text
Cho tôi NPV và IRR hiện tại
```

```text
Tăng lãi vay lên 8%
```

```text
TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

### Setup một file Excel mới

1. Mở app.
2. Chọn tab `Setup new Excel file`.
3. Upload workbook Excel `.xlsx` hoặc `.xlsm`.
4. Nhập tên profile ngắn, ví dụ `project_a`.
5. Bấm `Analyze workbook`.
6. Review các ô hệ thống đề xuất.
7. Chọn `Approve`, `Review`, hoặc `Reject`.
8. Nếu hệ thống đoán sai ô, nhập lại `Correct Sheet` và `Correct Cell`.
9. Bấm `Validate selection`.
10. Nếu validation đạt, bấm `Create profile`.
11. Quay lại tab `Run scenario` để chạy mô hình.

Quy tắc quan trọng: ô input không nên là ô công thức. Ô output có thể là ô công thức vì hệ thống chỉ đọc kết quả.

## Setup cho kỹ thuật/admin

Nếu cần setup thủ công:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Chạy app:

```powershell
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Mở:

```text
http://127.0.0.1:8000
```

## Cấu hình `.env`

Các cấu hình quan trọng:

```env
AI_PROVIDER=mock
EXCEL_ENGINE=xlwings
MODEL_PATH=models/ten_file_excel.xlsx
MODEL_PROFILE=yen_my
OUTPUT_DIR=outputs
AUDIT_LOG_PATH=logs/audit_log.csv
```

Ghi chú:

- `AI_PROVIDER=mock`: dùng parser nội bộ, không cần API key.
- `AI_PROVIDER=openai`: dùng OpenAI làm fallback để hiểu câu lệnh linh hoạt hơn.
- `EXCEL_ENGINE=xlwings`: khuyến nghị cho Windows có Microsoft Excel.
- `openpyxl` có thể sửa ô nhưng không tính lại công thức Excel phức tạp đáng tin cậy.

## Profile mô hình

Mỗi workbook tài chính nên có một profile riêng:

```text
config/profiles/<ten_profile>/model_map.yaml
```

Profile cho biết:

- Ô nào là input được phép sửa.
- Ô nào là output để đọc kết quả.
- Min/max, kiểu dữ liệu, đơn vị và alias cho từng parameter.

Không hardcode ô mới trong parser. Khi có workbook mới, tạo profile mới.

## Workflow review mapping không cần sửa YAML

Nếu không dùng web wizard, admin có thể tạo file review Excel:

```powershell
python scripts/generate_mapping_review.py --excel "models/project_a.xlsx"
```

Mở file:

```text
outputs/analysis/project_a/mapping_review.xlsx
```

Trong sheet `Review Mapping`:

- `Approve`: dùng mapping này.
- `Reject`: bỏ mapping này.
- Nếu sai ô, điền `Correct Sheet` và `Correct Cell/Range`.

Tạo profile từ file review:

```powershell
python scripts/build_profile_from_review.py --review "outputs/analysis/project_a/mapping_review.xlsx" --profile project_a --force
```

Validate profile:

```powershell
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

## Quy tắc an toàn

- Không commit file Excel thật lên Git.
- Không commit `.env`.
- Luôn validate profile mới trước khi chạy scenario ghi dữ liệu.
- Chỉ map input vào ô nhập liệu thật, không map vào ô công thức.
- Với mô hình tài chính thật, ưu tiên `xlwings` và Microsoft Excel để tính lại kết quả.
