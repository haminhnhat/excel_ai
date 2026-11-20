# Excel AI Controller

## Tải App Và Dùng Ngay

**[TẢI APP CHO WINDOWS: ZIP PUBLIC, KHOẢNG 26 MB](https://github.com/haminhnhat/excel_ai/releases/download/public-portable-v0.1.0/Excel_AI_Controller_Public_20261120_155111.zip)**

Không cần biết lập trình, cài Python, tạo môi trường hay chạy lệnh. Bản ZIP này
đã kèm Python và các thư viện cần thiết, không kèm workbook hay dữ liệu khách hàng.

> **Tránh tải nhầm:** không chọn **Code → Download ZIP**, **Source code (zip)**
> hoặc **Source code (tar.gz)**. Đây là mã nguồn dành cho người phát triển,
> không phải bản app mở lên dùng ngay. Hãy dùng link **TẢI APP** ở trên.

**Máy cần có:** Windows 10/11 64-bit và Microsoft Excel bản cài trên máy để tính
lại công thức. Excel Online trên trình duyệt không thay thế được Excel desktop.

### Mở App Trong 3 Bước

1. Bấm **TẢI APP** ở trên để tải file `Excel_AI_Controller_Public_20261120_155111.zip`.
2. Nhấp chuột phải vào ZIP → **Extract All / Giải nén tất cả**. Chọn một thư mục
   như Documents, rồi mở thư mục vừa giải nén. **Không chạy app ngay trong ZIP.**
3. Nhấp đúp **`Excel AI Controller.bat`**. Trình duyệt sẽ tự mở app khi sẵn sàng.

Giữ cửa sổ app đang chạy mở trong lúc sử dụng. Khi dùng xong, nhấn **Ctrl+C**
trong cửa sổ đó để dừng app; chỉ đóng tab trình duyệt thì app vẫn chạy.
Nếu trình duyệt không tự mở, dùng địa chỉ `http://127.0.0.1:...` hiện trong cửa sổ app.

### Thêm File Excel Của Bạn

1. Trong **Setup new Excel file**, chọn file `.xlsx` hoặc `.xlsm` của bạn.
2. Bấm **Analyze workbook**, kiểm tra các ô đầu vào/đầu ra và mapping được đề xuất.
3. Sửa các ô chưa đúng, bấm **Validate selection**, rồi **Create profile**.
4. Chuyển sang **Run scenario** để chạy với file và profile vừa tạo.

Mapping là đề xuất tự động, cần được kiểm tra trước khi chạy. Khi mở lại app,
chọn profile đã lưu và chọn lại workbook trong **Run scenario**.

**Lưu ý dữ liệu:** chỉ chia sẻ ZIP Public nguyên gốc. Không chia sẻ lại thư mục
đã sử dụng vì file tải lên, profile, kết quả và log có thể chứa dữ liệu riêng.

Nếu cần tìm bản phát hành khác, mở [GitHub Releases](https://github.com/haminhnhat/excel_ai/releases)
và chọn file **`Excel_AI_Controller_Public_*.zip`** trong phần **Assets**, không chọn Source code.

---

## Thông Tin Kỹ Thuật

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

### Chạy Từ Mã Nguồn

**Phần này dành cho người phát triển hoặc IT. Nếu đã tải ZIP Public ở đầu trang,
bạn không cần thực hiện các bước dưới đây hoặc chạy `Setup_First_Time.bat`.**

1. Mở thư mục project.
2. Double-click `Setup_First_Time.bat` trong lần đầu tiên.
3. Sau khi setup xong, double-click `Excel AI Controller.lnk` neu co shortcut, hoac `Excel AI Controller.bat`.
4. Trình duyệt sẽ mở app tại:

```text
http://127.0.0.1:8000
```

Giữ cửa sổ màu đen đang chạy app mở trong lúc sử dụng.

Hướng dẫn chi tiết hơn nằm trong:

```text
USER_GUIDE_LOW_TECH.md
```

Nếu admin/IT đã chuẩn bị sẵn offline package cache trong `vendor/wheels`, bước setup sẽ không cần tải thư viện từ internet.

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

## Chuẩn bị bản offline cho nội bộ

Khuyến nghị cho công ty: admin/IT chuẩn bị package một lần trên máy có internet, sau đó copy nguyên folder project cho user.

Trên máy admin có internet:

```text
Build_Offline_Package.bat
```

Script này tải toàn bộ dependency trong `requirements.txt` vào:

```text
vendor/wheels/
```

Sau đó copy cả thư mục project sang máy user. Trên máy user:

```text
Setup_First_Time.bat
Create_Shortcuts.bat
Excel AI Controller.lnk
Excel AI Controller.bat
```

Khi `vendor/wheels` có package, `Setup_First_Time.bat` sẽ ưu tiên cài offline bằng:

```text
pip install --no-index --find-links vendor/wheels -r requirements.txt
```

Như vậy user không cần Git Bash, không cần PowerShell, và thường không cần internet để cài dependencies.

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
- `EXCEL_RECALC_MODE=calculate`: chế độ tính nhanh hơn cho workbook lớn. Chỉ dùng `full_rebuild` khi cần đối soát rất kỹ vì chậm hơn nhiều.
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

## Nâng accuracy cho output

Project có 3 công cụ để đo và cải thiện độ chính xác:

1. Log accuracy khi user chạy scenario:

```text
logs/accuracy_events.jsonl
```

File này lưu câu lệnh user, profile, action plan parser tạo ra, trạng thái validate/chạy Excel, output keys và lỗi nếu có. Admin có thể dùng file này để xem câu nào parse sai và bổ sung alias/test case.

Khi scenario chạy lâu, xem thêm runtime log:

```text
logs/scenario_runtime.log
```

File này ghi các stage như `xlwings:open_workbook`, `xlwings:recalculate`, `xlwings:read_outputs`, kèm số giây đã chạy. Nếu bị treo, stage cuối cùng trong file thường cho biết đang mắc ở bước nào.

2. Chạy eval parser:

```powershell
python scripts/evaluate_parser_accuracy.py
```

Test case mặc định nằm ở:

```text
tests/parser_eval_cases.jsonl
```

Mỗi dòng mô tả một câu lệnh, expected changes và expected outputs. Khi thêm alias hoặc đổi model parser, chạy script này để biết accuracy tăng hay giảm.

3. QA profile/mapping:

```powershell
python scripts/qa_profile_mapping.py --profile yen_my
```

Nếu có workbook thật:

```powershell
python scripts/qa_profile_mapping.py --profile yen_my --excel "models/ten_file.xlsx"
```

Script này kiểm tra các rủi ro như thiếu alias, thiếu min/max, default output sai, sheet/cell không tồn tại, input trỏ vào ô công thức, output rỗng.

Quy trình cải thiện nên là:

```text
1. Chạy QA profile
2. Chạy parser eval
3. Cho user dùng thật
4. Xem logs/accuracy_events.jsonl
5. Thêm alias/test case/correction
6. Chạy eval lại
```

## Quy tắc an toàn

- Không commit file Excel thật lên Git.
- Không commit `.env`.
- Luôn validate profile mới trước khi chạy scenario ghi dữ liệu.
- Chỉ map input vào ô nhập liệu thật, không map vào ô công thức.
- Với mô hình tài chính thật, ưu tiên `xlwings` và Microsoft Excel để tính lại kết quả.
