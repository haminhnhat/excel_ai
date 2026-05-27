# Excel AI Controller — Hướng dẫn cài đặt cho người mới hoàn toàn

Tài liệu này dành cho người **chưa từng tiếp xúc với code**.

Bạn không cần hiểu sâu về Python. Chỉ cần làm theo từng bước theo thứ tự.

Tài liệu này có hướng dẫn cho cả:

```text
Windows PowerShell
Git Bash
```

Bạn có thể chọn một trong hai. Không nên trộn lẫn lệnh PowerShell và Git Bash nếu chưa quen.

---

# Dự án này dùng để làm gì?

Dự án này cho phép bạn điều khiển một file Excel financial model bằng câu lệnh tự nhiên.

Ví dụ:

```text
Tăng lãi vay lên 8%
```

Hoặc:

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

Hệ thống sẽ:

```text
Đọc câu lệnh của bạn
→ hiểu bạn muốn thay đổi giả định tài chính nào
→ copy file Excel gốc thành một file scenario mới
→ chỉ sửa những ô input đã được cho phép
→ dùng Microsoft Excel để tính lại công thức
→ trả về kết quả như NPV, IRR, lợi nhuận, ROI
```

Hệ thống **không tự ý sửa lung tung trong Excel**. Nó chỉ sửa các ô đã được khai báo trong profile mapping.

---

# 1. Cần cài những gì trước?

## 1.1 Python

Tải Python tại:

```text
https://www.python.org/downloads/
```

Khi cài đặt, nhớ tick lựa chọn:

```text
Add Python to PATH
```

Sau khi cài xong, mở PowerShell hoặc Git Bash và kiểm tra:

### PowerShell

```powershell
python --version
```

### Git Bash

```bash
python --version
```

Nếu cài đúng, bạn sẽ thấy dạng:

```text
Python 3.11.x
```

Khuyến nghị dùng Python 3.11 hoặc mới hơn.

---

## 1.2 Git

Tải Git tại:

```text
https://git-scm.com/downloads
```

Sau khi cài xong, kiểm tra:

### PowerShell

```powershell
git --version
```

### Git Bash

```bash
git --version
```

Nếu cài đúng, bạn sẽ thấy dạng:

```text
git version ...
```

---

## 1.3 Microsoft Excel bản Desktop

Bạn cần cài Microsoft Excel bản desktop thông thường.

Dự án này dùng chính Excel để tính lại công thức như:

```text
NPV
IRR
Cash flow
Công thức liên kết giữa các sheet
```

Chỉ dùng Excel Online là không đủ.

---

# 2. Chọn terminal để chạy lệnh

Bạn có thể dùng một trong hai:

```text
Cách A: PowerShell
Cách B: Git Bash
```

## Khi nào dùng PowerShell?

Dùng PowerShell nếu bạn quen với Windows hơn.

Lệnh PowerShell thường trông như:

```powershell
Copy-Item .env.example .env
.\.venv\Scripts\Activate.ps1
```

## Khi nào dùng Git Bash?

Dùng Git Bash nếu bạn thích kiểu lệnh giống Linux hơn.

Lệnh Git Bash thường trông như:

```bash
cp .env.example .env
source .venv/Scripts/activate
```

---

# 3. Tải project từ GitHub

## 3.1 Chọn một thư mục để lưu project

Bạn có thể chọn bất kỳ thư mục nào, ví dụ:

```text
Desktop
Documents
```

---

## 3.2 Clone repository từ GitHub

Thay `YOUR_GITHUB_REPO_URL` bằng link GitHub thật của project.

### PowerShell

```powershell
cd Desktop
git clone YOUR_GITHUB_REPO_URL
cd excel-ai-controller
```

Ví dụ:

```powershell
cd Desktop
git clone https://github.com/your-name/excel-ai-controller.git
cd excel-ai-controller
```

### Git Bash

```bash
cd ~/Desktop
git clone YOUR_GITHUB_REPO_URL
cd excel-ai-controller
```

Ví dụ:

```bash
cd ~/Desktop
git clone https://github.com/your-name/excel-ai-controller.git
cd excel-ai-controller
```

Nếu folder repo của bạn có tên khác, hãy dùng đúng tên folder đó trong lệnh `cd`.

---

# 4. Tạo môi trường Python cho project

## 4.1 Tạo virtual environment

### PowerShell

```powershell
python -m venv .venv
```

### Git Bash

```bash
python -m venv .venv
```

Lệnh này tạo một môi trường Python riêng cho project.

---

## 4.2 Kích hoạt virtual environment

### PowerShell

```powershell
.\.venv\Scripts\Activate.ps1
```

Nếu PowerShell chặn không cho chạy script, chạy lệnh này một lần:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Sau đó activate lại:

```powershell
.\.venv\Scripts\Activate.ps1
```

### Git Bash

```bash
source .venv/Scripts/activate
```

Nếu activate thành công, bạn sẽ thấy:

```text
(.venv)
```

ở đầu dòng terminal.

Ví dụ:

```text
(.venv) user@computer MINGW64 ~/Desktop/excel-ai-controller
```

---

## 4.3 Cài các thư viện cần thiết

### PowerShell

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Git Bash

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Nếu cần cài thêm `xlwings` hoặc `pytest`:

### PowerShell

```powershell
python -m pip install xlwings pytest
```

### Git Bash

```bash
python -m pip install xlwings pytest
```

---

# 5. Chuẩn bị các thư mục cần thiết

Project nên có các thư mục sau:

```text
models/
uploads/
outputs/
logs/
config/profiles/
```

Nếu thiếu thư mục nào, tạo bằng lệnh sau.

### PowerShell

```powershell
New-Item models -ItemType Directory -Force
New-Item uploads -ItemType Directory -Force
New-Item outputs -ItemType Directory -Force
New-Item logs -ItemType Directory -Force
```

### Git Bash

```bash
mkdir -p models uploads outputs logs
```

Các thư mục này thường bị Git ignore vì có thể chứa file Excel riêng tư hoặc file output được tạo ra khi chạy scenario.

---

# 6. Thêm file Excel financial model

Copy file Excel financial model của bạn vào thư mục:

```text
models/
```

Ví dụ:

```text
models/project_a.xlsx
```

Tên file khuyến nghị:

```text
project_a.xlsx
yen_my.xlsx
financial_model.xlsx
```

Nên tránh ký tự đặc biệt nếu có thể.

Tốt:

```text
project_a.xlsx
```

Ít phù hợp cho người mới:

```text
2025.10.26_NOXH_Yên Mỹ_BKD.xlsx
```

Tên có dấu tiếng Việt vẫn có thể chạy, nhưng tên đơn giản sẽ ít lỗi hơn.

---

# 7. Tạo file `.env`

File `.env` cho app biết file Excel nào và profile nào sẽ được dùng mặc định.

## 7.1 Copy `.env.example`

### PowerShell

```powershell
Copy-Item .env.example .env
```

### Git Bash

```bash
cp .env.example .env
```

---

## 7.2 Mở file `.env`

### PowerShell

```powershell
notepad .env
```

### Git Bash

```bash
notepad .env
```

Bạn cũng có thể mở `.env` bằng VS Code hoặc bất kỳ text editor nào.

---

## 7.3 Ví dụ nội dung `.env`

Giả sử file Excel tên là `project_a.xlsx` và profile tên là `project_a`:

```env
AI_PROVIDER=mock
EXCEL_ENGINE=xlwings
MODEL_PROFILE=project_a
MODEL_PATH=models/project_a.xlsx
MODEL_MAP_PATH=config/profiles/project_a/model_map.yaml
OUTPUT_DIR=outputs
AUDIT_LOG_PATH=logs/audit_log.csv
```

Giải thích:

```text
AI_PROVIDER=mock
```

Không cần API key. Hệ thống dùng parser local.

```text
EXCEL_ENGINE=xlwings
```

Hệ thống dùng Microsoft Excel Desktop để tính công thức.

```text
MODEL_PROFILE=project_a
```

Hệ thống dùng profile mapping tên `project_a`.

```text
MODEL_PATH=models/project_a.xlsx
```

File Excel mặc định nằm ở `models/project_a.xlsx`.

Không upload file `.env` lên GitHub.

---

# 8. Profile là gì?

Profile cho hệ thống biết các ô quan trọng nằm ở đâu trong file Excel.

Ví dụ profile sẽ nói:

```text
Lãi vay nằm ở ô nào
NPV dự án nằm ở ô nào
IRR dự án nằm ở ô nào
VAT nằm ở ô nào
Thuế TNDN nằm ở ô nào
```

Profile nằm trong thư mục:

```text
config/profiles/
```

Ví dụ:

```text
config/profiles/project_a/model_map.yaml
config/profiles/yen_my/model_map.yaml
```

Nếu file Excel của bạn là:

```text
project_a.xlsx
```

thì nên dùng profile tên:

```text
project_a
```

---

# 9. Nếu đã có profile sẵn

Nếu đã có file:

```text
config/profiles/project_a/model_map.yaml
```

hãy validate trước khi chạy scenario.

### PowerShell

```powershell
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

### Git Bash

```bash
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

Kết quả đúng sẽ có:

```json
"ok": true
```

Nếu có lỗi, cần sửa profile trước khi chạy.

---

# 10. Nếu chưa có profile

Nếu bạn đang dùng một file Excel mới và chưa có profile, dùng workflow tạo mapping review.

Workflow này giúp hệ thống tạo một file Excel review để người kiểm tra không cần hiểu code.

---

## 10.1 Tạo file mapping review

### PowerShell

```powershell
python scripts/generate_mapping_review.py --excel "models/project_a.xlsx"
```

### Git Bash

```bash
python scripts/generate_mapping_review.py --excel "models/project_a.xlsx"
```

Lệnh này sẽ tạo file dạng:

```text
outputs/analysis/project_a/mapping_review.xlsx
```

---

## 10.2 Mở file review

Mở file:

```text
outputs/analysis/project_a/mapping_review.xlsx
```

Vào sheet:

```text
Review Mapping
```

Bạn sẽ thấy các cột như:

```text
Decision
Role
Parameter Key
Friendly Name
Sheet
Cell
Current Value
Formula?
Confidence
Correct Sheet
Correct Cell/Range
```

---

## 10.3 Approve hoặc Reject từng dòng

Trong cột `Decision`, dùng:

```text
Approve = mapping này đúng
Reject = mapping này sai
Review = chưa duyệt
```

Nếu hệ thống đoán sai sheet hoặc cell, điền vào:

```text
Correct Sheet
Correct Cell/Range
```

Ví dụ:

```text
Correct Sheet: Assumptions
Correct Cell/Range: C8
```

Người review không cần biết Python hoặc YAML. Chỉ cần biết ô Excel được đề xuất có đúng hay không.

---

## 10.4 Build profile từ file review

Sau khi save file review:

### PowerShell

```powershell
python scripts/build_profile_from_review.py --review "outputs/analysis/project_a/mapping_review.xlsx" --profile project_a --force
```

### Git Bash

```bash
python scripts/build_profile_from_review.py --review "outputs/analysis/project_a/mapping_review.xlsx" --profile project_a --force
```

Lệnh này sẽ tạo:

```text
config/profiles/project_a/model_map.yaml
```

---

## 10.5 Validate profile mới

### PowerShell

```powershell
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

### Git Bash

```bash
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

Chỉ tiếp tục khi kết quả có:

```json
"ok": true
```

---

# 11. Chạy command bằng terminal

## 11.1 Đọc NPV và IRR hiện tại

### PowerShell

```powershell
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Cho tôi NPV và IRR hiện tại" --engine xlwings
```

### Git Bash

```bash
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Cho tôi NPV và IRR hiện tại" --engine xlwings
```

---

## 11.2 Chạy scenario thay đổi lãi vay

### PowerShell

```powershell
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Tăng lãi vay lên 8%" --engine xlwings
```

### Git Bash

```bash
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Tăng lãi vay lên 8%" --engine xlwings
```

---

## 11.3 Chạy downside scenario

### PowerShell

```powershell
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%" --engine xlwings
```

### Git Bash

```bash
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%" --engine xlwings
```

---

# 12. Chạy web app

### PowerShell

```powershell
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

### Git Bash

```bash
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Sau đó mở trình duyệt:

```text
http://127.0.0.1:8000
```

Trong web page:

```text
Excel file: chọn file Excel, hoặc để trống để dùng MODEL_PATH trong .env
Profile: nhập tên profile, ví dụ project_a
Engine: xlwings
Command: Cho tôi NPV và IRR hiện tại
```

Bấm:

```text
Run scenario
```

---

# 13. Các command mẫu có thể dùng

## Đọc output

```text
Cho tôi NPV và IRR hiện tại
Xem NPV dự án
Xem IRR dự án
Cho tôi lợi nhuận sau thuế
Cho tôi ROI
Cho tôi NPV chủ đầu tư và IRR chủ đầu tư
```

## Thay đổi lãi vay

```text
Tăng lãi vay lên 8%
Cho lãi ngân hàng là 8%
Lãi suất vay 8.5%
Interest rate 9%
Chi phí vốn vay 7.5%
```

## Thay đổi tổng mức đầu tư / TMĐT

```text
Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ
Tăng TMĐT từ 1000 tỷ lên 1200 tỷ
Tăng tổng mức đầu tư 20%
TMĐT tăng 15%
Chi phí đầu tư tăng 10%
Tổng mức đầu tư giảm 5%
```

## Thay đổi giá bán

```text
Giảm giá bán 5%
Giá bán tăng 10%
Đơn giá bán giảm 7%
Nếu giá bán giảm 5% thì sao
```

## Thay đổi thuế

```text
VAT 10%
Thuế VAT 8%
TNDN 20%
Thuế TNDN 20%
CIT 20%
VAT 10% TNDN 20%
```

## Thay đổi chi phí

```text
Chi phí bán hàng 3%
Hoa hồng 2%
Môi giới 2.5%
Marketing 2%
Chi phí marketing 2%
Truyền thông 1.5%
Chi phí quản lý 4%
Chi phí quản lý doanh nghiệp 4%
G&A 3%
```

## Scenario nhiều biến

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
TMĐT tăng 10%, giá bán giảm 3%, lãi vay 8.5%
Tổng mức đầu tư tăng 20%, giá bán giảm 5%, VAT 10%, TNDN 20%, lãi vay 8%
```

---

# 14. Những command hệ thống nên từ chối

Các command này quá mơ hồ hoặc không an toàn:

```text
Làm model tốt hơn
Chỉnh sao cho IRR cao hơn
Tối ưu lợi nhuận
Sửa công thức giúp tôi
Làm NPV dương
Tự tìm chỗ cần chỉnh
```

Hệ thống nên từ chối thay vì tự ý sửa file Excel.

---

# 15. File output nằm ở đâu?

Các file scenario mới sẽ được lưu trong:

```text
outputs/
```

File gốc trong:

```text
models/
```

không nên bị thay đổi.

Mỗi lần chạy sẽ tạo một file scenario mới.

---

# 16. An toàn khi dùng GitHub

Trước khi push code lên GitHub, luôn chạy:

### PowerShell

```powershell
git status
```

### Git Bash

```bash
git status
```

Không commit các file sau:

```text
.env
models/*.xlsx
uploads/*.xlsx
outputs/*.xlsx
logs/*.csv
```

Các file này có thể chứa dữ liệu nhạy cảm.

File `.gitignore` nên có:

```gitignore
.env

models/*
!models/.gitkeep

uploads/*
!uploads/.gitkeep

outputs/*
!outputs/.gitkeep

logs/*
!logs/.gitkeep

*.xlsx
*.xlsm
*.xls
~$*.xlsx

.venv/
__pycache__/
*.pyc
```

Nếu lỡ stage file Excel:

### PowerShell

```powershell
git rm --cached "models/file_name.xlsx"
```

### Git Bash

```bash
git rm --cached "models/file_name.xlsx"
```

Commit các file an toàn:

### PowerShell

```powershell
git add app config scripts web tests README.md requirements.txt .env.example .gitignore pytest.ini
git add models/.gitkeep uploads/.gitkeep outputs/.gitkeep logs/.gitkeep
git commit -m "Update Excel AI Controller"
git push
```

### Git Bash

```bash
git add app config scripts web tests README.md requirements.txt .env.example .gitignore pytest.ini
git add models/.gitkeep uploads/.gitkeep outputs/.gitkeep logs/.gitkeep
git commit -m "Update Excel AI Controller"
git push
```

---

# 17. Lỗi thường gặp

## Lỗi: `python` is not recognized

Python chưa được cài đúng, hoặc chưa được thêm vào PATH.

Cách xử lý:

```text
Cài lại Python
Tick "Add Python to PATH"
Mở lại terminal
```

---

## Lỗi: Git Bash không tìm thấy Python

Thử:

```bash
py --version
```

Nếu `py` chạy được nhưng `python` không chạy được, dùng:

```bash
py -m venv .venv
```

Sau đó activate:

```bash
source .venv/Scripts/activate
```

Sau khi activate, thử lại:

```bash
python --version
```

---

## Lỗi: PowerShell không activate được `.venv`

Chạy:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Sau đó:

```powershell
.\.venv\Scripts\Activate.ps1
```

---

## Lỗi: Git Bash activate không được

Dùng đúng lệnh:

```bash
source .venv/Scripts/activate
```

Không dùng lệnh PowerShell trong Git Bash:

```text
.\.venv\Scripts\Activate.ps1
```

Lệnh đó chỉ dành cho PowerShell.

---

## Lỗi: `No module named pytest`

### PowerShell

```powershell
python -m pip install pytest
```

### Git Bash

```bash
python -m pip install pytest
```

---

## Lỗi: `No module named app`

Hãy chắc chắn bạn đang ở đúng folder project.

Bạn nên thấy các thư mục:

```text
app
config
scripts
web
tests
```

Sau đó chạy:

### PowerShell

```powershell
python -m pytest -v
```

### Git Bash

```bash
python -m pytest -v
```

Nếu cần, tạo file `pytest.ini`:

```ini
[pytest]
pythonpath = .
testpaths = tests
```

---

## Lỗi: `FileNotFoundError`

Đường dẫn file Excel bị sai.

Kiểm tra file trong thư mục `models`.

### PowerShell

```powershell
dir models
```

### Git Bash

```bash
ls models
```

Dùng đúng tên file trong command.

---

## Lỗi: output bị trống khi dùng `openpyxl`

Dùng:

```text
--engine xlwings
```

`openpyxl` không tính lại công thức Excel.

---

## Lỗi: `Refusing to overwrite formula cell`

Profile đang trỏ input vào một ô công thức.

Việc này bị chặn để tránh phá model.

Cách sửa: dùng đúng ô input, hoặc tạo lại profile thông qua:

```text
mapping_review.xlsx
```

---

## Lỗi: Excel bị treo

Đóng toàn bộ cửa sổ Excel.

### PowerShell

```powershell
taskkill /F /IM EXCEL.EXE
```

### Git Bash

```bash
taskkill //F //IM EXCEL.EXE
```

Sau đó chạy lại.

---

## Lỗi: web UI không dùng đúng profile

Nhập thủ công profile trong ô Profile.

Ví dụ:

```text
project_a
```

Restart backend:

### PowerShell

```powershell
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

### Git Bash

```bash
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Refresh trình duyệt bằng:

```text
Ctrl + F5
```

---

# 18. Checklist cho máy mới hoàn toàn

Dùng checklist này khi setup trên một máy mới:

```text
[ ] Cài Python
[ ] Cài Git
[ ] Cài Microsoft Excel Desktop
[ ] Clone repository từ GitHub
[ ] Mở PowerShell hoặc Git Bash trong folder project
[ ] Tạo .venv
[ ] Activate .venv
[ ] Cài requirements
[ ] Copy file Excel vào models/
[ ] Copy .env.example thành .env
[ ] Sửa .env
[ ] Tạo hoặc validate profile
[ ] Test CLI: "Cho tôi NPV và IRR hiện tại"
[ ] Test scenario
[ ] Chạy web app
[ ] Test trên browser
```

---

# 19. Tóm tắt thiết kế hệ thống

```text
User command
→ parser
→ action plan
→ validator
→ Excel controller
→ copied scenario workbook
→ Excel recalculation
→ output reader
→ result
```

Nguyên tắc chính:

```text
Ngôn ngữ tự nhiên có thể linh hoạt.
Việc chỉnh Excel phải được kiểm soát.
```

Hệ thống không nên cho AI hoặc parser tự do sửa bất kỳ ô Excel nào.
