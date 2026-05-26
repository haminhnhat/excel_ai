# Excel AI Controller — Financial Model Agent

Dự án mẫu này biến file Excel financial model thành một **AI-controlled scenario engine**.

Người dùng nhập câu lệnh tự nhiên, ví dụ:

```text
Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ và cho tôi NPV, IRR mới.
```

Hệ thống sẽ:

1. Parse câu lệnh thành JSON action plan.
2. Validate các parameter được phép sửa.
3. Copy file Excel gốc thành file scenario mới.
4. Sửa đúng input cell trong Excel.
5. Recalculate workbook.
6. Đọc các output cell như lợi nhuận sau thuế, NPV, IRR.
7. Trả kết quả và file Excel scenario mới.

## Cấu trúc

```text
excel_ai_controller_project/
├── app/
│   ├── main.py                # FastAPI backend
│   ├── ai_parser.py           # AI / fallback parser
│   ├── excel_controller.py    # Sửa Excel + recalculate + đọc output
│   ├── validator.py           # Guardrail cho input/output
│   ├── config_loader.py       # Load YAML map
│   ├── schemas.py             # Pydantic schemas
│   └── utils.py
├── config/
│   └── model_map.yaml         # Mapping input/output cho workbook
├── models/
│   └── .gitkeep               # Đặt Excel gốc ở đây
├── outputs/                   # File scenario xuất ra
├── logs/                      # Audit log CSV
├── scripts/
│   ├── run_cli.py             # Chạy không cần web server
│   └── inspect_workbook.py    # Inspect label/cell trong workbook
├── web/
│   └── index.html             # UI đơn giản
├── tests/
│   └── test_fallback_parser.py
├── .env.example
└── requirements.txt
```

## Cài đặt

```bash
cd excel_ai_controller_project
python -m venv .venv
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

macOS/Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

Copy file Excel model vào thư mục `models/`, ví dụ:

```text
models/2025.10.26_NOXH_Yen_My_BKD.xlsx
```

Sau đó copy `.env.example` thành `.env` và sửa:

```text
MODEL_PATH=models/2025.10.26_NOXH_Yen_My_BKD.xlsx
MODEL_MAP_PATH=config/model_map.yaml
EXCEL_ENGINE=xlwings
AI_PROVIDER=mock
```

## Chạy bằng API + web UI

```bash
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Mở trình duyệt:

```text
http://127.0.0.1:8000
```

Nhập prompt, ví dụ:

```text
Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ. Trả ra lợi nhuận sau thuế, NPV dự án, IRR dự án, NPV chủ đầu tư, IRR chủ đầu tư.
```

## Chạy bằng CLI

```bash
python scripts/run_cli.py \
  --excel "models/2025.10.26_NOXH_Yen_My_BKD.xlsx" \
  --command "Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ" \
  --engine xlwings
```

## Dùng AI API thật

Mặc định project dùng parser nội bộ đơn giản (`AI_PROVIDER=mock`) để demo các câu như:

```text
Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ
Tăng lãi vay lên 8%
Giảm giá bán 5%
```

Muốn dùng AI API thật, chỉnh `.env`:

```text
AI_PROVIDER=openai
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4o-mini
```

Phần AI chỉ được trả về JSON action plan. Nó không được tự chọn cell ngoài `config/model_map.yaml`.

## Engine Excel

Khuyến nghị dùng:

```text
EXCEL_ENGINE=xlwings
```

Vì `xlwings` dùng Microsoft Excel thật để recalculate workbook. Cách này phù hợp nhất cho financial model có công thức phức tạp như IRR, NPV, XNPV, XIRR.

Fallback:

```text
EXCEL_ENGINE=openpyxl
```

`openpyxl` có thể sửa cell và save file, nhưng **không tính lại công thức Excel một cách đáng tin cậy**. Output có thể là cached value cũ cho tới khi mở bằng Excel.

## Mapping hiện tại

`config/model_map.yaml` đã được tạo theo file workbook đã gửi. Một số mapping quan trọng:

| Parameter | Sheet | Cell | Ghi chú |
|---|---|---:|---|
| `investment_cost_change` | `1. Tổng hợp` | `C38` | Tỷ lệ thay đổi TMĐT |
| `selling_price_change` | `1. Tổng hợp` | `C37` | Tỷ lệ thay đổi giá bán |
| `loan_interest_rate` | `1. Tổng hợp` | `D26` | Lãi vay ngân hàng |
| `vat_rate` | `1. Tổng hợp` | `D24` | Thuế VAT |
| `cit_rate` | `1. Tổng hợp` | `D25` | Thuế TNDN |
| `project_irr` | `1. Tổng hợp` | `D5` | IRR dự án |
| `project_npv` | `1. Tổng hợp` | `D6` | NPV dự án |
| `equity_irr` | `1. Tổng hợp` | `D7` | IRR chủ đầu tư |
| `equity_npv` | `1. Tổng hợp` | `D8` | NPV chủ đầu tư |
| `profit_after_tax` | `7. P&L TOÀN DỰ ÁN` | `C24` | Lợi nhuận sau thuế |

Với câu:

```text
Tăng tổng mức đầu tư từ 1000 tỷ lên 1200 tỷ
```

parser sẽ hiểu là tăng TMĐT thêm:

```text
1200 / 1000 - 1 = 20%
```

và set `investment_cost_change = 0.2` vào `1. Tổng hợp!C38`.

## Guardrail

Project có các guardrail cơ bản:

- Chỉ sửa parameter nằm trong `inputs` của `model_map.yaml`.
- Không sửa formula cell nếu `allow_formula_overwrite=false`.
- Validate min/max.
- Luôn copy workbook gốc trước khi sửa.
- Ghi audit log vào `logs/audit_log.csv`.

## Lưu ý quan trọng

Trước khi dùng cho số liệu thật, cần kiểm tra lại workbook gốc có lỗi `#REF!`, `#VALUE!`, `#DIV/0!` hay không. Nếu model gốc đã có lỗi công thức, AI chỉnh input đúng nhưng output vẫn có thể sai.
