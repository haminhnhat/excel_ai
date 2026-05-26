# Excel AI Controller

This project controls an Excel financial model through natural-language commands. The safe design is:

```text
User command -> parser -> JSON action -> validator -> Excel controller -> recalculation -> mapped outputs
```

The AI/parser never edits arbitrary workbook cells. It can only modify input parameters listed in a trusted `model_map.yaml`.

## Quick start

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Put the real Excel workbook in `models/`, for example:

```text
models/2025.10.26_NOXH_Yen_My_BKD.xlsx
```

Run the web app:

```powershell
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Open:

```text
http://127.0.0.1:8000
```

## CLI examples

```powershell
python scripts/run_cli.py --profile yen_my --excel "models/2025.10.26_NOXH_Yen_My_BKD.xlsx" --command "Tăng lãi vay lên 8%" --engine xlwings
```

```powershell
python scripts/run_cli.py --profile yen_my --excel "models/2025.10.26_NOXH_Yen_My_BKD.xlsx" --command "Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%" --engine xlwings
```

## Model profiles

Different financial Excel workbooks need different mappings. Do not hardcode new workbook cells in the parser.

Use profiles:

```text
config/profiles/<profile_name>/model_map.yaml
```

Example:

```text
config/profiles/yen_my/model_map.yaml
config/profiles/project_a/model_map.yaml
```

Set profile in `.env`:

```env
MODEL_PROFILE=yen_my
```

Or pass it in CLI:

```powershell
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Cho tôi NPV và IRR hiện tại" --engine xlwings
```

## Onboard a new financial Excel workbook

1. Put the workbook in `models/`, for example:

```text
models/project_a.xlsx
```

2. Analyze the workbook:

```powershell
python scripts/analyze_workbook.py --excel "models/project_a.xlsx"
```

This creates:

```text
outputs/analysis/project_a/
  workbook_structure.json
  candidate_inputs.csv
  candidate_outputs.csv
  formula_cells.csv
  formula_errors.csv
  draft_model_map.yaml
  analysis_summary.json
```

3. Create a profile from the draft:

```powershell
python scripts/create_profile.py --profile project_a --draft-map "outputs/analysis/project_a/draft_model_map.yaml"
```

4. Manually edit:

```text
config/profiles/project_a/model_map.yaml
```

Verify every input and output cell. Do not map inputs to formula cells unless you intentionally set `allow_formula_overwrite: true`.

5. Validate the profile:

```powershell
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

6. Test read-only command first:

```powershell
python scripts/run_cli.py --profile project_a --excel "models/project_a.xlsx" --command "Cho tôi NPV và IRR hiện tại" --engine xlwings
```

7. Then test one write command at a time.

## Mapping format

Single input cell:

```yaml
inputs:
  loan_interest_rate:
    sheet: "Assumptions"
    cell: "C12"
    type: "percent"
    editable: true
    min: 0.0
    max: 0.25
    unit: "decimal"
    aliases:
      - "lãi vay"
      - "interest rate"
```

Multiple hardcoded cells for one assumption:

```yaml
inputs:
  loan_interest_rate:
    sheet: "Loan"
    cells: ["E7", "F7", "G7", "H7"]
    type: "percent"
    editable: true
    min: 0.0
    max: 0.25
    unit: "decimal"
```

Range:

```yaml
inputs:
  loan_interest_rate:
    sheet: "Loan"
    range: "E7:AS7"
    type: "percent"
    editable: true
    min: 0.0
    max: 0.25
    unit: "decimal"
```

Outputs:

```yaml
outputs:
  project_npv:
    sheet: "Summary"
    cell: "D6"
    type: "currency"
    unit: "VND"
    aliases:
      - "npv dự án"
      - "project npv"
```

## Important safety rules

- Keep real Excel models out of GitHub.
- Keep `.env` out of GitHub.
- Use `xlwings` on Windows with Microsoft Excel for reliable financial-model recalculation.
- `openpyxl` can edit cells but does not calculate formulas like Excel.
- Always validate a new profile before running write commands.
