# Low-tech Excel Upload Workflow

This version adds a web-based onboarding wizard for new Excel financial models.

## Goal

A low-tech user should not need to edit `model_map.yaml`, run multiple Python scripts, or understand folders like `config/profiles`.

The new workflow is:

```text
Upload Excel
→ Analyze workbook
→ Review suggested input/output cells in the browser
→ Validate selection
→ Create profile
→ Run scenario
```

## How to use

1. Start the app:

```powershell
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

2. Open:

```text
http://127.0.0.1:8000
```

3. Click:

```text
Setup new Excel file
```

4. Upload an `.xlsx` file.

5. Type a simple profile name, for example:

```text
project_a
```

6. Click:

```text
Analyze workbook
```

7. Review the candidates.

For each row:

```text
Approve = use this mapping
Review = not approved yet
Reject = do not use it
```

If the suggested sheet or cell is wrong, fill:

```text
Correct Sheet
Correct Cell
```

8. Click:

```text
Validate selection
```

9. If validation passes, click:

```text
Create profile
```

10. Go back to:

```text
Run scenario
```

The profile and uploaded workbook will already be selected internally. You can now run:

```text
Cho tôi NPV và IRR hiện tại
```

or:

```text
Tạo scenario xấu: TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%
```

## New API endpoints

```text
GET  /api/profiles
POST /api/onboarding/analyze
POST /api/onboarding/validate-selection
POST /api/onboarding/create-profile
```

## Important safety rule

Inputs must not point to formula cells. If a selected input cell contains a formula, validation will fail.

Outputs can be formula cells because the system only reads them.
