# Profile: project_a

Đây là profile mẫu cho workbook test `project_a.xlsx`.

Profile này chủ yếu dùng để kiểm thử luồng tạo mapping và chạy scenario. Không dùng cho workbook thật nếu cấu trúc sheet/cell không giống file test.

Validate bằng command line:

```powershell
python scripts/validate_profile.py --profile project_a --excel "models/project_a.xlsx"
```

Người dùng không kỹ thuật nên dùng giao diện web thay vì chạy lệnh.
