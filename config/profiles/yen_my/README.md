# Profile: yen_my

Profile này map các ô input/output cho mô hình tài chính NOXH Yên Mỹ.

Chỉ dùng profile này với workbook có cùng cấu trúc sheet và cell. Nếu workbook khác cấu trúc, hãy tạo profile mới bằng tab `Setup new Excel file` trong app.

Ví dụ chạy bằng command line cho admin/kỹ thuật:

```powershell
python scripts/run_cli.py --profile yen_my --excel "models/2025.10.26_NOXH_Yen_My_BKD.xlsx" --command "Tăng lãi vay lên 8%" --engine xlwings
```

Người dùng không kỹ thuật nên mở app bằng `Start_App.bat` và chạy scenario trên trình duyệt.
