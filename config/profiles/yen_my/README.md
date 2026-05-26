# Profile: yen_my

This profile maps the uploaded NOXH Yên Mỹ financial model. It is safe only for workbooks with the same sheet/cell structure.

Use it with:

```powershell
python scripts/run_cli.py --profile yen_my --excel "models/2025.10.26_NOXH_Yen_My_BKD.xlsx" --command "Tăng lãi vay lên 8%" --engine xlwings
```

For another workbook, create a new profile under `config/profiles/<profile_name>/model_map.yaml`.
