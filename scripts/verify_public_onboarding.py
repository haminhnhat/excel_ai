"""Exercise first-run onboarding against a running public portable package."""
from __future__ import annotations

import argparse
import hashlib
import io
from pathlib import Path

import httpx
from openpyxl import Workbook, load_workbook


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True)
    parser.add_argument("--work-dir", type=Path, required=True)
    args = parser.parse_args()
    args.work_dir.mkdir(parents=True, exist_ok=True)
    source = args.work_dir / "user_workbook.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws.append(["Lai vay", .05])
    ws.append(["NPV", "=B1*1000"])
    wb.save(source)
    wb.close()
    original = hashlib.sha256(source.read_bytes()).hexdigest()

    with httpx.Client(base_url=args.url, timeout=90, trust_env=False) as client:
        health = client.get("/health").raise_for_status().json()
        assert health["available_profiles"] == [], health
        assert not health["model_profile"]
        response = client.post("/api/onboarding/analyze", files={"file": (source.name, source.read_bytes())},
                               data={"suggested_profile": "public_smoke"}).raise_for_status().json()
        assert response["ok"] and response["summary"]["sheet_count"] == 1
        assert response["candidates"]
        print("PASS: Empty public app accepts a new workbook and analyzes its sheet", flush=True)
        payload = {
            "profile_name": "public_smoke", "excel_path": response["excel_path"],
            "candidates": [
                {"decision": "Approve", "role": "input", "parameter_key": "loan_interest_rate",
                 "sheet": "Model", "cell": "B1", "type": "percent", "min": 0, "max": 1,
                 "aliases": ["lai vay"]},
                {"decision": "Approve", "role": "output", "parameter_key": "project_npv",
                 "sheet": "Model", "cell": "B2", "type": "number"},
            ],
        }
        for endpoint in ("validate-selection", "create-profile"):
            result = client.post("/api/onboarding/" + endpoint, json=payload).raise_for_status().json()
            assert result["ok"], result
        assert "public_smoke" in client.get("/health").json()["available_profiles"]
        print("PASS: User-reviewed mapping validates and saves a new profile", flush=True)

        data = {"command": "Tang lai vay len 8%", "model_profile": "public_smoke", "engine": "xlwings",
                "excel_path": response["excel_path"]}
        result = client.post("/api/scenario", data=data).raise_for_status().json()
        assert result["ok"] and result["result"]["recalculated"], result
        assert result["result"]["outputs"]["project_npv"]["value"] == 80
        output = client.get(result["download_url"]).raise_for_status().content
        wb = load_workbook(io.BytesIO(output), data_only=True)
        assert wb["Model"]["B1"].value == .08 and wb["Model"]["B2"].value == 80
        wb.close()
        print("PASS: New profile runs a real Excel calculation and downloads the correct workbook", flush=True)

        # Returning users can select a saved profile and reselect their own workbook.
        data.pop("excel_path")
        data["command"] = "Tang lai vay len 9%"
        result = client.post("/api/scenario", data=data,
                             files={"file": (source.name, source.read_bytes())}).raise_for_status().json()
        assert result["ok"] and result["result"]["outputs"]["project_npv"]["value"] == 90
        assert hashlib.sha256(source.read_bytes()).hexdigest() == original
        print("PASS: Saved profile accepts a reselected workbook; source remains unchanged", flush=True)


if __name__ == "__main__":
    main()
