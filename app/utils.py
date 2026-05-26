from __future__ import annotations

import csv
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


def normalize_text(text: str) -> str:
    text = text.lower().strip()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d")
    return re.sub(r"\s+", " ", text)


def parse_number_token(token: str) -> float:
    token = token.strip().replace(" ", "")
    # Vietnamese often uses comma as decimal separator.
    if "," in token and "." not in token:
        token = token.replace(",", ".")
    # If both exist, assume comma is thousand separator.
    if "," in token and "." in token:
        token = token.replace(",", "")
    return float(token)


def safe_filename(name: str, max_len: int = 80) -> str:
    name = normalize_text(name)
    name = re.sub(r"[^a-z0-9._-]+", "_", name).strip("_")
    return name[:max_len] or "scenario"


def format_value(value: Any, value_type: str | None = None, unit: str | None = None) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        if value_type == "percent":
            return f"{value * 100:,.2f}%"
        if unit == "VND" or value_type == "currency":
            return f"{value:,.0f} VND"
        return f"{value:,.4f}"
    return str(value)


def append_audit_log(path: str | Path, rows: Iterable[dict[str, Any]]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "time",
        "scenario_name",
        "parameter",
        "sheet",
        "cell",
        "old_value",
        "new_value",
        "reason",
    ]
    exists = p.exists()
    with p.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not exists:
            writer.writeheader()
        now = datetime.now().isoformat(timespec="seconds")
        for row in rows:
            out = {k: row.get(k) for k in fieldnames}
            out["time"] = now
            writer.writerow(out)
