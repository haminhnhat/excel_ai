from __future__ import annotations

from pathlib import Path
from typing import Any, Dict
import yaml


def load_model_map(path: str | Path) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Model map not found: {p}")
    with p.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if "inputs" not in data or "outputs" not in data:
        raise ValueError("model_map.yaml must contain 'inputs' and 'outputs'.")
    return data
