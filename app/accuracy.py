from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any


def write_accuracy_event(path: str | Path, event: dict[str, Any]) -> None:
    """Append one scenario parsing/execution event for later accuracy review."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "time": datetime.now().isoformat(timespec="seconds"),
        **event,
    }
    with target.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
