from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.ai_parser import parse_user_command
from app.config_loader import load_model_map_for_profile
from app.validator import ValidationError, validate_action_plan


@dataclass
class CaseResult:
    name: str
    ok: bool
    error: str | None = None


def _load_cases(path: Path) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line_no, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                cases.append(json.loads(line))
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSONL at {path}:{line_no}: {e}") from e
    return cases


def _changes_dict(plan: Any) -> dict[str, Any]:
    return {change.parameter: change.value for change in plan.changes}


def _close_enough(actual: Any, expected: Any, tolerance: float) -> bool:
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        return abs(float(actual) - float(expected)) <= tolerance
    return actual == expected


def _compare_changes(actual: dict[str, Any], expected: dict[str, Any], tolerance: float) -> list[str]:
    errors: list[str] = []
    if set(actual) != set(expected):
        errors.append(f"change keys expected={sorted(expected)} actual={sorted(actual)}")
        return errors
    for key, expected_value in expected.items():
        if not _close_enough(actual[key], expected_value, tolerance):
            errors.append(f"{key} expected={expected_value!r} actual={actual[key]!r}")
    return errors


def evaluate_case(case: dict[str, Any], default_profile: str | None, tolerance: float) -> CaseResult:
    name = str(case.get("name") or case.get("command") or "unnamed")
    profile = case.get("profile") or default_profile
    try:
        model_map, _ = load_model_map_for_profile(profile=profile)
        plan = parse_user_command(str(case["command"]), model_map)
        plan = validate_action_plan(plan, model_map)
    except (ValidationError, Exception) as e:
        return CaseResult(name=name, ok=False, error=f"parse/validation failed: {e}")

    errors: list[str] = []
    expected_changes = case.get("expected_changes", {})
    errors.extend(_compare_changes(_changes_dict(plan), expected_changes, tolerance))

    expected_outputs = case.get("expected_outputs")
    if expected_outputs is not None and set(plan.requested_outputs) != set(expected_outputs):
        errors.append(
            f"outputs expected={sorted(expected_outputs)} actual={sorted(plan.requested_outputs)}"
        )

    if errors:
        return CaseResult(name=name, ok=False, error="; ".join(errors))
    return CaseResult(name=name, ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Evaluate parser accuracy against JSONL test cases.")
    parser.add_argument("--cases", default=str(ROOT / "tests" / "parser_eval_cases.jsonl"))
    parser.add_argument("--profile", default=None, help="Default profile if a case does not specify one.")
    parser.add_argument("--tolerance", type=float, default=1e-9)
    args = parser.parse_args()

    cases = _load_cases(Path(args.cases))
    results = [evaluate_case(case, args.profile, args.tolerance) for case in cases]
    passed = sum(1 for result in results if result.ok)
    total = len(results)

    print(f"Parser accuracy: {passed}/{total} ({(passed / total * 100 if total else 0):.1f}%)")
    for result in results:
        status = "PASS" if result.ok else "FAIL"
        print(f"{status} {result.name}")
        if result.error:
            print(f"  {result.error}")

    return 0 if passed == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
