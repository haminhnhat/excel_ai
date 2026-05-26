from __future__ import annotations

from typing import Any, Dict, List

from .schemas import ActionPlan


class ValidationError(Exception):
    pass


def validate_action_plan(plan: ActionPlan, model_map: Dict[str, Any]) -> ActionPlan:
    inputs = model_map.get("inputs", {})
    outputs = model_map.get("outputs", {})
    errors: list[str] = []

    if not plan.changes:
        errors.append("Action plan contains no changes.")

    for change in plan.changes:
        meta = inputs.get(change.parameter)
        if not meta:
            errors.append(f"Parameter not allowed: {change.parameter}")
            continue
        if not meta.get("editable", True):
            errors.append(f"Parameter is not editable: {change.parameter}")
            continue

        if isinstance(change.value, (int, float)):
            min_v = meta.get("min")
            max_v = meta.get("max")
            if min_v is not None and change.value < min_v:
                errors.append(f"{change.parameter}={change.value} is below min {min_v}.")
            if max_v is not None and change.value > max_v:
                errors.append(f"{change.parameter}={change.value} is above max {max_v}.")

    if not plan.requested_outputs:
        plan.requested_outputs = model_map.get("settings", {}).get("default_outputs", list(outputs.keys()))

    for output_key in plan.requested_outputs:
        if output_key not in outputs:
            errors.append(f"Output not allowed: {output_key}")

    if errors:
        raise ValidationError("\n".join(errors))

    return plan
