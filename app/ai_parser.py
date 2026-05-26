from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List

from .schemas import ActionPlan, Change
from .utils import normalize_text, parse_number_token


def _default_outputs(model_map: Dict[str, Any]) -> List[str]:
    configured = model_map.get("settings", {}).get("default_outputs") or []
    if configured:
        return configured
    return list(model_map.get("outputs", {}).keys())


def _extract_percent_after_keyword(text_norm: str, keywords: list[str]) -> float | None:
    # Examples: "tang lai vay len 8%", "giam gia ban 5%"
    if not any(k in text_norm for k in keywords):
        return None
    m = re.search(r"([0-9]+(?:[\.,][0-9]+)?)\s*%", text_norm)
    if not m:
        return None
    value = parse_number_token(m.group(1)) / 100.0
    if "giam" in text_norm or "giảm" in text_norm:
        value = -abs(value)
    return value


def fallback_parse(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    """Rule-based Vietnamese parser for a small MVP.

    This keeps the demo usable even without an AI API key. It intentionally only
    handles common financial-model scenario commands.
    """
    text_norm = normalize_text(command)
    changes: list[Change] = []

    # Tổng mức đầu tư from old value to new value.
    # Example: "tang tong muc dau tu tu 1000 ty len 1200 ty"
    has_tmdt = any(k in text_norm for k in ["tong muc dau tu", "tmdt", "chi phi dau tu"])
    if has_tmdt:
        m = re.search(
            r"tu\s*([0-9]+(?:[\.,][0-9]+)?)\s*(?:ty|ti|ty dong|ti dong)?\s*(?:len|thanh|den)\s*([0-9]+(?:[\.,][0-9]+)?)\s*(?:ty|ti|ty dong|ti dong)?",
            text_norm,
        )
        if m:
            old_v = parse_number_token(m.group(1))
            new_v = parse_number_token(m.group(2))
            if old_v == 0:
                raise ValueError("Cannot calculate investment_cost_change because old TMĐT is zero.")
            pct_change = new_v / old_v - 1
            changes.append(
                Change(
                    parameter="investment_cost_change",
                    value=pct_change,
                    operation="set",
                    reason=f"TMĐT changed from {old_v:g} tỷ to {new_v:g} tỷ => {pct_change:.4%}",
                )
            )
        else:
            # Example: "tang tong muc dau tu 20%"
            pct = _extract_percent_after_keyword(text_norm, ["tong muc dau tu", "tmdt", "chi phi dau tu"])
            if pct is not None:
                changes.append(
                    Change(
                        parameter="investment_cost_change",
                        value=pct,
                        operation="set",
                        reason="User requested a percentage change to TMĐT.",
                    )
                )

    pct = _extract_percent_after_keyword(text_norm, ["lai vay", "lai suat"])
    if pct is not None:
        # For interest rate, "tăng lãi vay lên 8%" usually means set to 8%, not +8%.
        value = abs(pct)
        changes.append(
            Change(parameter="loan_interest_rate", value=value, operation="set", reason="User requested loan interest rate.")
        )

    pct = _extract_percent_after_keyword(text_norm, ["gia ban", "don gia"])
    if pct is not None:
        changes.append(
            Change(parameter="selling_price_change", value=pct, operation="set", reason="User requested selling price change.")
        )

    pct = _extract_percent_after_keyword(text_norm, ["vat", "thue vat"])
    if pct is not None:
        changes.append(Change(parameter="vat_rate", value=abs(pct), operation="set", reason="User requested VAT rate."))

    pct = _extract_percent_after_keyword(text_norm, ["tndn", "cit", "thue thu nhap doanh nghiep"])
    if pct is not None:
        changes.append(Change(parameter="cit_rate", value=abs(pct), operation="set", reason="User requested CIT/TNDN rate."))

    if not changes:
        raise ValueError(
            "Could not parse command with fallback parser. Use AI_PROVIDER=openai or add a parsing rule."
        )

    return ActionPlan(
        scenario_name=command[:80],
        changes=changes,
        requested_outputs=_default_outputs(model_map),
        raw_command=command,
    )


def _openai_parse(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    from openai import OpenAI

    inputs = model_map.get("inputs", {})
    outputs = model_map.get("outputs", {})
    default_outputs = _default_outputs(model_map)

    system = """
You are a strict controller for an Excel financial model.
Convert the user's natural-language request into a JSON action plan.
Rules:
- Use only parameter keys listed in ALLOWED_INPUTS.
- Use only output keys listed in ALLOWED_OUTPUTS.
- Do not invent cell references.
- Do not modify formulas.
- Return JSON only.
- Percent values must be decimals: 8% => 0.08, +20% => 0.2, -5% => -0.05.
- If the user says total investment changes from X to Y, convert to percentage change: Y / X - 1, and use parameter investment_cost_change.
""".strip()

    user = {
        "user_command": command,
        "allowed_inputs": {
            k: {
                "description": v.get("description"),
                "type": v.get("type"),
                "min": v.get("min"),
                "max": v.get("max"),
                "aliases": v.get("aliases", []),
            }
            for k, v in inputs.items()
        },
        "allowed_outputs": list(outputs.keys()),
        "default_outputs": default_outputs,
        "required_json_shape": {
            "scenario_name": "string",
            "changes": [{"parameter": "string", "value": "number|string|boolean|null", "operation": "set", "reason": "string"}],
            "requested_outputs": ["string"],
        },
    }

    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
    model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )
    content = resp.choices[0].message.content or "{}"
    data = json.loads(content)
    data.setdefault("raw_command", command)
    if not data.get("requested_outputs"):
        data["requested_outputs"] = default_outputs
    return ActionPlan.model_validate(data)


def parse_user_command(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    provider = os.getenv("AI_PROVIDER", "mock").lower().strip()
    if provider == "openai":
        return _openai_parse(command, model_map)
    return fallback_parse(command, model_map)
