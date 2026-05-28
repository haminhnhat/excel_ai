from __future__ import annotations

import json
import os
import re
import urllib.request
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .schemas import ActionPlan, Change
from .utils import normalize_text, parse_number_token

# Matches both 8% and 8.5% after text normalization.
PERCENT_RE = r"([+-]?[0-9]+(?:[\.,][0-9]+)?)\s*(?:%|phan tram|percent|percentage)"
NUMBER_RE = r"([0-9]+(?:[\.,][0-9]+)?)"

# Parameters that represent an absolute rate. A command like "lãi vay tăng lên 8%"
# should set the rate to +8%, not +8 percentage points.
ABSOLUTE_RATE_PARAMS = {
    "loan_interest_rate",
    "vat_rate",
    "cit_rate",
    "selling_cost_rate",
    "marketing_rate",
    "admin_cost_rate",
}

# Parameters that represent relative scenario adjustments. A command like
# "giá bán giảm 5%" should write -0.05.
ADJUSTMENT_PARAMS = {
    "investment_cost_change",
    "selling_price_change",
    "premium_change",
}

NEGATIVE_WORDS = [
    "giam",
    "giam xuong",
    "ha",
    "ha xuong",
    "thap hon",
    "sut",
    "sut giam",
    "mat",
    "am",
    "worse",
    "decrease",
    "reduce",
    "reduction",
    "lower",
]

POSITIVE_WORDS = [
    "tang",
    "tang len",
    "nang",
    "nang len",
    "cao hon",
    "them",
    "upside",
    "tot",
    "increase",
    "raise",
    "higher",
    "add",
]

CLAUSE_SPLIT_RE = re.compile(r"[,;\n]+|\s+va\s+|\s+voi\s+|\s+kem\s+|\s+sau do\s+|\s+roi\s+")


def _clean_command(command: str) -> str:
    """Normalize Vietnamese text and standardize common percentage phrases."""
    text = normalize_text(command)
    text = text.replace("phan tram", "%")
    text = text.replace("percent", "%")
    text = text.replace("percentage", "%")
    return text


def _default_outputs(model_map: Dict[str, Any]) -> List[str]:
    configured = model_map.get("settings", {}).get("default_outputs") or []
    if configured:
        return configured
    return list(model_map.get("outputs", {}).keys())


def _split_clauses(text_norm: str) -> list[str]:
    parts = [p.strip(" .:-") for p in CLAUSE_SPLIT_RE.split(text_norm) if p.strip(" .:-")]
    return parts or [text_norm]


def _first_percent(text: str) -> float | None:
    m = re.search(PERCENT_RE, text)
    if not m:
        return None
    return parse_number_token(m.group(1)) / 100.0


def _first_number(text: str) -> float | None:
    m = re.search(NUMBER_RE, text)
    if not m:
        return None
    return parse_number_token(m.group(1))


def _target_number(text: str) -> float | None:
    m = re.search(
        rf"(?:la|=|bang|thanh|len|toi|den|sang)\s*\$?\s*{NUMBER_RE}",
        text,
    )
    if m:
        return parse_number_token(m.group(1))
    return _first_number(text)


def _signed_percent_for_change(text: str) -> float | None:
    value = _first_percent(text)
    if value is None:
        return None
    if any(w in text for w in NEGATIVE_WORDS):
        return -abs(value)
    if any(w in text for w in POSITIVE_WORDS):
        return abs(value)
    # For adjustment parameters, an unsigned percentage means +x% by default.
    return value


def _absolute_rate_percent(text: str) -> float | None:
    value = _first_percent(text)
    if value is not None:
        return abs(value)

    # Fallback for common rate commands where the user omits the % sign:
    # "lãi vay là 8" => 0.08. Limit to values in a human rate range.
    n = _first_number(text)
    if n is None:
        return None
    if 0 <= n <= 1:
        return n
    if 1 < n <= 100:
        return n / 100.0
    return None


def _extract_tmdt_from_to(text: str) -> float | None:
    """Return percentage change for 'from X to Y' total investment commands."""
    m = re.search(
        rf"tu\s*{NUMBER_RE}\s*(?:ty|ti|ty dong|ti dong|bn|billion)?\s*(?:len|thanh|den|toi|sang|xuong)\s*{NUMBER_RE}\s*(?:ty|ti|ty dong|ti dong|bn|billion)?",
        text,
    )
    if not m:
        return None
    old_v = parse_number_token(m.group(1))
    new_v = parse_number_token(m.group(2))
    if old_v == 0:
        raise ValueError("Cannot calculate investment_cost_change because old TMĐT is zero.")
    return new_v / old_v - 1


def _extract_rate_from_to(text: str) -> float | None:
    """Return new absolute rate for 'from X% to Y%' or 'from X to Y' rate commands."""
    m = re.search(rf"tu\s*{PERCENT_RE}\s*(?:len|thanh|den|toi|sang|xuong)\s*{PERCENT_RE}", text)
    if m:
        return abs(parse_number_token(m.group(2)) / 100.0)

    # If no percent signs are present, still accept "lãi vay từ 7 lên 8".
    m2 = re.search(rf"tu\s*{NUMBER_RE}\s*(?:len|thanh|den|toi|sang|xuong)\s*{NUMBER_RE}", text)
    if m2:
        new_v = parse_number_token(m2.group(2))
        return new_v if 0 <= new_v <= 1 else new_v / 100.0
    return None


def _add_or_replace(changes: list[Change], change: Change) -> None:
    """Keep one final change per parameter. Later clauses win."""
    for i, existing in enumerate(changes):
        if existing.parameter == change.parameter:
            changes[i] = change
            return
    changes.append(change)


def _normalized_aliases(model_map: Dict[str, Any]) -> dict[str, list[str]]:
    aliases: dict[str, list[str]] = {}
    for key, meta in model_map.get("inputs", {}).items():
        raw_aliases = list(meta.get("aliases", [])) + [key.replace("_", " ")]
        aliases[key] = sorted({_clean_command(a) for a in raw_aliases if a}, key=len, reverse=True)
    return aliases


def _find_exact_alias_matches(text: str, aliases: dict[str, list[str]]) -> list[tuple[int, int, str, str]]:
    """Return exact alias matches as (start, end, parameter, alias)."""
    matches: list[tuple[int, int, str, str]] = []
    for param, alias_list in aliases.items():
        for alias in alias_list:
            if not alias:
                continue
            pattern = r"(?<![a-z0-9])" + re.escape(alias) + r"(?![a-z0-9])"
            for m in re.finditer(pattern, text):
                matches.append((m.start(), m.end(), param, alias))
    # Remove overlapping shorter matches by preferring earlier and longer aliases.
    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    filtered: list[tuple[int, int, str, str]] = []
    occupied: list[range] = []
    for match in matches:
        span = range(match[0], match[1])
        if any(max(span.start, r.start) < min(span.stop, r.stop) for r in occupied):
            continue
        filtered.append(match)
        occupied.append(span)
    return sorted(filtered, key=lambda x: x[0])


def _best_fuzzy_param(text: str, aliases: dict[str, list[str]]) -> tuple[str | None, float, str | None]:
    """Return best fuzzy parameter match for a whole clause.

    This catches slightly different wording, but keeps a high threshold so the
    parser fails safely instead of changing the wrong Excel cell.
    """
    best_param: str | None = None
    best_alias: str | None = None
    best_score = 0.0

    for param, alias_list in aliases.items():
        for alias in alias_list:
            if not alias:
                continue
            if alias in text:
                score = 1.0
            else:
                # Compare alias against same-length windows and whole clause.
                windows = [text]
                words = text.split()
                alias_len = max(1, len(alias.split()))
                for i in range(0, max(1, len(words) - alias_len + 1)):
                    windows.append(" ".join(words[i : i + alias_len]))
                    if alias_len + 1 <= len(words):
                        windows.append(" ".join(words[i : i + alias_len + 1]))
                score = max(SequenceMatcher(None, alias, w).ratio() for w in windows)
            if score > best_score:
                best_param, best_alias, best_score = param, alias, score

    return best_param, best_score, best_alias


def _change_from_segment(param: str, segment: str, meta: Dict[str, Any] | None = None) -> Change | None:
    """Convert one parameter-specific clause segment into a Change."""
    meta = meta or {}
    if param == "investment_cost_change":
        pct = _extract_tmdt_from_to(segment)
        if pct is None:
            pct = _signed_percent_for_change(segment)
        if pct is None:
            return None
        return Change(parameter=param, value=pct, operation="set", reason="Parsed TMĐT / investment cost change.")

    if param == "selling_price_change":
        pct = _signed_percent_for_change(segment)
        if pct is None:
            return None
        return Change(parameter=param, value=pct, operation="set", reason="Parsed selling price / revenue change.")

    if param in ADJUSTMENT_PARAMS or param.endswith("_change"):
        pct = _signed_percent_for_change(segment)
        if pct is None and meta.get("base_value") not in (None, ""):
            target = _target_number(segment)
            base = float(meta["base_value"])
            if target is not None and base:
                pct = target / base - 1
        if pct is None:
            return None
        return Change(parameter=param, value=pct, operation="set", reason=f"Parsed {param} adjustment.")

    if param in ABSOLUTE_RATE_PARAMS:
        value = _extract_rate_from_to(segment)
        if value is None:
            value = _absolute_rate_percent(segment)
        if value is None:
            return None
        return Change(parameter=param, value=value, operation="set", reason=f"Parsed {param}.")

    # Generic fallback for future numeric inputs.
    pct = _first_percent(segment)
    if pct is not None:
        return Change(parameter=param, value=pct, operation="set", reason=f"Parsed {param}.")
    value = _target_number(segment)
    if value is not None:
        return Change(parameter=param, value=value, operation="set", reason=f"Parsed {param}.")
    return None


def _parse_with_aliases(command: str, model_map: Dict[str, Any]) -> list[Change]:
    text_norm = _clean_command(command)
    clauses = _split_clauses(text_norm)
    aliases = _normalized_aliases(model_map)
    changes: list[Change] = []

    for clause in clauses:
        exact_matches = _find_exact_alias_matches(clause, aliases)

        if exact_matches:
            # A clause may contain multiple commands, e.g. "VAT 10% TNDN 20%".
            for idx, (start, end, param, alias) in enumerate(exact_matches):
                next_start = exact_matches[idx + 1][0] if idx + 1 < len(exact_matches) else len(clause)
                # Include a little text before adjustment aliases to capture
                # "giảm giá bán 5%". For absolute rates, start at the alias so
                # "VAT 10% TNDN 20%" does not leak 10% into TNDN.
                segment_start = max(0, start - 18) if param in ADJUSTMENT_PARAMS or param.endswith("_change") else start
                segment = clause[segment_start:next_start]
                change = _change_from_segment(param, segment, model_map.get("inputs", {}).get(param, {}))
                if change is not None:
                    _add_or_replace(changes, change)
            continue

        # Fuzzy fallback when no exact alias matched.
        param, score, alias = _best_fuzzy_param(clause, aliases)
        if param and score >= 0.78:
            change = _change_from_segment(param, clause, model_map.get("inputs", {}).get(param, {}))
            if change is not None:
                change.reason = f"Parsed by fuzzy alias matching: '{alias}' score={score:.2f}."
                _add_or_replace(changes, change)

    return changes


def _output_aliases(model_map: Dict[str, Any]) -> dict[str, list[str]]:
    defaults: dict[str, list[str]] = {
        "profit_after_tax": ["loi nhuan sau thue", "profit after tax", "pat", "lnst"],
        "project_npv": ["npv du an", "npv project", "project npv", "npv"],
        "project_irr": ["irr du an", "irr project", "project irr", "irr"],
        "equity_npv": ["npv chu dau tu", "investor npv", "equity npv"],
        "equity_irr": ["irr chu dau tu", "investor irr", "equity irr"],
        "roi": ["roi", "return on investment"],
        "total_investment": ["tong muc dau tu", "tmdt", "tmđt", "total investment"],
        "bank_loan": ["vay ngan hang", "bank loan", "khoan vay"],
    }
    out: dict[str, list[str]] = {}
    for key, meta in model_map.get("outputs", {}).items():
        raw = list(meta.get("aliases", [])) + defaults.get(key, []) + [key.replace("_", " ")]
        out[key] = sorted({_clean_command(a) for a in raw if a}, key=len, reverse=True)
    return out


def _requested_outputs(command: str, model_map: Dict[str, Any], has_changes: bool) -> list[str]:
    text_norm = _clean_command(command)

    # For write commands, do not treat input mentions such as "TMĐT tăng 15%"
    # as output requests. Return default scenario outputs unless the command
    # explicitly asks for results or names clear output metrics such as NPV/IRR.
    explicit_output_signal = any(
        phrase in text_norm
        for phrase in [
            "cho toi", "xem", "tra ra", "hien thi", "ket qua", "bao nhieu",
            "output", "show", "return", "read", "current", "hien tai",
            "npv", "irr", "roi", "loi nhuan", "lnst", "profit",
        ]
    )
    if has_changes and not explicit_output_signal:
        return _default_outputs(model_map)

    outputs = []
    for key, aliases in _output_aliases(model_map).items():
        if any(re.search(r"(?<![a-z0-9])" + re.escape(alias) + r"(?![a-z0-9])", text_norm) for alias in aliases):
            outputs.append(key)

    # If user says only "NPV" or "IRR", include both project and equity variants.
    if "npv" in text_norm:
        for k in ["project_npv", "equity_npv"]:
            if k in model_map.get("outputs", {}) and k not in outputs:
                outputs.append(k)
    if "irr" in text_norm:
        for k in ["project_irr", "equity_irr"]:
            if k in model_map.get("outputs", {}) and k not in outputs:
                outputs.append(k)

    if outputs:
        return outputs
    return _default_outputs(model_map)


def fallback_parse(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    """Hybrid no-API parser: exact rules + alias dictionary + fuzzy matching.

    This mode needs no API key. It covers common Vietnamese/English commands
    without trying to give the model arbitrary freedom over the workbook.
    """
    changes = _parse_with_aliases(command, model_map)
    requested_outputs = _requested_outputs(command, model_map, has_changes=bool(changes))

    if not changes:
        # Allow read-only questions such as "Cho tôi NPV và IRR hiện tại".
        text_norm = _clean_command(command)
        read_words = ["cho toi", "xem", "hien tai", "bao nhieu", "ket qua", "show", "read", "current"]
        has_output_reference = requested_outputs != _default_outputs(model_map) or any(k in text_norm for k in ["npv", "irr", "roi", "loi nhuan"])
        if not (has_output_reference and any(w in text_norm for w in read_words)):
            raise ValueError(
                "Could not parse command safely. Try specifying one approved input such as: "
                "'Tăng lãi vay lên 8%', 'Giảm giá bán 5%', "
                "'Tăng TMĐT từ 1000 tỷ lên 1200 tỷ', or "
                "'TMĐT tăng 15%, giá bán giảm 5%, lãi vay 8%'."
            )

    return ActionPlan(
        scenario_name=command[:80] or "Read current outputs",
        changes=changes,
        requested_outputs=requested_outputs,
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
- For rate parameters such as loan_interest_rate, VAT, and CIT, "increase to 8%" means set value = 0.08.
- For adjustment parameters such as selling_price_change, "decrease 5%" means value = -0.05.
- If the user says total investment changes from X to Y, convert to percentage change: Y / X - 1, and use parameter investment_cost_change.
- If the user only asks to read outputs, return changes: [].
""".strip()

    user = {
        "user_command": command,
        "ALLOWED_INPUTS": {
            k: {
                "description": v.get("description"),
                "type": v.get("type"),
                "min": v.get("min"),
                "max": v.get("max"),
                "aliases": v.get("aliases", []),
            }
            for k, v in inputs.items()
        },
        "ALLOWED_OUTPUTS": list(outputs.keys()),
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


def _ollama_parse(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    """Optional local LLM parser. Requires Ollama running locally.

    Set AI_PROVIDER=ollama and optionally OLLAMA_MODEL=qwen2.5:7b-instruct.
    This is not required for MVP; it is a privacy-preserving fallback when you
    do not want an external API key.
    """
    url = os.getenv("OLLAMA_URL", "http://localhost:11434/api/chat")
    model = os.getenv("OLLAMA_MODEL", "qwen2.5:7b-instruct")
    prompt = {
        "task": "Convert user command into strict JSON for an Excel financial model.",
        "rules": [
            "Use only allowed input keys.",
            "Use only allowed output keys.",
            "Return JSON only. No markdown.",
            "Percent values are decimals: 8% => 0.08, -5% => -0.05.",
            "If command only asks outputs, return changes as an empty list.",
        ],
        "user_command": command,
        "allowed_inputs": {
            k: {"description": v.get("description"), "aliases": v.get("aliases", []), "min": v.get("min"), "max": v.get("max")}
            for k, v in model_map.get("inputs", {}).items()
        },
        "allowed_outputs": list(model_map.get("outputs", {}).keys()),
        "default_outputs": _default_outputs(model_map),
        "json_shape": {
            "scenario_name": "string",
            "changes": [{"parameter": "string", "value": "number", "operation": "set", "reason": "string"}],
            "requested_outputs": ["string"],
        },
    }
    body = json.dumps(
        {
            "model": model,
            "messages": [
                {"role": "system", "content": "You return valid JSON only."},
                {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
            ],
            "stream": False,
            "format": "json",
        },
        ensure_ascii=False,
    ).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    content = payload.get("message", {}).get("content", "{}")
    data = json.loads(content)
    data.setdefault("raw_command", command)
    if not data.get("requested_outputs"):
        data["requested_outputs"] = _default_outputs(model_map)
    return ActionPlan.model_validate(data)


def parse_user_command(command: str, model_map: Dict[str, Any]) -> ActionPlan:
    provider = os.getenv("AI_PROVIDER", "mock").lower().strip()

    if provider == "openai":
        try:
            return _openai_parse(command, model_map)
        except Exception:
            if os.getenv("AI_FALLBACK_TO_RULES", "true").lower() == "true":
                return fallback_parse(command, model_map)
            raise

    if provider in {"ollama", "local", "local_llm"}:
        try:
            return _ollama_parse(command, model_map)
        except Exception:
            if os.getenv("AI_FALLBACK_TO_RULES", "true").lower() == "true":
                return fallback_parse(command, model_map)
            raise

    return fallback_parse(command, model_map)
