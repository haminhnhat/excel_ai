from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict
import yaml


BASE_DIR = Path(__file__).resolve().parents[1]


def resolve_model_map_path(
    explicit_path: str | Path | None = None,
    profile: str | None = None,
    base_dir: str | Path | None = None,
) -> Path:
    """Resolve the model_map.yaml path.

    Resolution order:
    1. explicit_path argument
    2. MODEL_MAP_PATH env var
    3. config/profiles/<profile>/model_map.yaml, where profile comes from argument or MODEL_PROFILE
    4. config/model_map.yaml legacy fallback
    """
    root = Path(base_dir) if base_dir is not None else BASE_DIR

    candidate: str | Path | None = explicit_path or os.getenv("MODEL_MAP_PATH")
    if candidate:
        path = Path(candidate)
        return path if path.is_absolute() else root / path

    selected_profile = profile or os.getenv("MODEL_PROFILE")
    if selected_profile:
        path = root / "config" / "profiles" / selected_profile / "model_map.yaml"
        return path

    legacy = root / "config" / "model_map.yaml"
    return legacy


def load_model_map(path: str | Path) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Model map not found: {p}")
    with p.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    if "inputs" not in data or "outputs" not in data:
        raise ValueError("model_map.yaml must contain 'inputs' and 'outputs'.")
    return data


def load_model_map_for_profile(profile: str | None = None, explicit_path: str | Path | None = None) -> tuple[Dict[str, Any], Path]:
    path = resolve_model_map_path(explicit_path=explicit_path, profile=profile)
    return load_model_map(path), path


def list_profiles(base_dir: str | Path | None = None) -> list[str]:
    root = Path(base_dir) if base_dir is not None else BASE_DIR
    profiles_dir = root / "config" / "profiles"
    if not profiles_dir.exists():
        return []
    return sorted(
        p.name for p in profiles_dir.iterdir()
        if p.is_dir() and (p / "model_map.yaml").exists()
    )
