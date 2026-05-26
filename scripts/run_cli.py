from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

# Allow running from project root without installing as package.
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
from app.ai_parser import parse_user_command
from app.config_loader import load_model_map
from app.excel_controller import ExcelController
from app.utils import format_value
from app.validator import validate_action_plan


def main() -> None:
    load_dotenv(ROOT / ".env")
    parser = argparse.ArgumentParser(description="Run one AI Excel scenario from CLI.")
    parser.add_argument("--excel", required=True, help="Path to source Excel model")
    parser.add_argument("--command", required=True, help="Natural language command")
    parser.add_argument("--map", default=os.getenv("MODEL_MAP_PATH", "config/model_map.yaml"))
    parser.add_argument("--engine", default=os.getenv("EXCEL_ENGINE", "auto"), choices=["auto", "xlwings", "openpyxl"])
    parser.add_argument("--output-dir", default=os.getenv("OUTPUT_DIR", "outputs"))
    args = parser.parse_args()

    model_map_path = Path(args.map)
    if not model_map_path.is_absolute():
        model_map_path = ROOT / model_map_path

    model_map = load_model_map(model_map_path)
    plan = parse_user_command(args.command, model_map)
    plan = validate_action_plan(plan, model_map)

    controller = ExcelController(model_map, output_dir=ROOT / args.output_dir, engine=args.engine)
    result = controller.run_scenario(Path(args.excel), plan)

    print("\nACTION PLAN")
    print(json.dumps(plan.model_dump(), ensure_ascii=False, indent=2))
    print("\nAPPLIED CHANGES")
    for c in result.applied_changes:
        print(f"- {c.parameter}: {c.old_value} -> {c.new_value} ({c.sheet}!{c.cell})")
    print("\nOUTPUTS")
    for key, value in result.outputs.items():
        print(f"- {key}: {format_value(value.value, value.type, value.unit)}")
    print(f"\nScenario file: {result.scenario_file}")
    if result.warnings:
        print("\nWARNINGS")
        for w in result.warnings:
            print(f"- {w}")


if __name__ == "__main__":
    main()
