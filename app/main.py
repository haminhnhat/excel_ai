from __future__ import annotations

import os
import shutil
import time
import traceback
from pathlib import Path
from typing import Any, Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from .ai_parser import parse_user_command
from .accuracy import write_accuracy_event
from .config_loader import list_profiles, load_model_map_for_profile, resolve_model_map_path
from .excel_controller import ExcelController
from .utils import format_value, safe_filename
from .validator import ValidationError, validate_action_plan
from .onboarding import (
    build_model_map_from_candidates,
    save_profile_yaml,
    scan_workbook_for_mapping,
    validate_model_map_against_workbook,
)

load_dotenv()

BASE_DIR = Path(__file__).resolve().parents[1]

OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", BASE_DIR / "outputs"))
if not OUTPUT_DIR.is_absolute():
    OUTPUT_DIR = BASE_DIR / OUTPUT_DIR

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

AUDIT_LOG_PATH = Path(os.getenv("AUDIT_LOG_PATH", BASE_DIR / "logs" / "audit_log.csv"))
if not AUDIT_LOG_PATH.is_absolute():
    AUDIT_LOG_PATH = BASE_DIR / AUDIT_LOG_PATH

ACCURACY_LOG_PATH = Path(os.getenv("ACCURACY_LOG_PATH", BASE_DIR / "logs" / "accuracy_events.jsonl"))
if not ACCURACY_LOG_PATH.is_absolute():
    ACCURACY_LOG_PATH = BASE_DIR / ACCURACY_LOG_PATH

ERROR_LOG_PATH = Path(os.getenv("ERROR_LOG_PATH", BASE_DIR / "logs" / "app_errors.log"))
if not ERROR_LOG_PATH.is_absolute():
    ERROR_LOG_PATH = BASE_DIR / ERROR_LOG_PATH

SCENARIO_RUNTIME_LOG_PATH = Path(os.getenv("SCENARIO_RUNTIME_LOG_PATH", BASE_DIR / "logs" / "scenario_runtime.log"))
if not SCENARIO_RUNTIME_LOG_PATH.is_absolute():
    SCENARIO_RUNTIME_LOG_PATH = BASE_DIR / SCENARIO_RUNTIME_LOG_PATH

app = FastAPI(title="Excel AI Controller", version="0.1.0")

WEB_DIR = BASE_DIR / "web"
if WEB_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")


def _resolve_model_path() -> Path:
    model_path = Path(os.getenv("MODEL_PATH", BASE_DIR / "models" / "base_model.xlsx"))
    if not model_path.is_absolute():
        model_path = BASE_DIR / model_path
    return model_path


class OnboardingCandidate(BaseModel):
    decision: str = "Review"
    role: str
    parameter_key: str
    friendly_name: str | None = None
    sheet: str
    cell: str
    current_value: str | None = None
    formula: bool | None = None
    type: str = "number"
    unit: str | None = None
    min: float | None = None
    max: float | None = None
    editable: bool | None = True
    confidence: str | None = None
    nearby_label: str | None = None
    description: str | None = None
    aliases: list[str] | str | None = Field(default_factory=list)
    correct_sheet: str | None = None
    correct_cell: str | None = None
    target_mode: str | None = "cell"
    notes: str | None = None


class CreateProfileRequest(BaseModel):
    profile_name: str
    excel_path: str
    candidates: list[OnboardingCandidate]
    currency: str = "VND"
    overwrite: bool = True


def _resolve_any_path(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def _rel(path: Path) -> str:
    try:
        return str(path.relative_to(BASE_DIR)).replace("\\", "/")
    except ValueError:
        return str(path)


def _log_error(context: str, exc: Exception) -> None:
    try:
        ERROR_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with ERROR_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"\n[{context}]\n")
            f.write(traceback.format_exc())
            f.write("\n")
    except Exception:
        pass


def _append_runtime_log(message: str) -> None:
    try:
        SCENARIO_RUNTIME_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with SCENARIO_RUNTIME_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception:
        pass


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    index_path = WEB_DIR / "index.html"
    if index_path.exists():
        return index_path.read_text(encoding="utf-8")
    return "<h1>Excel AI Controller</h1>"


@app.get("/favicon.ico")
def favicon() -> FileResponse:
    icon_path = WEB_DIR / "favicon.svg"
    if icon_path.exists():
        return FileResponse(icon_path, media_type="image/svg+xml")
    raise HTTPException(status_code=404, detail="Favicon not found")


@app.get("/health")
def health() -> dict:
    selected_profile = os.getenv("MODEL_PROFILE")
    map_path = resolve_model_map_path(profile=selected_profile)
    return {
        "status": "ok",
        "model_profile": selected_profile,
        "available_profiles": list_profiles(BASE_DIR),
        "model_map_path": str(map_path),
        "default_model_path": str(_resolve_model_path()),
        "engine": os.getenv("EXCEL_ENGINE", "auto"),
        "ai_provider": os.getenv("AI_PROVIDER", "mock"),
    }




@app.get("/api/profiles")
def api_profiles() -> dict:
    profiles = list_profiles(BASE_DIR)
    return {"ok": True, "profiles": profiles, "default_profile": os.getenv("MODEL_PROFILE")}


@app.post("/api/onboarding/analyze")
async def onboarding_analyze(
    file: UploadFile = File(...),
    suggested_profile: Optional[str] = Form(default=None),
) -> dict:
    if not file.filename:
        raise HTTPException(status_code=400, detail="No Excel file uploaded.")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise HTTPException(status_code=400, detail="Please upload an .xlsx or .xlsm file.")

    stem = safe_filename(suggested_profile or Path(file.filename).stem)
    target_dir = UPLOAD_DIR / "onboarding" / stem
    target_dir.mkdir(parents=True, exist_ok=True)
    excel_path = target_dir / f"{stem}{suffix}"
    with excel_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        analysis = scan_workbook_for_mapping(excel_path)
    except Exception as e:
        _log_error(f"onboarding_analyze file={excel_path}", e)
        raise HTTPException(status_code=500, detail=f"Could not analyze workbook: {e}")

    analysis["excel_path"] = _rel(excel_path)
    analysis["suggested_profile"] = safe_filename(suggested_profile or analysis.get("suggested_profile") or stem)
    analysis["available_profiles"] = list_profiles(BASE_DIR)
    return {"ok": True, **analysis}


@app.post("/api/onboarding/validate-selection")
def onboarding_validate_selection(payload: CreateProfileRequest) -> dict:
    excel_path = _resolve_any_path(payload.excel_path)
    if not excel_path.exists():
        raise HTTPException(status_code=400, detail=f"Excel file not found: {excel_path}")
    candidates = [item.model_dump() for item in payload.candidates]
    try:
        model_map = build_model_map_from_candidates(payload.profile_name, candidates, currency=payload.currency)
        validation = validate_model_map_against_workbook(model_map, excel_path)
    except Exception as e:
        raise HTTPException(status_code=422, detail=str(e))
    return {"ok": validation["ok"], "validation": validation, "model_map_preview": model_map}


@app.post("/api/onboarding/create-profile")
def onboarding_create_profile(payload: CreateProfileRequest) -> dict:
    excel_path = _resolve_any_path(payload.excel_path)
    if not excel_path.exists():
        raise HTTPException(status_code=400, detail=f"Excel file not found: {excel_path}")
    candidates = [item.model_dump() for item in payload.candidates]
    try:
        profile = safe_filename(payload.profile_name)
        model_map = build_model_map_from_candidates(profile, candidates, currency=payload.currency)
        validation = validate_model_map_against_workbook(model_map, excel_path)
        if not validation["ok"]:
            return {
                "ok": False,
                "message": "Profile was not saved because validation failed.",
                "validation": validation,
                "model_map_preview": model_map,
            }
        out_path = save_profile_yaml(BASE_DIR, profile, model_map, overwrite=payload.overwrite)
        return {
            "ok": True,
            "profile": profile,
            "model_map_path": _rel(out_path),
            "excel_path": _rel(excel_path),
            "validation": validation,
            "available_profiles": list_profiles(BASE_DIR),
        }
    except Exception as e:
        raise HTTPException(status_code=422, detail=str(e))

@app.post("/api/scenario")
async def run_scenario(
    command: str = Form(...),
    file: Optional[UploadFile] = File(default=None),
    engine: Optional[str] = Form(default=None),
    model_profile: Optional[str] = Form(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    started_at = time.perf_counter()
    stage_timings: list[dict[str, Any]] = []

    def mark(stage: str) -> None:
        elapsed = round(time.perf_counter() - started_at, 3)
        item = {"elapsed_seconds": elapsed, "stage": stage}
        stage_timings.append(item)
        _append_runtime_log(f"{elapsed:>8.3f}s profile={model_profile or os.getenv('MODEL_PROFILE') or ''} stage={stage}")

    accuracy_event: dict[str, Any] = {
        "endpoint": "/api/scenario",
        "command": command,
        "requested_profile": model_profile,
        "engine": engine or os.getenv("EXCEL_ENGINE", "auto"),
        "status": "started",
    }
    try:
        mark("request:start")
        model_map, model_map_path = load_model_map_for_profile(profile=model_profile)
        mark("profile:loaded")
        accuracy_event["model_map_path"] = str(model_map_path)

        if file is not None and file.filename:
            safe_name = Path(file.filename).name
            source_path = UPLOAD_DIR / safe_name
            with source_path.open("wb") as f:
                shutil.copyfileobj(file.file, f)
            accuracy_event["excel_source"] = "uploaded_file"
            accuracy_event["excel_path"] = _rel(source_path)
        elif excel_path:
            source_path = _resolve_any_path(excel_path)
            accuracy_event["excel_source"] = "onboarded_path"
            accuracy_event["excel_path"] = _rel(source_path)
        else:
            source_path = _resolve_model_path()
            accuracy_event["excel_source"] = "default_model_path"
            accuracy_event["excel_path"] = _rel(source_path)

        if not source_path.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Excel model not found: {source_path}. Upload a file or set MODEL_PATH in .env.",
            )

        mark("parse_command:start")
        plan = parse_user_command(command, model_map)
        mark("parse_command:done")
        accuracy_event["parsed_action_plan"] = plan.model_dump()
        mark("validate_action_plan:start")
        plan = validate_action_plan(plan, model_map)
        mark("validate_action_plan:done")
        accuracy_event["validated_action_plan"] = plan.model_dump()

        controller = ExcelController(
            model_map=model_map,
            output_dir=OUTPUT_DIR,
            audit_log_path=AUDIT_LOG_PATH,
            engine=engine or os.getenv("EXCEL_ENGINE", "auto"),
            progress_callback=mark,
        )
        mark("excel_controller:start")
        result = controller.run_scenario(source_path, plan)
        mark("excel_controller:done")
        accuracy_event["status"] = "ok"
        accuracy_event["result_summary"] = {
            "scenario_file": result.scenario_file,
            "engine": result.engine,
            "recalculated": result.recalculated,
            "applied_change_count": len(result.applied_changes),
            "output_keys": list(result.outputs.keys()),
            "warnings": result.warnings,
            "duration_seconds": round(time.perf_counter() - started_at, 3),
        }
        scenario_file = Path(result.scenario_file)
        return {
            "ok": True,
            "model_map_path": str(model_map_path),
            "model_profile": model_profile or os.getenv("MODEL_PROFILE"),
            "action_plan": plan.model_dump(),
            "result": result.model_dump(),
            "formatted_outputs": {
                key: format_value(v.value, v.type, v.unit) for key, v in result.outputs.items()
            },
            "duration_seconds": round(time.perf_counter() - started_at, 3),
            "stage_timings": stage_timings,
            "download_url": f"/download/{scenario_file.name}",
        }
    except ValidationError as e:
        accuracy_event["status"] = "validation_error"
        accuracy_event["error"] = str(e)
        raise HTTPException(status_code=422, detail=str(e))
    except HTTPException:
        accuracy_event["status"] = "http_error"
        raise
    except Exception as e:
        accuracy_event["status"] = "error"
        accuracy_event["error"] = str(e)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        mark(f"request:end:{accuracy_event.get('status')}")
        accuracy_event["duration_seconds"] = round(time.perf_counter() - started_at, 3)
        accuracy_event["stage_timings"] = stage_timings
        try:
            write_accuracy_event(ACCURACY_LOG_PATH, accuracy_event)
        except Exception:
            pass


@app.get("/download/{filename}")
def download(filename: str) -> FileResponse:
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, filename=path.name)
