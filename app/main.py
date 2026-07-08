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
from .sensitivity import (
    export_sensitivity_by_prompt,
    export_sensitivity_range,
    preview_sensitivity_by_prompt,
    preview_sensitivity_range,
)
from .workbook_transfer import (
    export_workbook_value_transfer,
    plan_workbook_value_transfer,
)
from .format_normalizer import (
    export_format_normalization,
    preview_format_normalization,
)
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
    base_value: float | None = None
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
    overwrite: bool = False


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


def _resolve_excel_source(file: Optional[UploadFile], excel_path: Optional[str]) -> Path:
    if file is not None and file.filename:
        safe_name = Path(file.filename).name
        source_path = UPLOAD_DIR / safe_name
        with source_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        return source_path
    if excel_path:
        return _resolve_any_path(excel_path)
    return _resolve_model_path()


def _resolve_uploaded_or_path(file: Optional[UploadFile], excel_path: Optional[str], subdir: str) -> Path:
    if file is not None and file.filename:
        safe_name = Path(file.filename).name
        target_dir = UPLOAD_DIR / subdir
        target_dir.mkdir(parents=True, exist_ok=True)
        source_path = target_dir / f"{int(time.time() * 1000)}_{safe_name}"
        with source_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        return source_path
    if excel_path:
        return _resolve_any_path(excel_path)
    raise HTTPException(status_code=400, detail="Please upload a workbook file.")


def _ensure_request_allowed() -> None:
    return None


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
    _ensure_request_allowed()
    profiles = list_profiles(BASE_DIR)
    return {"ok": True, "profiles": profiles, "default_profile": os.getenv("MODEL_PROFILE")}


@app.delete("/api/profiles/{profile_name}")
def delete_profile(profile_name: str) -> dict:
    _ensure_request_allowed()
    profile = safe_filename(profile_name)
    if not profile:
        raise HTTPException(status_code=400, detail="Profile name is required.")
    profiles_dir = BASE_DIR / "config" / "profiles"
    profile_dir = profiles_dir / profile
    if not profile_dir.exists() or not (profile_dir / "model_map.yaml").exists():
        raise HTTPException(status_code=404, detail=f"Profile not found: {profile}")
    try:
        resolved = profile_dir.resolve()
        root = profiles_dir.resolve()
        if root not in resolved.parents:
            raise HTTPException(status_code=400, detail="Invalid profile path.")
        shutil.rmtree(profile_dir)
        remaining = list_profiles(BASE_DIR)
        return {
            "ok": True,
            "deleted_profile": profile,
            "available_profiles": remaining,
            "deleted_default_profile": profile == os.getenv("MODEL_PROFILE"),
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error(f"delete_profile profile={profile}", e)
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/onboarding/analyze")
async def onboarding_analyze(
    file: UploadFile = File(...),
    suggested_profile: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
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
    _ensure_request_allowed()
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
    _ensure_request_allowed()
    excel_path = _resolve_any_path(payload.excel_path)
    if not excel_path.exists():
        raise HTTPException(status_code=400, detail=f"Excel file not found: {excel_path}")
    candidates = [item.model_dump() for item in payload.candidates]
    try:
        profile = safe_filename(payload.profile_name)
        profile_path = BASE_DIR / "config" / "profiles" / profile / "model_map.yaml"
        if profile_path.exists() and not payload.overwrite:
            return {
                "ok": False,
                "conflict": True,
                "profile": profile,
                "model_map_path": _rel(profile_path),
                "message": f"Profile already exists: {profile}. Confirm overwrite to replace it.",
                "available_profiles": list_profiles(BASE_DIR),
            }
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
    _ensure_request_allowed()
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


@app.post("/api/command/preview")
async def preview_command(
    command: str = Form(...),
    model_profile: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        model_map, model_map_path = load_model_map_for_profile(profile=model_profile)
        plan = parse_user_command(command, model_map)
        plan = validate_action_plan(plan, model_map)
        input_map = model_map.get("inputs", {})
        output_map = model_map.get("outputs", {})
        changes = []
        for change in plan.changes:
            meta = input_map.get(change.parameter, {})
            changes.append({
                "parameter": change.parameter,
                "value": change.value,
                "operation": change.operation,
                "reason": change.reason,
                "sheet": meta.get("sheet"),
                "cell": meta.get("cell") or meta.get("range") or meta.get("cells"),
                "description": meta.get("description"),
            })
        outputs = [
            {
                "parameter": key,
                "sheet": output_map.get(key, {}).get("sheet"),
                "cell": output_map.get(key, {}).get("cell"),
                "description": output_map.get(key, {}).get("description"),
            }
            for key in plan.requested_outputs
        ]
        return {
            "ok": True,
            "model_profile": model_profile or os.getenv("MODEL_PROFILE"),
            "model_map_path": str(model_map_path),
            "action_plan": plan.model_dump(),
            "preview": {"changes": changes, "outputs": outputs},
        }
    except ValidationError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/scenario/preview-results")
async def preview_scenario_results(
    command: str = Form(...),
    file: Optional[UploadFile] = File(default=None),
    engine: Optional[str] = Form(default=None),
    model_profile: Optional[str] = Form(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    started_at = time.perf_counter()
    stage_timings: list[dict[str, Any]] = []
    scenario_path: Path | None = None

    def mark(stage: str) -> None:
        elapsed = round(time.perf_counter() - started_at, 3)
        item = {"elapsed_seconds": elapsed, "stage": stage}
        stage_timings.append(item)
        _append_runtime_log(
            f"{elapsed:>8.3f}s preview=true profile={model_profile or os.getenv('MODEL_PROFILE') or ''} stage={stage}"
        )

    try:
        mark("request:start")
        model_map, model_map_path = load_model_map_for_profile(profile=model_profile)
        mark("profile:loaded")
        source_path = _resolve_excel_source(file, excel_path)
        if not source_path.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Excel model not found: {source_path}. Upload a file or set MODEL_PATH in .env.",
            )

        mark("parse_command:start")
        plan = parse_user_command(command, model_map)
        mark("parse_command:done")
        mark("validate_action_plan:start")
        plan = validate_action_plan(plan, model_map)
        mark("validate_action_plan:done")

        controller = ExcelController(
            model_map=model_map,
            output_dir=OUTPUT_DIR / "preview",
            audit_log_path=AUDIT_LOG_PATH,
            engine=engine or os.getenv("EXCEL_ENGINE", "auto"),
            progress_callback=mark,
            save_workbook=False,
            write_audit_log=False,
        )
        mark("excel_controller:start")
        result = controller.run_scenario(source_path, plan)
        mark("excel_controller:done")
        scenario_path = Path(result.scenario_file)
        duration = round(time.perf_counter() - started_at, 3)
        return {
            "ok": True,
            "preview_only": True,
            "model_map_path": str(model_map_path),
            "model_profile": model_profile or os.getenv("MODEL_PROFILE"),
            "action_plan": plan.model_dump(),
            "result": result.model_dump(),
            "formatted_outputs": {
                key: format_value(v.value, v.type, v.unit) for key, v in result.outputs.items()
            },
            "duration_seconds": duration,
            "stage_timings": stage_timings,
        }
    except ValidationError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:
        _log_error("scenario_preview_results", e)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if scenario_path and scenario_path.exists():
            try:
                scenario_path.unlink()
                mark("preview_file:deleted")
            except Exception as e:
                _log_error(f"scenario_preview_cleanup file={scenario_path}", e)
                mark("preview_file:delete_failed")
        mark("request:end")


@app.post("/api/sensitivity/export")
async def export_sensitivity(
    sheet: str = Form(...),
    range_address: str = Form(...),
    output_name: Optional[str] = Form(default=None),
    file: Optional[UploadFile] = File(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_excel_source(file, excel_path)

        if not source_path.exists():
            raise HTTPException(status_code=400, detail=f"Excel model not found: {source_path}")

        out_path = export_sensitivity_range(
            source_path=source_path,
            sheet_name=sheet.strip(),
            range_address=range_address.strip(),
            output_dir=OUTPUT_DIR,
            output_name=output_name,
        )
        return {
            "ok": True,
            "source_path": _rel(source_path),
            "sheet": sheet,
            "range_address": range_address,
            "file": _rel(out_path),
            "download_url": f"/download/{out_path.name}",
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error("sensitivity_export", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/sensitivity/export-by-prompt")
async def export_sensitivity_by_prompt_api(
    prompt: str = Form(...),
    output_name: Optional[str] = Form(default=None),
    file: Optional[UploadFile] = File(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_excel_source(file, excel_path)

        if not source_path.exists():
            raise HTTPException(status_code=400, detail=f"Excel model not found: {source_path}")

        out_path, detected = export_sensitivity_by_prompt(
            source_path=source_path,
            prompt=prompt,
            output_dir=OUTPUT_DIR,
            output_name=output_name,
        )
        return {
            "ok": True,
            "source_path": _rel(source_path),
            "prompt": prompt,
            "detected": detected,
            "file": _rel(out_path),
            "download_url": f"/download/{out_path.name}",
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error("sensitivity_export_by_prompt", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/sensitivity/preview")
async def preview_sensitivity(
    sheet: str = Form(...),
    range_address: str = Form(...),
    file: Optional[UploadFile] = File(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_excel_source(file, excel_path)
        if not source_path.exists():
            raise HTTPException(status_code=400, detail=f"Excel model not found: {source_path}")

        preview = preview_sensitivity_range(
            source_path=source_path,
            sheet_name=sheet.strip(),
            range_address=range_address.strip(),
        )
        return {"ok": True, "source_path": _rel(source_path), "preview": preview}
    except HTTPException:
        raise
    except Exception as e:
        _log_error("sensitivity_preview", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/sensitivity/preview-by-prompt")
async def preview_sensitivity_by_prompt_api(
    prompt: str = Form(...),
    file: Optional[UploadFile] = File(default=None),
    excel_path: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_excel_source(file, excel_path)
        if not source_path.exists():
            raise HTTPException(status_code=400, detail=f"Excel model not found: {source_path}")

        preview = preview_sensitivity_by_prompt(source_path=source_path, prompt=prompt)
        return {"ok": True, "source_path": _rel(source_path), "prompt": prompt, "preview": preview}
    except HTTPException:
        raise
    except Exception as e:
        _log_error("sensitivity_preview_by_prompt", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/workbook-transfer/preview")
async def preview_workbook_transfer(
    original_file: Optional[UploadFile] = File(default=None),
    edited_file: Optional[UploadFile] = File(default=None),
    original_excel_path: Optional[str] = Form(default=None),
    edited_excel_path: Optional[str] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
    include_formula_cells: bool = Form(default=False),
    include_blank_cells: bool = Form(default=False),
) -> dict:
    _ensure_request_allowed()
    try:
        original_path = _resolve_uploaded_or_path(original_file, original_excel_path, "transfer")
        edited_path = _resolve_uploaded_or_path(edited_file, edited_excel_path, "transfer")
        preview = plan_workbook_value_transfer(
            original_path=original_path,
            edited_path=edited_path,
            sheet_name=sheet_name.strip() if sheet_name and sheet_name.strip() else None,
            include_formula_cells=include_formula_cells,
            include_blank_cells=include_blank_cells,
        )
        return {"ok": True, "preview": preview}
    except HTTPException:
        raise
    except Exception as e:
        _log_error("workbook_transfer_preview", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/workbook-transfer/export")
async def export_workbook_transfer(
    original_file: Optional[UploadFile] = File(default=None),
    edited_file: Optional[UploadFile] = File(default=None),
    original_excel_path: Optional[str] = Form(default=None),
    edited_excel_path: Optional[str] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
    include_formula_cells: bool = Form(default=False),
    include_blank_cells: bool = Form(default=False),
    output_name: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        original_path = _resolve_uploaded_or_path(original_file, original_excel_path, "transfer")
        edited_path = _resolve_uploaded_or_path(edited_file, edited_excel_path, "transfer")
        out_path, summary = export_workbook_value_transfer(
            original_path=original_path,
            edited_path=edited_path,
            output_dir=OUTPUT_DIR,
            sheet_name=sheet_name.strip() if sheet_name and sheet_name.strip() else None,
            include_formula_cells=include_formula_cells,
            include_blank_cells=include_blank_cells,
            output_name=output_name,
        )
        return {
            "ok": True,
            "file": _rel(out_path),
            "download_url": f"/download/{out_path.name}",
            "summary": summary,
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error("workbook_transfer_export", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/format-normalize/preview")
async def preview_format_normalize(
    source_file: Optional[UploadFile] = File(default=None),
    template_file: Optional[UploadFile] = File(default=None),
    source_excel_path: Optional[str] = Form(default=None),
    template_excel_path: Optional[str] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
    min_confidence: float = Form(default=0.72),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_uploaded_or_path(source_file, source_excel_path, "format_normalize")
        template_path = _resolve_uploaded_or_path(template_file, template_excel_path, "format_normalize")
        preview = preview_format_normalization(
            source_path=source_path,
            template_path=template_path,
            sheet_name=sheet_name.strip() if sheet_name and sheet_name.strip() else None,
            min_confidence=min_confidence,
        )
        return {"ok": True, "preview": preview}
    except HTTPException:
        raise
    except Exception as e:
        _log_error("format_normalize_preview", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.post("/api/format-normalize/export")
async def export_format_normalize(
    source_file: Optional[UploadFile] = File(default=None),
    template_file: Optional[UploadFile] = File(default=None),
    source_excel_path: Optional[str] = Form(default=None),
    template_excel_path: Optional[str] = Form(default=None),
    sheet_name: Optional[str] = Form(default=None),
    min_confidence: float = Form(default=0.72),
    output_name: Optional[str] = Form(default=None),
) -> dict:
    _ensure_request_allowed()
    try:
        source_path = _resolve_uploaded_or_path(source_file, source_excel_path, "format_normalize")
        template_path = _resolve_uploaded_or_path(template_file, template_excel_path, "format_normalize")
        out_path, summary = export_format_normalization(
            source_path=source_path,
            template_path=template_path,
            output_dir=OUTPUT_DIR,
            sheet_name=sheet_name.strip() if sheet_name and sheet_name.strip() else None,
            min_confidence=min_confidence,
            output_name=output_name,
        )
        return {
            "ok": True,
            "file": _rel(out_path),
            "download_url": f"/download/{out_path.name}",
            "summary": summary,
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error("format_normalize_export", e)
        raise HTTPException(status_code=422, detail=str(e))


@app.get("/download/{filename}")
def download(filename: str) -> FileResponse:
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, filename=path.name)
