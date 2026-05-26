from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from .ai_parser import parse_user_command
from .config_loader import list_profiles, load_model_map_for_profile, resolve_model_map_path
from .excel_controller import ExcelController
from .utils import format_value
from .validator import ValidationError, validate_action_plan

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

app = FastAPI(title="Excel AI Controller", version="0.1.0")

WEB_DIR = BASE_DIR / "web"
if WEB_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")


def _resolve_model_path() -> Path:
    model_path = Path(os.getenv("MODEL_PATH", BASE_DIR / "models" / "base_model.xlsx"))
    if not model_path.is_absolute():
        model_path = BASE_DIR / model_path
    return model_path


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    index_path = WEB_DIR / "index.html"
    if index_path.exists():
        return index_path.read_text(encoding="utf-8")
    return "<h1>Excel AI Controller</h1>"


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


@app.post("/api/scenario")
async def run_scenario(
    command: str = Form(...),
    file: Optional[UploadFile] = File(default=None),
    engine: Optional[str] = Form(default=None),
    model_profile: Optional[str] = Form(default=None),
) -> dict:
    try:
        model_map, model_map_path = load_model_map_for_profile(profile=model_profile)

        if file is not None and file.filename:
            safe_name = Path(file.filename).name
            source_path = UPLOAD_DIR / safe_name
            with source_path.open("wb") as f:
                shutil.copyfileobj(file.file, f)
        else:
            source_path = _resolve_model_path()

        if not source_path.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Excel model not found: {source_path}. Upload a file or set MODEL_PATH in .env.",
            )

        plan = parse_user_command(command, model_map)
        plan = validate_action_plan(plan, model_map)

        controller = ExcelController(
            model_map=model_map,
            output_dir=OUTPUT_DIR,
            audit_log_path=AUDIT_LOG_PATH,
            engine=engine or os.getenv("EXCEL_ENGINE", "auto"),
        )
        result = controller.run_scenario(source_path, plan)
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
            "download_url": f"/download/{scenario_file.name}",
        }
    except ValidationError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/download/{filename}")
def download(filename: str) -> FileResponse:
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, filename=path.name)
