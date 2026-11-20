import hashlib
import json
from pathlib import Path
import shutil
import sys
import zipfile

import pytest

from scripts import build_portable_client as builder


@pytest.fixture
def isolated_builder(tmp_path, monkeypatch):
    root = tmp_path / "source"
    root.mkdir()
    shutil.copytree(builder.ROOT / "packaging", root / "packaging")
    for name, content in {
        "app/__init__.py": "",
        "web/index.html": "<html>Test application</html>",
        "assets/icons/excel_ai_controller.ico": "test icon",
        "models/private.xlsx": "PRIVATE_WORKBOOK_SENTINEL",
        "config/profiles/private/model_map.yaml": "PRIVATE_MAPPING_SENTINEL",
        ".env": "OPENAI_API_KEY=PRIVATE_ENV_SENTINEL",
        "outputs/private.json": "PRIVATE_OUTPUT_SENTINEL",
        "uploads/private.xlsx": "PRIVATE_UPLOAD_SENTINEL",
        "logs/private.log": "PRIVATE_LOG_SENTINEL",
    }.items():
        target = root / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
    cache = root / "outputs/portable/download-cache"
    cache.mkdir(parents=True)
    archive = cache / f"python-{builder.PYTHON_VERSION}-embed-amd64.zip"
    with zipfile.ZipFile(archive, "w") as runtime:
        runtime.writestr("python.exe", b"test-runtime")
    monkeypatch.setattr(builder, "ROOT", root)
    monkeypatch.setattr(builder, "PYTHON_SHA256", hashlib.sha256(archive.read_bytes()).hexdigest())
    # Runtime installation/execution is covered separately by the real ZIP verifier.
    monkeypatch.setattr(builder.subprocess, "run", lambda *args, **kwargs: None)
    return root


def test_public_build_excludes_all_private_source_data(isolated_builder, monkeypatch):
    monkeypatch.setattr(sys, "argv", ["build_portable_client.py", "--public"])
    builder.build()
    archive = next((isolated_builder / "outputs/portable").glob("Excel_AI_Controller_Public_*.zip"))
    with zipfile.ZipFile(archive) as package:
        names = package.namelist()
        prefix = names[0].split("/", 1)[0] + "/"
        metadata = json.loads(package.read(prefix + "build-info.json"))
        assert metadata["edition"] == "public"
        assert metadata["profile"] == ""
        assert "workbook_sha256" not in metadata
        assert "MODEL_PROFILE=" in package.read(prefix + ".env").decode().splitlines()
        assert not any(Path(name).suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"} for name in names)
        for name in names:
            assert b"SENTINEL" not in package.read(name)
            relative = name.removeprefix(prefix)
            if relative.startswith(("config/", "models/", "logs/", "uploads/", "outputs/")):
                assert name.endswith("/"), name


@pytest.mark.parametrize("option", ["--model", "--profile-map", "--profile"])
def test_public_build_rejects_client_options(option, monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["build_portable_client.py", "--public", option, "private"])
    with pytest.raises(SystemExit) as exc:
        builder.build()
    assert exc.value.code == 2
    assert "cannot include client" in capsys.readouterr().err


def test_client_build_still_requires_private_inputs(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["build_portable_client.py"])
    with pytest.raises(SystemExit) as exc:
        builder.build()
    assert exc.value.code == 2
    assert "--model" in capsys.readouterr().err


def test_client_build_preserves_explicit_inputs(isolated_builder, monkeypatch):
    monkeypatch.setattr(sys, "argv", [
        "build_portable_client.py", "--model", str(isolated_builder / "models/private.xlsx"),
        "--profile-map", str(isolated_builder / "config/profiles/private/model_map.yaml"),
        "--profile", "selected_profile",
    ])
    builder.build()
    archive = next((isolated_builder / "outputs/portable").glob("Excel_AI_Controller_Client_*.zip"))
    with zipfile.ZipFile(archive) as package:
        prefix = package.namelist()[0].split("/", 1)[0] + "/"
        metadata = json.loads(package.read(prefix + "build-info.json"))
        assert metadata["edition"] == "client" and metadata["profile"] == "selected_profile"
        assert package.read(prefix + "models/base_model.xlsx") == b"PRIVATE_WORKBOOK_SENTINEL"
        assert package.read(prefix + "config/profiles/selected_profile/model_map.yaml") == b"PRIVATE_MAPPING_SENTINEL"
        assert b"PRIVATE_ENV_SENTINEL" not in package.read(prefix + ".env")
