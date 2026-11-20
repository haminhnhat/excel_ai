"""Build an allowlisted Windows x64 client ZIP with a private Python runtime."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import urllib.request
import zipfile

ROOT = Path(__file__).resolve().parents[1]
PYTHON_VERSION = "3.13.7"
PYTHON_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip"
# SHA-256 from the official python.org Sigstore bundle, fetched over HTTPS.
PYTHON_SHA256 = "f6cca216a359be84797cabb54149ce5e062afb16cc7567eb7fc51cacb2d86b65"


def digest(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def copy_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def build() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "outputs/portable")
    parser.add_argument("--public", action="store_true", help="Build without any workbook or customer profile")
    parser.add_argument("--model", type=Path, help="Private .xlsx workbook to include locally")
    parser.add_argument("--profile-map", type=Path, help="Private model_map.yaml to include locally")
    parser.add_argument("--profile", help="Profile identifier inside the client package")
    args = parser.parse_args()
    if sys.platform != "win32" or sys.version_info[:2] != (3, 13) or struct.calcsize("P") != 8:
        parser.error("Build with 64-bit Python 3.13 on Windows.")
    model = profile = None
    if args.public:
        if args.model is not None or args.profile_map is not None or args.profile is not None:
            parser.error("Public packages cannot include client --model, --profile-map or --profile options.")
        profile_name = ""
    else:
        if args.model is None or args.profile_map is None:
            parser.error("Client packages require --model and --profile-map; use --public for an empty app.")
        profile_name = args.profile or "client_model"
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", profile_name):
            parser.error("Profile must contain only letters, digits, underscores or hyphens.")
        model = args.model.resolve()
        profile = args.profile_map.resolve()
        if model.suffix.lower() != ".xlsx":
            parser.error("The default workbook must be an .xlsx file.")
        for required in (model, profile):
            if not required.is_file():
                parser.error(f"Missing required client data: {required}")

    output = args.output_dir.resolve()
    output.mkdir(parents=True, exist_ok=True)
    cache = output / "download-cache"
    cache.mkdir(exist_ok=True)
    runtime_zip = cache / f"python-{PYTHON_VERSION}-embed-amd64.zip"
    if not runtime_zip.exists():
        print(f"Downloading {PYTHON_URL}", flush=True)
        urllib.request.urlretrieve(PYTHON_URL, runtime_zip)
    if digest(runtime_zip) != PYTHON_SHA256:
        raise RuntimeError(f"Python archive checksum mismatch: {runtime_zip}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    edition = "Public" if args.public else "Client"
    package = output / f"Excel_AI_Controller_{edition}_{stamp}"
    package.mkdir()  # Never overwrite a previously built client package.
    runtime = package / "runtime"
    with zipfile.ZipFile(runtime_zip) as archive:
        archive.extractall(runtime)
    (runtime / "python313._pth").write_text(
        "python313.zip\n.\nLib/site-packages\n..\nimport site\n", encoding="ascii"
    )
    requirements = ROOT / "packaging/requirements-portable.txt"
    subprocess.run([
        sys.executable, "-m", "pip", "install", "--disable-pip-version-check",
        "--only-binary=:all:", "--no-compile", "--no-deps",
        "--target", str(runtime / "Lib/site-packages"), "-r", str(requirements),
    ], check=True)

    for source in sorted((ROOT / "app").glob("*.py")):
        copy_file(source, package / "app" / source.name)
    for source in sorted((ROOT / "web").rglob("*")):
        if source.is_file() and source.suffix.lower() in {".html", ".css", ".js", ".svg", ".png", ".ico", ".woff2"}:
            copy_file(source, package / source.relative_to(ROOT))
    copy_file(ROOT / "assets/icons/excel_ai_controller.ico", package / "assets/icons/excel_ai_controller.ico")
    if not args.public:
        copy_file(model, package / "models/base_model.xlsx")
        copy_file(profile, package / "config/profiles" / profile_name / "model_map.yaml")
    for name in ("launch.py", "Excel AI Controller.bat", "READ_ME.txt"):
        copy_file(ROOT / "packaging/client" / name, package / name)
    if args.public:
        copy_file(ROOT / "packaging/public/READ_ME.txt", package / "READ_ME.txt")
    env_template = (ROOT / "packaging/client/client.env").read_text(encoding="utf-8")
    (package / ".env").write_text(
        env_template.replace("MODEL_PROFILE=client_model", f"MODEL_PROFILE={profile_name}"), encoding="utf-8"
    )
    copy_file(requirements, package / requirements.name)
    empty_dirs = ("outputs", "uploads", "logs", "models", "config/profiles")
    for name in empty_dirs:
        (package / name).mkdir(parents=True, exist_ok=True)

    metadata = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "platform": "Windows x64", "python_version": PYTHON_VERSION,
        "python_url": PYTHON_URL, "python_sha256": PYTHON_SHA256,
        "edition": edition.lower(), "profile": profile_name,
        "configuration": "Local mock parser; no API keys; no license administration",
    }
    if model is not None:
        metadata["workbook_sha256"] = digest(model)
    (package / "build-info.json").write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    subprocess.run([str(runtime / "python.exe"), "-I", "-B", str(package / "launch.py"), "--check"],
                   cwd=package, check=True, timeout=60)
    # Keep third-party notices, but omit caches, tests and machine-specific console scripts.
    excluded = {"__pycache__", "tests", "test", "bin", "Scripts"}
    excluded_suffixes = {".pyc"}
    if args.public:
        # xlwings ships example workbooks; they are not needed by this app.
        excluded_suffixes.update({".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".xlam"})
    files = [path for path in sorted(package.rglob("*")) if path.is_file()
             and not excluded.intersection(path.relative_to(package).parts)
             and path.suffix.lower() not in excluded_suffixes]
    manifest = package / "SHA256SUMS.txt"
    manifest.write_text("".join(f"{digest(path)}  {path.relative_to(package).as_posix()}\n" for path in files), encoding="utf-8")
    files.append(manifest)
    zip_path = package.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in files:
            archive.write(path, f"{package.name}/{path.relative_to(package).as_posix()}")
        for name in empty_dirs:
            archive.writestr(f"{package.name}/{name}/", "")
    print(json.dumps({"zip": str(zip_path), "sha256": digest(zip_path), "bytes": zip_path.stat().st_size}, indent=2), flush=True)


if __name__ == "__main__":
    build()
