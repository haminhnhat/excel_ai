"""Verify a client ZIP after relocating it, without Python on the child PATH."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import socket
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("archive", type=Path)
    parser.add_argument("--smoke-script", type=Path)
    args = parser.parse_args()
    work = Path(tempfile.mkdtemp(prefix="Client test ", dir=ROOT / "outputs"))
    report: dict = {"archive": str(args.archive.resolve()), "test_directory": str(work), "checks": []}

    def passed(name: str) -> None:
        report["checks"].append(name)
        print(f"PASS: {name}", flush=True)

    with zipfile.ZipFile(args.archive) as archive:
        names = archive.namelist()
        assert archive.testzip() is None
        forbidden = {".git", ".venv", ".agents", "secrets", "__pycache__", "distributor"}
        assert not any(forbidden.intersection(Path(name).parts) for name in names)
        assert not any(re.search(r"license[_-](admin|private)|private[_-]key", name, re.I) for name in names)
        archive.extractall(work)
    package = next(path for path in work.iterdir() if path.is_dir())
    for line in (package / "SHA256SUMS.txt").read_text(encoding="utf-8").splitlines():
        expected, name = line.split("  ", 1)
        assert hashlib.sha256((package / name).read_bytes()).hexdigest() == expected, name
    def read_settings(path: Path) -> dict[str, str]:
        return dict(line.split("=", 1) for line in path.read_text(encoding="utf-8").splitlines()
                    if line.strip() and not line.startswith("#"))

    settings = read_settings(package / ".env")
    template = read_settings(ROOT / "packaging/client/client.env")
    metadata = json.loads((package / "build-info.json").read_text(encoding="utf-8"))
    public = metadata.get("edition") == "public"
    assert settings.keys() == template.keys()
    for key in settings.keys() - {"MODEL_PATH", "MODEL_PROFILE"}:
        assert settings[key] == template[key], key
    assert settings["MODEL_PROFILE"] == metadata["profile"]
    model = (package / settings["MODEL_PATH"]).resolve()
    assert model.is_relative_to(package / "models")
    if public:
        assert not settings["MODEL_PROFILE"] and not model.exists()
        assert "workbook_sha256" not in metadata
        workbook_suffixes = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".xlam"}
        assert not any(Path(name).suffix.lower() in workbook_suffixes for name in names)
        assert not any(path.is_file() for path in (package / "config").rglob("*"))
        assert not any(path.is_file() for path in (package / "models").rglob("*"))
    else:
        assert model.is_file()
    for name in ("logs", "uploads", "outputs"):
        assert not list((package / name).iterdir())
    passed("Archive integrity, complete manifest, clean configuration and no private/admin files")

    env = {key.upper(): value for key, value in os.environ.items()}
    env.update({"PATH": os.pathsep.join([str(Path(env["SYSTEMROOT"]) / "System32"), env["SYSTEMROOT"]]),
                "PYTHONPATH": str(work / "nonexistent-pythonpath"),
                "PYTHONHOME": str(work / "nonexistent-pythonhome"),
                "MODEL_PATH": "invalid-inherited-model.xlsx", "MODEL_MAP_PATH": "invalid-map.yaml"})
    python = package / "runtime/python.exe"
    check = subprocess.run([str(python), "-I", "-B", str(package / "launch.py"), "--check"],
                           cwd=work, env=env, capture_output=True, text=True, encoding="utf-8", check=True, timeout=60)
    diagnostic = json.loads(check.stdout)
    assert diagnostic["isolated"] and diagnostic["user_site_disabled"]
    assert all(Path(path).is_relative_to(package) for path in diagnostic["sys_path"])
    report["runtime"] = diagnostic
    passed("Bundled Python imports without system Python, user-site or inherited configuration")

    bat = subprocess.run(f'"{env["COMSPEC"]}" /d /s /c ""{package / "Excel AI Controller.bat"}" --check"',
                         cwd=work, env=env, capture_output=True, text=True, encoding="utf-8", timeout=60)
    assert bat.returncode == 0, bat.stdout + bat.stderr
    assert json.loads(bat.stdout)["user_site_disabled"]
    passed("Named BAT launcher works from a relocated path containing spaces")

    if not public:
        original_hash = hashlib.sha256(model.read_bytes()).hexdigest()
        assert original_hash == metadata["workbook_sha256"]
    with socket.socket() as occupied:
        occupied.bind(("127.0.0.1", 0))
        occupied.listen()
        port = occupied.getsockname()[1]
        with (work / "server.log").open("w", encoding="utf-8") as log:
            process = subprocess.Popen([str(python), "-I", "-B", str(package / "launch.py"),
                                        "--no-browser", "--port", str(port)],
                                       cwd=work, env=env, stdout=log, stderr=subprocess.STDOUT)
            try:
                url = None
                deadline = time.monotonic() + 40
                while time.monotonic() < deadline:
                    assert process.poll() is None, (work / "server.log").read_text(encoding="utf-8")
                    match = re.search(r"Excel AI Controller: (http://127\.0\.0\.1:\d+)",
                                      (work / "server.log").read_text(encoding="utf-8"))
                    if match:
                        url = match[1]
                        try:
                            with urllib.request.urlopen(url + "/health", timeout=1) as response:
                                health = json.load(response)
                            break
                        except OSError:
                            pass
                    time.sleep(.2)
                else:
                    raise RuntimeError("Portable server did not become ready")
                assert int(url.rsplit(":", 1)[1]) > port
                assert health["model_profile"] == settings["MODEL_PROFILE"]
                with urllib.request.urlopen(url, timeout=10) as response:
                    assert b"Excel AI Controller" in response.read()
                passed("HTTP interface and health work; occupied port is handled automatically")
                if public:
                    assert health["available_profiles"] == []
                    result = subprocess.run([
                        str(python), "-I", "-B", str(ROOT / "scripts/verify_public_onboarding.py"),
                        "--url", url, "--work-dir", str(work / "synthetic-input"),
                    ], cwd=package, env=env, capture_output=True, text=True, encoding="utf-8",
                        errors="replace", timeout=180)
                    (work / "public-onboarding.log").write_text(result.stdout + result.stderr, encoding="utf-8")
                    print(result.stdout, flush=True)
                    assert result.returncode == 0, result.stderr
                    passed("Public first-run upload, mapping, Excel scenario, download and profile reuse")
                else:
                    body = urllib.parse.urlencode({"command": "Cho toi NPV va IRR hien tai",
                                               "model_profile": settings["MODEL_PROFILE"], "engine": "xlwings"}).encode()
                    with urllib.request.urlopen(url + "/api/scenario", body, timeout=90) as response:
                        result = json.load(response)
                    assert result["ok"], result
                    assert result["result"]["outputs"]
                    assert not result["result"]["applied_changes"]
                    with urllib.request.urlopen(url + result["download_url"], timeout=10) as response:
                        assert response.read(2) == b"PK"
                    assert hashlib.sha256(model.read_bytes()).hexdigest() == original_hash
                    passed("Real Excel scenario and download work with default workbook; source unchanged")
            finally:
                process.terminate()
                try:
                    process.wait(timeout=15)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=10)

    if args.smoke_script:
        target = package / "outputs/verify_restored_app.py"
        shutil.copy2(args.smoke_script, target)
        smoke_env = env.copy()
        smoke_env.pop("MODEL_MAP_PATH", None)
        for engine in ("openpyxl", "xlwings"):
            result = subprocess.run([str(python), "-I", "-B", str(target), engine],
                                    cwd=package, env=smoke_env, capture_output=True, text=True,
                                    encoding="utf-8", errors="replace", timeout=180)
            (work / f"smoke-{engine}.log").write_text(result.stdout + result.stderr, encoding="utf-8")
            print(result.stdout, flush=True)
            assert result.returncode == 0, result.stderr
            passed(f"Full API smoke suite with bundled runtime: {engine}")
    (work / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Verification report: {work / 'report.json'}", flush=True)


if __name__ == "__main__":
    main()
