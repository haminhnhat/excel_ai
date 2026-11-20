"""Start the portable client using its private Python runtime."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import socket
import sys
import threading
import time
import webbrowser
import winreg

ROOT = Path(__file__).resolve().parent


def excel_available() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            return True
    except OSError:
        return False


def bind_local_socket(port: int) -> socket.socket:
    candidates = [0] if port == 0 else range(port, min(port + 32, 65536))
    for candidate in candidates:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            sock.bind(("127.0.0.1", candidate))
            sock.listen(128)
            return sock
        except OSError as exc:
            sock.close()
            if exc.winerror != 10048:
                raise
    raise RuntimeError("No free local port. Close an older app window and try again.")


def main() -> None:
    if not sys.flags.no_user_site:
        raise RuntimeError("Open Excel AI Controller.bat to use the isolated bundled runtime.")
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()
    if not 0 <= args.port <= 65535:
        parser.error("port must be between 0 and 65535")
    os.chdir(ROOT)
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    from dotenv import load_dotenv

    # A portable copy owns its configuration, not the launching shell's model path.
    for name in ("MODEL_MAP_PATH", "OPENAI_API_KEY", "OLLAMA_URL", "OLLAMA_MODEL"):
        os.environ.pop(name, None)
    if not (ROOT / ".env").is_file():
        raise RuntimeError("Missing .env. Extract the complete ZIP before starting the app.")
    load_dotenv(ROOT / ".env", override=True)
    for folder in ("outputs", "uploads", "logs", "config/profiles"):
        path = ROOT / folder
        path.mkdir(parents=True, exist_ok=True)
        probe = path / f".write-test-{os.getpid()}"
        try:
            probe.write_text("", encoding="ascii")
        except OSError as exc:
            raise RuntimeError("Extract the app to a writable folder, such as Documents.") from exc
        finally:
            probe.unlink(missing_ok=True)

    import pythoncom
    import pywintypes
    import win32com.client
    import xlwings
    import uvicorn
    from app.main import app

    model = ROOT / os.environ["MODEL_PATH"]
    profile = ROOT / "config/profiles" / os.environ["MODEL_PROFILE"] / "model_map.yaml"
    if not model.is_file() or not profile.is_file():
        raise RuntimeError("Default workbook/profile missing. Extract the complete ZIP again.")
    if args.check:
        print(json.dumps({
            "ok": True,
            "python": sys.version,
            "executable": sys.executable,
            "isolated": bool(sys.flags.isolated),
            "user_site_disabled": bool(sys.flags.no_user_site),
            "excel_registered": excel_available(),
            "profile": os.environ["MODEL_PROFILE"],
            "model_exists": model.is_file(),
            "sys_path": sys.path,
        }, indent=2))
        return

    if not excel_available():
        print("WARNING: Microsoft Excel desktop is required to calculate workbook formulas.")
    sock = bind_local_socket(args.port)
    port = sock.getsockname()[1]
    url = f"http://127.0.0.1:{port}"
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="info"))

    def open_when_ready() -> None:
        deadline = time.monotonic() + 60
        while not server.started and not server.should_exit and time.monotonic() < deadline:
            time.sleep(0.1)
        if server.started:
            webbrowser.open(url)

    if not args.no_browser:
        threading.Thread(target=open_when_ready, daemon=True).start()
    print(f"Excel AI Controller: {url}", flush=True)
    print("Keep this window open. Press Ctrl+C to stop the app.", flush=True)
    try:
        server.run(sockets=[sock])
    finally:
        sock.close()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
