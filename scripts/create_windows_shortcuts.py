from __future__ import annotations

import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
ICON_DIR = BASE_DIR / "assets" / "icons"
APP_BAT = BASE_DIR / "Excel AI Controller.bat"


def _font(size: int) -> ImageFont.ImageFont:
    from PIL import ImageFont

    for name in ("segoeuib.ttf", "arialbd.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _rounded_icon(path: Path, label: str, background: tuple[int, int, int], accent: tuple[int, int, int]) -> None:
    from PIL import Image, ImageDraw

    size = 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((18, 18, size - 18, size - 18), radius=48, fill=background)
    draw.rounded_rectangle((34, 44, size - 34, size - 44), radius=28, fill=(255, 255, 255, 235))
    draw.rectangle((56, 74, size - 56, 90), fill=accent)
    for y in (116, 148, 180):
        draw.rectangle((56, y, size - 56, y + 12), fill=(215, 224, 235, 255))
    font = _font(72)
    bbox = draw.textbbox((0, 0), label, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    draw.rounded_rectangle((70, 86, size - 70, 170), radius=20, fill=accent)
    draw.text(((size - text_w) / 2, 124 - text_h / 2), label, fill="white", font=font)
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


def _create_shortcut(name: str, target: Path, icon: Path) -> None:
    try:
        import win32com.client
    except ImportError:
        import subprocess

        script = (
            "$shell = New-Object -ComObject WScript.Shell; "
            f"$shortcut = $shell.CreateShortcut('{(BASE_DIR / name).as_posix()}'); "
            f"$shortcut.TargetPath = '{target.as_posix()}'; "
            f"$shortcut.WorkingDirectory = '{BASE_DIR.as_posix()}'; "
            f"$shortcut.IconLocation = '{icon.as_posix()}'; "
            "$shortcut.Save()"
        )
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], check=True)
        return

    shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(BASE_DIR / name))
    shortcut.TargetPath = str(target)
    shortcut.WorkingDirectory = str(BASE_DIR)
    shortcut.IconLocation = str(icon)
    shortcut.Save()


def main() -> int:
    missing = [path.name for path in (APP_BAT,) if not path.exists()]
    if missing:
        print(f"Missing target file(s): {', '.join(missing)}", file=sys.stderr)
        return 1

    app_icon = ICON_DIR / "excel_ai_controller.ico"
    if not app_icon.exists():
        _rounded_icon(app_icon, "AI", (37, 99, 235, 255), (20, 184, 166, 255))

    _create_shortcut("Excel AI Controller.lnk", APP_BAT, app_icon)
    print("Created shortcut: Excel AI Controller.lnk")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
