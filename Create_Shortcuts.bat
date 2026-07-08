@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo.
  echo Python environment was not found.
  echo Please run Setup_First_Time.bat first.
  echo.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" scripts\create_windows_shortcuts.py
if errorlevel 1 (
  echo.
  echo Could not create shortcuts.
  echo Make sure dependencies are installed by running Setup_First_Time.bat.
  echo.
  pause
  exit /b 1
)

echo.
echo Shortcut is ready:
echo - Excel AI Controller.lnk
echo.
pause

endlocal
