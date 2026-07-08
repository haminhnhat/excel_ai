@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo.
  echo Python environment was not found.
  echo Please run Setup_First_Time.bat first, then run Excel AI Controller.bat again.
  echo.
  pause
  exit /b 1
)

if not exist ".env" (
  if exist ".env.example" (
    copy ".env.example" ".env" >nul
  )
)

echo.
echo Starting Excel AI Controller...
echo.
echo App URL: http://127.0.0.1:8000
echo Keep this window open while using the app.
echo.

start "" "http://127.0.0.1:8000"
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8000

echo.
echo The app stopped.
echo If this happened immediately, copy or screenshot the error message above.
echo Common causes: missing dependencies, broken Python environment, or port 8000 already in use.
echo.
pause

endlocal
