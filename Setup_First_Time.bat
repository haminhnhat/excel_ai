@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=python"

if "%PYTHON_CMD%"=="" (
  where py >nul 2>nul
  if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if "%PYTHON_CMD%"=="" (
  echo.
  echo Python was not found.
  echo Please install Python 3.11 or newer from https://www.python.org/downloads/
  echo During install, tick "Add python.exe to PATH".
  echo.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating Python environment...
  %PYTHON_CMD% -m venv .venv
  if errorlevel 1 (
    echo.
    echo Could not create Python environment.
    pause
    exit /b 1
  )
)

echo Installing app dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto install_failed

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto install_failed

if not exist ".env" (
  if exist ".env.example" (
    copy ".env.example" ".env" >nul
  )
)

echo.
echo Setup complete.
echo You can now double-click Start_App.bat.
echo.
pause
exit /b 0

:install_failed
echo.
echo Dependency install failed.
echo Check your internet connection, then run Setup_First_Time.bat again.
echo.
pause
exit /b 1

endlocal
