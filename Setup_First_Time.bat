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

if exist "vendor\wheels\*.whl" (
  echo Offline package cache found: vendor\wheels
  echo Installing without internet...
  ".venv\Scripts\python.exe" -m pip install --no-index --find-links "vendor\wheels" -r requirements.txt
  if errorlevel 1 goto offline_install_failed
) else (
  echo No offline package cache found.
  echo Installing from internet...
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  if errorlevel 1 goto install_failed

  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto install_failed
)

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

:offline_install_failed
echo.
echo Offline dependency install failed.
echo The package cache in vendor\wheels may be incomplete or built for a different Python version.
echo Ask admin/IT to run Build_Offline_Package.bat on a working machine, then copy the updated folder again.
echo.
pause
exit /b 1

endlocal
