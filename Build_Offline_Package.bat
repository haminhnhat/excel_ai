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
  echo Install Python 3.11 or newer, then run this file again.
  echo.
  pause
  exit /b 1
)

if not exist "vendor" mkdir "vendor"
if not exist "vendor\wheels" mkdir "vendor\wheels"

echo.
echo Building offline package cache in vendor\wheels...
echo This step needs internet and should be done by admin/IT.
echo.

%PYTHON_CMD% -m pip download --dest "vendor\wheels" -r requirements.txt
if errorlevel 1 goto failed

echo.
echo Offline package cache is ready.
echo Copy this whole project folder to user machines.
echo On user machines, run Setup_First_Time.bat, then Start_App.bat.
echo.
pause
exit /b 0

:failed
echo.
echo Could not build the offline package cache.
echo Check internet access, firewall/proxy, and Python installation.
echo.
pause
exit /b 1

endlocal
