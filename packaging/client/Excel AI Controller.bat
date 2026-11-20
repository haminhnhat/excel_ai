@echo off
setlocal
title Excel AI Controller
pushd "%~dp0"
if not exist "runtime\python.exe" (
    echo Missing bundled Python. Extract the complete ZIP before opening the app.
    pause
    popd
    exit /b 1
)
"%~dp0runtime\python.exe" -I -B "%~dp0launch.py" %*
set "APP_EXIT=%ERRORLEVEL%"
if not "%APP_EXIT%"=="0" (
    echo.
    echo The app could not start. Please send the error above to your support contact.
    pause
)
popd
exit /b %APP_EXIT%
