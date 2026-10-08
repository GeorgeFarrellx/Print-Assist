@echo off
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Create Print Assist Shortcut.ps1" -AppDir "%~dp0."

if errorlevel 1 (
    echo.
    echo The shortcut could not be created.
    if not "%PRINT_ASSIST_NO_PAUSE%"=="1" pause
    exit /b 1
)

echo.
echo Print Assist shortcut created successfully.
if not "%PRINT_ASSIST_NO_PAUSE%"=="1" pause
