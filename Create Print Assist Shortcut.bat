@echo off
setlocal

set "APP_DIR=%~dp0"
set "APP_DIR=%APP_DIR:~0,-1%"

set "TARGET=%APP_DIR%\PrintAssist.exe"
if not exist "%TARGET%" set "TARGET=%APP_DIR%\dist\PrintAssist\PrintAssist.exe"
if not exist "%TARGET%" set "TARGET=%APP_DIR%\dist\PrintAssist.exe"

if not exist "%TARGET%" (
    echo Print Assist executable was not found.
    echo.
    echo Expected one of:
    echo   %APP_DIR%\PrintAssist.exe
    echo   %APP_DIR%\dist\PrintAssist\PrintAssist.exe
    echo   %APP_DIR%\dist\PrintAssist.exe
    echo.
    echo Build the Windows executable first with:
    echo   python -m PyInstaller --noconfirm --clean PrintAssist.spec
    echo.
    if not "%PRINT_ASSIST_NO_PAUSE%"=="1" pause
    exit /b 1
)

set "ICON=%APP_DIR%\print_assist\assets\print-assist.ico"
if not exist "%ICON%" set "ICON=%TARGET%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$shortcutDir=$env:PRINT_ASSIST_SHORTCUT_DIR; if ([string]::IsNullOrWhiteSpace($shortcutDir)) { $shortcutDir=[Environment]::GetFolderPath('Desktop') }; if (-not (Test-Path -LiteralPath $shortcutDir)) { New-Item -ItemType Directory -Path $shortcutDir | Out-Null }; $shortcut=Join-Path $shortcutDir 'Print Assist.lnk'; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=$env:TARGET; $link.WorkingDirectory=Split-Path -Parent $env:TARGET; $link.IconLocation=$env:ICON; $link.Description='Open Print Assist'; $link.Save(); Write-Host ('Created shortcut: ' + $shortcut)"

if errorlevel 1 (
    echo.
    echo The shortcut could not be created.
    if not "%PRINT_ASSIST_NO_PAUSE%"=="1" pause
    exit /b 1
)

echo.
echo Print Assist shortcut created successfully.
if not "%PRINT_ASSIST_NO_PAUSE%"=="1" pause
