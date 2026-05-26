@echo off
setlocal
set "APP_DIR=C:\Users\saman\Documents\Codex\2026-05-12\pdf"
if exist "%APP_DIR%\auto_converter_gui.ps1" (
    cd /d "%APP_DIR%"
    start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP_DIR%\auto_converter_gui.ps1"
    exit /b
)
echo Cannot find PDF formula converter working files.
pause
