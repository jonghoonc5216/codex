@echo off
chcp 65001 >nul
setlocal

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_desktop_shortcut.ps1"
if errorlevel 1 (
    echo.
    echo 바로가기 생성에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo 바탕화면 바로가기 생성이 완료되었습니다.
pause
