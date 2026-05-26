@echo off
chcp 65001 >nul
setlocal

set "APP_DIR=D:\Ai 프로그래밍\엑셀산식자동화"

if not exist "%APP_DIR%\auto_converter_gui.ps1" (
    echo 자동화 프로그램 파일을 찾지 못했습니다.
    echo 확인 경로: %APP_DIR%
    pause
    exit /b 1
)

cd /d "%APP_DIR%"
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP_DIR%\auto_converter_gui.ps1"
exit /b 0
