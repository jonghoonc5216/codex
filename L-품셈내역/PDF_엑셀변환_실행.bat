@echo off
chcp 65001 >nul
setlocal

set "APP_DIR=%~dp0"

if not exist "%APP_DIR%auto_converter_gui.ps1" (
    echo 실행 파일을 찾지 못했습니다.
    echo 이 파일은 PDF 엑셀변환 폴더 안에서 실행해야 합니다.
    pause
    exit /b 1
)

call :find_python
if not defined PYTHON_CMD (
    echo Python 3을 찾지 못했습니다.
    echo 처음실행_필수설치.bat를 먼저 실행하거나 Python 3을 설치해 주세요.
    pause
    exit /b 1
)

%PYTHON_CMD% -c "import openpyxl, pdfplumber" >nul 2>nul
if errorlevel 1 (
    echo 필수 Python 모듈이 설치되어 있지 않습니다.
    echo 처음실행_필수설치.bat를 먼저 실행해 주세요.
    pause
    exit /b 1
)

start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP_DIR%auto_converter_gui.ps1"
exit /b 0

:find_python
where python >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)

where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)

exit /b 0
