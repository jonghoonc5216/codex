@echo off
chcp 65001 >nul
setlocal

where python >nul 2>nul
if errorlevel 1 (
    echo Python이 설치되어 있지 않습니다.
    echo 먼저 Python을 설치한 뒤 이 파일을 다시 실행해 주세요.
    pause
    exit /b 1
)

echo PDF를 JPEG로 변환하는 데 필요한 라이브러리를 설치합니다.
echo 설치 항목: pymupdf pillow
echo.

python -m pip install --user pymupdf pillow
if errorlevel 1 (
    echo.
    echo 설치에 실패했습니다. 인터넷 연결 또는 Python/pip 설치 상태를 확인해 주세요.
    pause
    exit /b 1
)

echo.
echo 설치가 완료되었습니다.
echo 이제 'PDF원클릭_JPEG변환.bat'을 실행하거나 PDF 파일을 그 위로 끌어다 놓으면 됩니다.
pause
