@echo off
chcp 65001 >nul
setlocal

where python >nul 2>nul
if errorlevel 1 (
    echo Python이 설치되어 있지 않습니다.
    echo Python 설치 후 다시 실행해 주세요.
    pause
    exit /b 1
)

pushd "%~dp0" >nul
if errorlevel 1 (
    echo 프로그램 폴더로 이동하지 못했습니다.
    pause
    exit /b 1
)

python "pdf_to_jpeg.py" %*
set "RESULT=%ERRORLEVEL%"
popd >nul
exit /b %RESULT%
