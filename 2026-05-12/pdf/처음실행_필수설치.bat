@echo off
chcp 65001 >nul
setlocal

set "APP_DIR=%~dp0"

call :find_python
if not defined PYTHON_CMD (
    echo Python 3을 찾지 못했습니다.
    echo https://www.python.org/downloads/ 에서 Python 3을 설치한 뒤 다시 실행해 주세요.
    echo 설치할 때 Add python.exe to PATH 항목을 체크하면 좋습니다.
    pause
    exit /b 1
)

echo 필수 모듈 설치를 시작합니다.
%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 (
    %PYTHON_CMD% -m ensurepip --upgrade
)

%PYTHON_CMD% -m pip install --upgrade pip
%PYTHON_CMD% -m pip install -r "%APP_DIR%requirements.txt"

if errorlevel 1 (
    echo 설치 중 오류가 발생했습니다.
    echo 인터넷 연결 또는 Python 설치 상태를 확인해 주세요.
    pause
    exit /b 1
)

echo 설치가 완료되었습니다.
echo 이제 PDF_엑셀변환_실행.bat를 실행해 주세요.
pause
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
