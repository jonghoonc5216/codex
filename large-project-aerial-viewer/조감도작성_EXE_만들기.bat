@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [조감도작성] PyInstaller 확인 중...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
  echo PyInstaller 설치에 실패했습니다.
  pause
  exit /b 1
)

echo [조감도작성] exe 빌드 중...
python -m PyInstaller --clean --noconfirm "조감도작성.spec"
if errorlevel 1 (
  echo exe 빌드에 실패했습니다.
  pause
  exit /b 1
)

echo.
echo 완료: %cd%\dist\조감도작성.exe
echo 이 파일을 실행하면 조감도작성 서버와 브라우저가 열립니다.
pause
