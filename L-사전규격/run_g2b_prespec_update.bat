@echo off
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"
python "%SCRIPT_DIR%g2b_prespec_updater.py"
exit /b %ERRORLEVEL%
