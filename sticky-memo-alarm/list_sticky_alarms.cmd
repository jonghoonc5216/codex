@echo off
setlocal
set "APP_DIR=%~dp0"
python "%APP_DIR%sticky_alarm.py" --list
pause
