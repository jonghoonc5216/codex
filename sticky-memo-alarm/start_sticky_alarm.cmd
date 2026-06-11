@echo off
setlocal
set "APP_DIR=%~dp0"
set "PYTHONW="
for %%P in (pythonw.exe) do set "PYTHONW=%%~$PATH:P"

if defined PYTHONW (
  start "" "%PYTHONW%" "%APP_DIR%sticky_alarm.py"
) else (
  start "Sticky Memo Alarm" /min python "%APP_DIR%sticky_alarm.py"
)

echo Sticky Memo Alarm monitor started.
echo You can close this window.
