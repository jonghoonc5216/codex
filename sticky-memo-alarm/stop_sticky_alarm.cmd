@echo off
setlocal
set "APP_DIR=%~dp0"
set "PIDFILE=%APP_DIR%sticky_alarm.pid"

if exist "%PIDFILE%" (
  set /p ALARM_PID=<"%PIDFILE%"
  if defined ALARM_PID (
    taskkill /PID %ALARM_PID% /F >nul 2>nul
  )
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'python.exe' -or $_.Name -eq 'pythonw.exe') -and $_.CommandLine -like '*sticky_alarm.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>nul

echo Sticky Memo Alarm monitor stopped.
