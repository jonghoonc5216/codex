@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p = Join-Path (Get-Location) 'auto_converter_gui.ps1'; $s = Get-Content -LiteralPath $p -Raw -Encoding UTF8; Invoke-Expression $s"
if errorlevel 1 pause
