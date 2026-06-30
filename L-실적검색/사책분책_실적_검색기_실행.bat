@echo off
for %%F in ("%~dp0*-v3.exe") do (
  start "" "%%~fF"
  exit /b
)
echo v3 exe not found.
pause
