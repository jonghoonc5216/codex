@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "$startup = [Environment]::GetFolderPath('Startup'); $shortcutPath = Join-Path $startup '스티커메모 알람.lnk'; if (Test-Path -LiteralPath $shortcutPath) { Remove-Item -LiteralPath $shortcutPath -Force; Write-Output ('removed: ' + $shortcutPath) } else { Write-Output 'startup shortcut not found' }"
pause
