$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath = Join-Path $scriptDir "run_g2b_prespec_update.bat"
$taskName = "G2B Prespec Auto Update"

if (-not (Test-Path -LiteralPath $batPath)) {
    throw "실행 배치파일을 찾을 수 없습니다: $batPath"
}

$desktop = [Environment]::GetFolderPath("DesktopDirectory")
$shortcutPath = Join-Path $desktop "G2B_Prespec_Update.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = Join-Path $env:SystemRoot "System32\cmd.exe"
$shortcut.Arguments = "/c `"$batPath`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.Description = "Update G2B prespec Excel and open it"
$shortcut.Save()

$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Update G2B prespec data and open Excel at logon." `
    -Force | Out-Null

Write-Host "Desktop shortcut created: $shortcutPath"
Write-Host "Scheduled task registered: $taskName"
