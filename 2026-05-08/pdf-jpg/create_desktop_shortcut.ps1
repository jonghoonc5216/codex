$ErrorActionPreference = "Stop"

$toolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $toolDir "PDF원클릭_JPEG변환.bat"
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "PDF원클릭 JPEG 변환.lnk"

if (-not (Test-Path -LiteralPath $target)) {
    throw "변환 실행 파일을 찾을 수 없습니다: $target"
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $target
$shortcut.WorkingDirectory = $toolDir
$shortcut.IconLocation = "$env:SystemRoot\System32\imageres.dll,109"
$shortcut.Description = "PDF 파일을 JPEG 이미지로 변환"
$shortcut.Save()

Write-Host "생성 완료: $shortcutPath"
