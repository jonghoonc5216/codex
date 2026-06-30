$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $scriptDir "open_g2b_prespec.ps1"
$powershell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path -LiteralPath $launcher)) {
    throw "실행 파일을 찾을 수 없습니다: $launcher"
}
if (-not (Test-Path -LiteralPath $powershell)) {
    throw "PowerShell 실행 파일을 찾을 수 없습니다: $powershell"
}

$protocolRoot = "HKCU:\Software\Classes\g2bprespec"
$commandRoot = Join-Path $protocolRoot "shell\open\command"

New-Item -Path $commandRoot -Force | Out-Null
Set-ItemProperty -Path $protocolRoot -Name "(default)" -Value "URL:G2B Prespec Detail"
Set-ItemProperty -Path $protocolRoot -Name "URL Protocol" -Value ""
Set-ItemProperty -Path $commandRoot -Name "(default)" -Value "`"$powershell`" -NoProfile -ExecutionPolicy Bypass -File `"$launcher`" `"%1`""

Write-Host "G2B prespec detail protocol registered: g2bprespec://"
