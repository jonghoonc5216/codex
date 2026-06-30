$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$extensionDir = Join-Path $scriptDir "g2b_prespec_deeplink_extension"

if (-not (Test-Path -LiteralPath (Join-Path $extensionDir "manifest.json"))) {
    throw "브라우저 확장 폴더를 찾을 수 없습니다: $extensionDir"
}

$browserCandidates = @(
    @{ Name = "Chrome"; Path = (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"); Url = "chrome://extensions/" },
    @{ Name = "Chrome"; Path = (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"); Url = "chrome://extensions/" },
    @{ Name = "Chrome"; Path = (Join-Path $env:LocalAppData "Google\Chrome\Application\chrome.exe"); Url = "chrome://extensions/" },
    @{ Name = "Edge"; Path = (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"); Url = "edge://extensions/" },
    @{ Name = "Edge"; Path = (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"); Url = "edge://extensions/" }
) | Where-Object { $_.Path -and (Test-Path -LiteralPath $_.Path) }

if ($browserCandidates.Count -gt 0) {
    Start-Process -FilePath $browserCandidates[0].Path -ArgumentList $browserCandidates[0].Url
} else {
    Start-Process "chrome://extensions/"
}

Start-Process explorer.exe -ArgumentList "`"$extensionDir`""

Write-Host ""
Write-Host "브라우저 확장 설치 안내"
Write-Host "1. 열린 브라우저 확장 프로그램 화면에서 '개발자 모드'를 켭니다."
Write-Host "2. '압축해제된 확장 프로그램을 로드합니다'를 누릅니다."
Write-Host "3. 열린 탐색기의 아래 폴더를 선택합니다."
Write-Host "   $extensionDir"
Write-Host "4. 설치 후 엑셀의 사업명을 클릭하면 같은 브라우저에서 나라장터 상세조회로 이동합니다."
