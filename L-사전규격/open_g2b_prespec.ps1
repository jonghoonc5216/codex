param(
    [Parameter(Mandatory = $true)]
    [string]$Uri
)

$ErrorActionPreference = "Stop"

function ConvertFrom-QueryString {
    param([string]$Query)

    $result = @{}
    if ([string]::IsNullOrWhiteSpace($Query)) {
        return $result
    }

    $rawQuery = $Query
    if ($rawQuery.StartsWith("?")) {
        $rawQuery = $rawQuery.Substring(1)
    }

    foreach ($pair in $rawQuery -split "&") {
        if ([string]::IsNullOrWhiteSpace($pair)) {
            continue
        }
        $parts = $pair -split "=", 2
        $key = [Uri]::UnescapeDataString($parts[0].Replace("+", " "))
        $value = ""
        if ($parts.Count -gt 1) {
            $value = [Uri]::UnescapeDataString($parts[1].Replace("+", " "))
        }
        $result[$key] = $value
    }
    return $result
}

function Get-ChromePath {
    $candidates = @(
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:LocalAppData "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }
    throw "Chrome 또는 Edge 실행 파일을 찾을 수 없습니다."
}

function Get-CdpItems {
    param([int]$Port)

    $data = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json" -TimeoutSec 2
    if ($null -ne $data.value) {
        return @($data.value)
    }
    return @($data)
}

function Get-CdpPage {
    param([int]$Port)

    $items = Get-CdpItems -Port $Port
    $page = @($items | Where-Object { $_.type -eq "page" -and $_.url -like "*g2b.go.kr*" } | Select-Object -First 1)
    if ($page.Count -gt 0) {
        return $page[0]
    }
    $page = @($items | Where-Object { $_.type -eq "page" } | Select-Object -First 1)
    if ($page.Count -gt 0) {
        return $page[0]
    }
    throw "Chrome DevTools 페이지를 찾을 수 없습니다."
}

function Invoke-Cdp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WebSocketUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [int]$TimeoutMs = 7000
    )

    $client = [System.Net.WebSockets.ClientWebSocket]::new()
    $connectCts = [Threading.CancellationTokenSource]::new($TimeoutMs)
    $client.ConnectAsync([Uri]$WebSocketUrl, $connectCts.Token).Wait()
    try {
        $json = $Payload | ConvertTo-Json -Compress -Depth 30
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $segment = [ArraySegment[byte]]::new($bytes)
        $sendCts = [Threading.CancellationTokenSource]::new($TimeoutMs)
        $client.SendAsync($segment, [Net.WebSockets.WebSocketMessageType]::Text, $true, $sendCts.Token).Wait()

        while ($true) {
            $receiveCts = [Threading.CancellationTokenSource]::new($TimeoutMs)
            $memory = [IO.MemoryStream]::new()
            do {
                $buffer = [byte[]]::new(65536)
                $bufferSegment = [ArraySegment[byte]]::new($buffer)
                $result = $client.ReceiveAsync($bufferSegment, $receiveCts.Token).Result
                $memory.Write($buffer, 0, $result.Count)
            } until ($result.EndOfMessage)

            $text = [Text.Encoding]::UTF8.GetString($memory.ToArray())
            $memory.Dispose()
            $message = $text | ConvertFrom-Json
            if ($message.id -eq $Payload.id) {
                return $message
            }
        }
    } finally {
        $client.Dispose()
    }
}

function Wait-G2BReady {
    param(
        [int]$Port,
        [int]$Seconds = 45
    )

    $deadline = (Get-Date).AddSeconds($Seconds)
    $expression = @"
Boolean(
  window.com &&
  typeof window.com.gfnOpenMenu === "function" &&
  window._gcm &&
  window._gcm.Global &&
  typeof window._gcm.Global.getMainLoaded === "function" &&
  window._gcm.Global.getMainLoaded()
)
"@
    while ((Get-Date) -lt $deadline) {
        try {
            $page = Get-CdpPage -Port $Port
            $result = Invoke-Cdp -WebSocketUrl $page.webSocketDebuggerUrl -Payload @{
                id = 101
                method = "Runtime.evaluate"
                params = @{
                    expression = $expression
                    returnByValue = $true
                }
            } -TimeoutMs 3000
            if ($result.result.result.value -eq $true) {
                return $page
            }
        } catch {
            Start-Sleep -Milliseconds 700
        }
        Start-Sleep -Milliseconds 700
    }
    throw "나라장터 초기화가 제한 시간 안에 완료되지 않았습니다."
}

$queryText = ""
if ($Uri.Contains("?")) {
    $queryText = ($Uri -split "\?", 2)[1]
}
$query = ConvertFrom-QueryString $queryText
$specNo = $query["bfSpecRgstNo"]
if ([string]::IsNullOrWhiteSpace($specNo)) {
    $specNo = $query["bfSpecRegNo"]
}
if ([string]::IsNullOrWhiteSpace($specNo)) {
    $specNo = $query["g2bPrespecRegNo"]
}
if ([string]::IsNullOrWhiteSpace($specNo)) {
    throw "사전규격등록번호가 없습니다: $Uri"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$chrome = Get-ChromePath
$profileDir = Join-Path $scriptDir "chrome_g2b_prespec_profile"
$port = 9322
$homeUrl = "https://www.g2b.go.kr/"

if (-not (Test-Path -LiteralPath $profileDir)) {
    New-Item -ItemType Directory -Force -Path $profileDir | Out-Null
}

$cdpReady = $false
try {
    Invoke-RestMethod -Uri "http://127.0.0.1:$port/json/version" -TimeoutSec 1 | Out-Null
    $cdpReady = $true
} catch {
    $cdpReady = $false
}

if (-not $cdpReady) {
    Start-Process -FilePath $chrome -ArgumentList @(
        "--remote-debugging-port=$port",
        "--user-data-dir=$profileDir",
        "--no-first-run",
        "--no-default-browser-check",
        "--new-window",
        $homeUrl
    )
}

$deadline = (Get-Date).AddSeconds(30)
do {
    try {
        Invoke-RestMethod -Uri "http://127.0.0.1:$port/json/version" -TimeoutSec 1 | Out-Null
        break
    } catch {
        Start-Sleep -Milliseconds 700
    }
} while ((Get-Date) -lt $deadline)

$page = Get-CdpPage -Port $port
Invoke-Cdp -WebSocketUrl $page.webSocketDebuggerUrl -Payload @{
    id = 1
    method = "Page.navigate"
    params = @{
        url = $homeUrl
    }
} | Out-Null

$page = Wait-G2BReady -Port $port

$param = @{
    bfSpecRegNo = $specNo
    bfSpecRgstNo = $specNo
}
if (-not [string]::IsNullOrWhiteSpace($query["prcmBsneSeCd"])) {
    $param.prcmBsneSeCd = $query["prcmBsneSeCd"]
}
if (-not [string]::IsNullOrWhiteSpace($query["bsnsDivNm"])) {
    $param.bsnsDivNm = $query["bsnsDivNm"]
}
$paramJson = $param | ConvertTo-Json -Compress -Depth 10
$openExpression = @"
(() => {
  const param = $paramJson;
  window.com.gfnOpenMenu("PRVA004_02", { isHistory: true, param });
  return true;
})()
"@

Invoke-Cdp -WebSocketUrl $page.webSocketDebuggerUrl -Payload @{
    id = 2
    method = "Runtime.evaluate"
    params = @{
        expression = $openExpression
        returnByValue = $true
    }
} | Out-Null
