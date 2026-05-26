Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$script:Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:PythonScript = Join-Path $script:Root "pdf_formula_converter.py"
$script:TemplatePath = Join-Path $script:Root "template.xlsx"
if (-not (Test-Path -LiteralPath $script:TemplatePath)) {
    $script:TemplatePath = "C:\Users\saman\Desktop\Ai 실습\_새 폴더\260127 실시계획인가 금액산정.xlsx"
}
$script:PdfPaths = New-Object System.Collections.Generic.List[string]
$script:Signatures = @{}
$script:Busy = $false
$script:UpdatingFactors = $false
$script:HasConverted = $false
$script:LastOutputPath = ""
$script:DefaultOutputFolder = [Environment]::GetFolderPath("DesktopDirectory")
$script:DefaultOutputName = "PDF 원본명칭"
$script:OutputNameAuto = $true
$script:UpdatingOutputName = $false
$script:LaborRateLabels = @("기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자", "고급숙련", "중급숙련", "초급숙련")
$script:LaborRateYear = 2026
$script:LaborRateByYear = @{
    2023 = @("432440", "335638", "282545", "261571", "205686", "240947", "220894", "186909")
    2024 = @("446055", "346855", "293799", "272915", "213496", "252328", "238259", "194029")
    2025 = @("452718", "358273", "300980", "284046", "223644", "267012", "240710", "204392")
    2026 = @("467217", "373353", "310884", "295138", "235459", "281075", "250087", "218142")
}
$script:LaborRateDefaults = $script:LaborRateByYear[$script:LaborRateYear]
$script:LaborRateBoxes = @()

function To-InvariantNumber([double]$Value) {
    return $Value.ToString("0.############", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Resolve-PythonRunner {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $python) {
        return [pscustomobject]@{
            Exe = $python.Source
            Args = @()
        }
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        return [pscustomobject]@{
            Exe = $py.Source
            Args = @("-3")
        }
    }

    return $null
}

function Invoke-Python([object[]]$Arguments) {
    if ($null -eq $script:PythonRunner) {
        throw "Python 실행 프로그램을 찾지 못했습니다. 처음실행_필수설치.bat를 실행하거나 Python 3을 설치해 주세요."
    }

    $fullArgs = @($script:PythonRunner.Args) + @($Arguments)
    $output = & $script:PythonRunner.Exe @fullArgs 2>&1
    $script:LastPythonExitCode = $LASTEXITCODE
    return $output
}

$script:PythonRunner = Resolve-PythonRunner
$script:LastPythonExitCode = 0

function Parse-Number([string]$Text, [string]$Name) {
    $clean = ($Text -replace ",", "").Trim()
    $value = 0.0
    $ok = [double]::TryParse(
        $clean,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$value
    )
    if (-not $ok) {
        throw "$Name 값이 숫자가 아닙니다: $Text"
    }
    return $value
}

function Parse-NumberFromText([string]$Text, [string]$Name) {
    $raw = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "$Name 값이 비어 있습니다."
    }
    $match = [regex]::Match($raw, "[-+]?\d[\d,]*(\.\d+)?")
    if (-not $match.Success) {
        throw "$Name 값에서 숫자를 찾지 못했습니다: $Text"
    }
    return Parse-Number $match.Value $Name
}

function Parse-Ratio([string]$Text, [string]$Name) {
    $raw = $Text.Trim()
    $value = Parse-NumberFromText $raw $Name
    if ($raw.Contains("%")) {
        return $value / 100.0
    }
    if ($value -gt 1) {
        return $value / 100.0
    }
    return $value
}

function Evaluate-FactorFormula([string]$Text, [double]$Area) {
    $formula = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($formula)) {
        throw "직접 보정계수 산식이 비어 있습니다."
    }

    $number = 0.0
    $isNumber = [double]::TryParse(
        ($formula -replace ",", ""),
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )
    if ($isNumber) {
        return $number
    }

    $raw = Invoke-Python @(
        $script:PythonScript,
        "--eval-factor",
        "--area", (To-InvariantNumber $Area),
        "--factor-formula", $formula
    )
    $exit = $script:LastPythonExitCode
    $rawText = ($raw | Out-String).Trim()
    if ($exit -ne 0) {
        throw $rawText
    }
    return Parse-Number $rawText "직접 보정계수 산식 결과"
}

function Get-BaseFactor([double]$Area) {
    if ($Area -lt 1000) {
        $unitArea = 0.1
    } else {
        $unitArea = [Math]::Round($Area / 10000.0, 2)
    }
    return [Math]::Round([Math]::Pow($unitArea, 0.6), 3)
}

function Get-FactorInputs([bool]$RequireValid) {
    $targetArea = Parse-NumberFromText $areaBox.Text "대상면적"
    $baseArea = Parse-NumberFromText $baseAreaBox.Text "기준면적"
    $exponent = Parse-Number $exponentBox.Text "승수"
    $factor1 = Parse-Number $conversionFactor1Box.Text "환산계수1"
    $factor2 = Parse-Number $conversionFactor2Box.Text "환산계수2"
    $ratio = Parse-Ratio $ratioBox.Text "적용비율"

    if ($baseArea -le 0) {
        if ($RequireValid) {
            throw "기준면적은 0보다 큰 값이어야 합니다."
        }
        return $null
    }

    return [pscustomobject]@{
        TargetArea = $targetArea
        TargetAreaText = $areaBox.Text.Trim()
        BaseArea = $baseArea
        BaseAreaText = $baseAreaBox.Text.Trim()
        Exponent = $exponent
        ConversionFactor1 = $factor1
        ConversionFactor2 = $factor2
        Ratio = $ratio
    }
}

function Set-ErrorText([string]$Message) {
    $errorBox.Text = $Message
}

function Clear-ErrorText {
    $errorBox.Text = "오류 없음"
}

function Get-FriendlyError([string]$Raw) {
    if ($Raw -match "PermissionError|액세스가 거부") {
        return "결과 엑셀 파일이 열려 있어서 저장하지 못했을 가능성이 큽니다.`r`n결과 엑셀을 닫은 뒤 다시 변환해 주세요.`r`n`r`n상세 내용:`r`n$Raw"
    }
    if ($Raw -match "기본업무|기준인원수|표를 찾지 못") {
        return "PDF에서 변환 가능한 표를 찾지 못했습니다.`r`n스캔 이미지 PDF이거나 표 구조가 기존 양식과 다를 수 있습니다.`r`n`r`n상세 내용:`r`n$Raw"
    }
    if ($Raw -match "템플릿|template|No such file|없습니다") {
        return "자동화용 템플릿 엑셀 파일을 찾지 못했거나 읽지 못했습니다.`r`n프로그램 폴더의 template.xlsx 또는 원본 템플릿 파일을 확인해 주세요.`r`n`r`n상세 내용:`r`n$Raw"
    }
    if ($Raw -match "python|Python") {
        return "Python 변환 엔진을 실행하지 못했습니다.`r`nPython 설치 또는 pdf_formula_converter.py 파일을 확인해 주세요.`r`n`r`n상세 내용:`r`n$Raw"
    }
    return "변환 중 오류가 발생했습니다.`r`n`r`n상세 내용:`r`n$Raw"
}

function Get-OutputPath {
    $name = $outputNameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        return ""
    }
    if ([IO.Path]::GetExtension($name).Length -eq 0) {
        $name = "$name.xlsx"
    }
    $folder = $outputFolderBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        return ""
    }
    return Join-Path $folder $name
}

function Get-AvailableOutputPath([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }
    $folder = Split-Path -Parent $Path
    $baseName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [IO.Path]::GetExtension($Path)
    for ($index = 1; $index -lt 10000; $index++) {
        $candidate = Join-Path $folder ("{0} ({1}){2}" -f $baseName, $index, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }
    throw "사용 가능한 저장 파일명을 만들지 못했습니다."
}

function Get-PdfDefaultOutputName {
    if ($script:PdfPaths.Count -gt 0) {
        return [IO.Path]::GetFileNameWithoutExtension($script:PdfPaths[0])
    }
    return $script:DefaultOutputName
}

function Set-OutputNameFromPdf {
    if (-not $script:OutputNameAuto) {
        return
    }
    $script:UpdatingOutputName = $true
    try {
        $outputNameBox.Text = Get-PdfDefaultOutputName
    } finally {
        $script:UpdatingOutputName = $false
    }
}

function Get-LaborRatesArgument {
    $values = New-Object System.Collections.Generic.List[string]
    $hasValue = $false
    for ($index = 0; $index -lt $script:LaborRateBoxes.Count; $index++) {
        $text = $script:LaborRateBoxes[$index].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            [void]$values.Add("")
            continue
        }
        $value = Parse-Number $text $script:LaborRateLabels[$index]
        [void]$values.Add((To-InvariantNumber $value))
        $hasValue = $true
    }
    if (-not $hasValue) {
        return $null
    }
    return ($values -join ",")
}

function Set-LaborRatesForYear([int]$Year) {
    if (-not $script:LaborRateByYear.ContainsKey($Year)) {
        return
    }
    $script:LaborRateYear = $Year
    $rates = $script:LaborRateByYear[$Year]
    for ($index = 0; $index -lt $script:LaborRateBoxes.Count; $index++) {
        if ($index -lt $rates.Count) {
            $script:LaborRateBoxes[$index].Text = $rates[$index]
        }
    }
}

function Refresh-OutputLabel {
    $path = Get-OutputPath
    if ([string]::IsNullOrWhiteSpace($path)) {
        $statusLabel.Text = "저장명칭과 저장위치를 입력해 주세요."
    } else {
        $statusLabel.Text = "대기 중"
    }
}

function Refresh-PdfList {
    $pdfList.Items.Clear()
    foreach ($path in $script:PdfPaths) {
        [void]$pdfList.Items.Add($path)
    }
    Refresh-OutputLabel
}

function Get-Signatures {
    $result = @{}
    foreach ($path in $script:PdfPaths) {
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path
            $result[$path] = "$($item.LastWriteTimeUtc.Ticks)|$($item.Length)"
        }
    }
    return $result
}

function Update-Signatures {
    $script:Signatures = Get-Signatures
}

function Signatures-Changed {
    $current = Get-Signatures
    if ($current.Count -ne $script:Signatures.Count) {
        $script:Signatures = $current
        return $true
    }
    foreach ($key in $current.Keys) {
        if (-not $script:Signatures.ContainsKey($key)) {
            $script:Signatures = $current
            return $true
        }
        if ($script:Signatures[$key] -ne $current[$key]) {
            $script:Signatures = $current
            return $true
        }
    }
    return $false
}

function Update-Factors([bool]$ResetManual) {
    if ($script:UpdatingFactors) {
        return
    }
    $script:UpdatingFactors = $true
    try {
        $inputs = Get-FactorInputs $false
        if ($null -eq $inputs) {
            $calcFactorBox.Text = ""
            $resultFactorBox.Text = ""
            Clear-ErrorText
            return
        }
        $workFactor = [Math]::Round([Math]::Pow(($inputs.TargetArea / $inputs.BaseArea), $inputs.Exponent), 2)
        $calcFactorBox.Text = To-InvariantNumber $workFactor
        $result = [Math]::Round($workFactor * $inputs.ConversionFactor1 * $inputs.ConversionFactor2 * $inputs.Ratio, 2)
        $resultFactorBox.Text = To-InvariantNumber $result
        Clear-ErrorText
    } catch {
        $calcFactorBox.Text = ""
        $resultFactorBox.Text = ""
        Set-ErrorText $_.Exception.Message
    } finally {
        $script:UpdatingFactors = $false
    }
}

function Add-PdfFiles([string[]]$Files, [bool]$AutoConvert) {
    $added = 0
    foreach ($file in $Files) {
        if ((Test-Path -LiteralPath $file) -and ([IO.Path]::GetExtension($file).ToLowerInvariant() -eq ".pdf")) {
            $full = [IO.Path]::GetFullPath($file)
            if (-not $script:PdfPaths.Contains($full)) {
                $script:PdfPaths.Add($full)
                $added += 1
            }
        }
    }
    Refresh-PdfList
    Update-Signatures
    if ($added -eq 0) {
        Set-ErrorText "추가된 PDF가 없습니다. PDF 파일만 변환 대상에 넣을 수 있습니다."
        return
    }
    Set-OutputNameFromPdf
    $statusLabel.Text = "PDF $added개 추가됨 - 변환 실행을 누르면 저장됩니다."
    if ($AutoConvert) {
        Invoke-Conversion "PDF 추가 자동 변환"
    }
}

function Invoke-Conversion([string]$Reason) {
    if ($script:Busy) {
        return
    }
    if ($script:PdfPaths.Count -eq 0) {
        Set-ErrorText "변환할 PDF가 없습니다. 변환 대상 PDF 영역에 PDF 파일을 끌어놓아 주세요."
        return
    }
    if (-not (Test-Path -LiteralPath $script:PythonScript)) {
        Set-ErrorText "변환 엔진 파일을 찾지 못했습니다: $script:PythonScript"
        return
    }
    if (-not (Test-Path -LiteralPath $script:TemplatePath)) {
        Set-ErrorText "템플릿 엑셀 파일을 찾지 못했습니다: $script:TemplatePath"
        return
    }

    try {
        $inputs = Get-FactorInputs $true
        $laborRates = Get-LaborRatesArgument
    } catch {
        Set-ErrorText $_.Exception.Message
        return
    }

    $requestedOutput = Get-OutputPath
    if ([string]::IsNullOrWhiteSpace($requestedOutput)) {
        Set-ErrorText "저장명칭 또는 저장위치가 비어 있습니다. 저장명칭과 저장위치를 입력해 주세요."
        return
    }
    $outputName = $outputNameBox.Text.Trim()
    if ($outputName -match '[\\/:*?"<>|]') {
        Set-ErrorText "저장명칭에는 \ / : * ? `" < > | 문자를 사용할 수 없습니다."
        return
    }
    if ([IO.Path]::GetExtension($requestedOutput).ToLowerInvariant() -ne ".xlsx") {
        Set-ErrorText "저장명칭에 확장자를 넣는 경우 .xlsx만 사용할 수 있습니다."
        return
    }
    $outputFolder = Split-Path -Parent $requestedOutput
    if (-not [string]::IsNullOrWhiteSpace($outputFolder) -and -not (Test-Path -LiteralPath $outputFolder)) {
        New-Item -ItemType Directory -Force -Path $outputFolder | Out-Null
    }

    if ($Reason -match "PDF 변경 자동 변환" -and $script:HasConverted -and -not [string]::IsNullOrWhiteSpace($script:LastOutputPath)) {
        $output = $script:LastOutputPath
    } else {
        $output = Get-AvailableOutputPath $requestedOutput
    }

    $script:Busy = $true
    $convertButton.Enabled = $false
    $clearButton.Enabled = $false
    $statusLabel.Text = "변환 중..."
    Clear-ErrorText
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $args = @(
            $script:PythonScript,
            "--template", $script:TemplatePath,
            "--output", $output,
            "--area", (To-InvariantNumber $inputs.TargetArea),
            "--area-text", $inputs.TargetAreaText,
            "--base-area", (To-InvariantNumber $inputs.BaseArea),
            "--base-area-text", $inputs.BaseAreaText,
            "--exponent", (To-InvariantNumber $inputs.Exponent),
            "--conversion-factor1", (To-InvariantNumber $inputs.ConversionFactor1),
            "--conversion-factor2", (To-InvariantNumber $inputs.ConversionFactor2),
            "--ratio", (To-InvariantNumber $inputs.Ratio),
            "--labor-rate-year", ([string][int]$laborYearBox.Value)
        )
        if (-not [string]::IsNullOrWhiteSpace($laborRates)) {
            $args += @("--labor-rates", $laborRates)
        }
        $args += @($script:PdfPaths)
        $raw = Invoke-Python $args
        $exit = $script:LastPythonExitCode
        $rawText = ($raw | Out-String).Trim()
        if ($exit -ne 0) {
            throw $rawText
        }
        $statusLabel.Text = "$Reason 완료: $($script:PdfPaths.Count)개 PDF 반영"
        $lastResultLabel.Text = "마지막 변환: " + (Get-Date -Format "HH:mm:ss")
        $script:HasConverted = $true
        $script:LastOutputPath = $output
        Clear-ErrorText
        Update-Signatures
    } catch {
        $statusLabel.Text = "변환 실패"
        Set-ErrorText (Get-FriendlyError ($_.Exception.Message))
    } finally {
        $convertButton.Enabled = $true
        $clearButton.Enabled = $true
        $script:Busy = $false
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "PDF 표 엑셀 산식 자동 변환기"
$form.Size = New-Object System.Drawing.Size(900, 720)
$form.MinimumSize = New-Object System.Drawing.Size(860, 660)
$form.StartPosition = "CenterScreen"

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = "Fill"
$main.Padding = New-Object System.Windows.Forms.Padding(14)
$main.ColumnCount = 1
$main.RowCount = 6
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 106))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 108))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
$form.Controls.Add($main)

$title = New-Object System.Windows.Forms.Label
$title.Dock = "Fill"
$title.Text = "PDF를 넣으면 엑셀 산식 결과가 자동 갱신됩니다."
$title.Font = New-Object System.Drawing.Font("맑은 고딕", 10, [System.Drawing.FontStyle]::Bold)
$main.Controls.Add($title, 0, 0)

$factorPanel = New-Object System.Windows.Forms.TableLayoutPanel
$factorPanel.Dock = "Fill"
$factorPanel.ColumnCount = 8
$factorPanel.RowCount = 3
$factorPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
$factorPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
$factorPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
for ($i = 0; $i -lt 8; $i++) {
    $factorPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 12.5))) | Out-Null
}
$main.Controls.Add($factorPanel, 0, 1)

$labels = @("대상면적", "기준면적", "승수", "작업량 계수", "환산계수1", "환산계수2", "적용비율", "적용 보정계수")
for ($i = 0; $i -lt $labels.Count; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labels[$i]
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleLeft"
    $label.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
    $factorPanel.Controls.Add($label, $i, 0)
}

$areaBox = New-Object System.Windows.Forms.TextBox
$areaBox.Text = "10,000㎡"
$baseAreaBox = New-Object System.Windows.Forms.TextBox
$baseAreaBox.Text = "10,000㎡"
$exponentBox = New-Object System.Windows.Forms.TextBox
$exponentBox.Text = "0.6"
$calcFactorBox = New-Object System.Windows.Forms.TextBox
$calcFactorBox.ReadOnly = $true
$conversionFactor1Box = New-Object System.Windows.Forms.TextBox
$conversionFactor1Box.Text = "1"
$conversionFactor2Box = New-Object System.Windows.Forms.TextBox
$conversionFactor2Box.Text = "1"
$ratioBox = New-Object System.Windows.Forms.TextBox
$ratioBox.Text = "100%"
$resultFactorBox = New-Object System.Windows.Forms.TextBox
$resultFactorBox.ReadOnly = $true

$boxes = @($areaBox, $baseAreaBox, $exponentBox, $calcFactorBox, $conversionFactor1Box, $conversionFactor2Box, $ratioBox, $resultFactorBox)
for ($i = 0; $i -lt $boxes.Count; $i++) {
    $boxes[$i].Dock = "Fill"
    $boxes[$i].Font = New-Object System.Drawing.Font("맑은 고딕", 10)
    $boxes[$i].Margin = New-Object System.Windows.Forms.Padding(0, 2, 10, 2)
    $factorPanel.Controls.Add($boxes[$i], $i, 1)
}

$factorNoteLabel = New-Object System.Windows.Forms.Label
$factorNoteLabel.Text = " * 대상면적에 대한 작업량 계수 a=(대상면적/기준면적)^(승수)"
$factorNoteLabel.Dock = "Fill"
$factorNoteLabel.TextAlign = "MiddleLeft"
$factorNoteLabel.ForeColor = [System.Drawing.Color]::Firebrick
$factorNoteLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$factorPanel.Controls.Add($factorNoteLabel, 0, 2)
$factorPanel.SetColumnSpan($factorNoteLabel, 8)

$outputPanel = New-Object System.Windows.Forms.TableLayoutPanel
$outputPanel.Dock = "Fill"
$outputPanel.ColumnCount = 3
$outputPanel.RowCount = 3
$outputPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 86))) | Out-Null
$outputPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$outputPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 96))) | Out-Null
$outputPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$outputPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$outputPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$main.Controls.Add($outputPanel, 0, 2)

$outputNameLabel = New-Object System.Windows.Forms.Label
$outputNameLabel.Text = "저장명칭"
$outputNameLabel.Dock = "Fill"
$outputNameLabel.TextAlign = "MiddleLeft"
$outputNameLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$outputPanel.Controls.Add($outputNameLabel, 0, 0)

$outputNameBox = New-Object System.Windows.Forms.TextBox
$outputNameBox.Text = $script:DefaultOutputName
$outputNameBox.Dock = "Fill"
$outputNameBox.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$outputNameBox.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 2)
$outputPanel.Controls.Add($outputNameBox, 1, 0)
$outputPanel.SetColumnSpan($outputNameBox, 2)

$outputFolderLabel = New-Object System.Windows.Forms.Label
$outputFolderLabel.Text = "저장위치"
$outputFolderLabel.Dock = "Fill"
$outputFolderLabel.TextAlign = "MiddleLeft"
$outputFolderLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$outputPanel.Controls.Add($outputFolderLabel, 0, 1)

$outputFolderBox = New-Object System.Windows.Forms.TextBox
$outputFolderBox.Text = $script:DefaultOutputFolder
$outputFolderBox.Dock = "Fill"
$outputFolderBox.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$outputFolderBox.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 2)
$outputPanel.Controls.Add($outputFolderBox, 1, 1)

$outputBrowseButton = New-Object System.Windows.Forms.Button
$outputBrowseButton.Text = "저장 위치"
$outputBrowseButton.Dock = "Fill"
$outputBrowseButton.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 2)
$outputPanel.Controls.Add($outputBrowseButton, 2, 1)

$networkNoticeLabel = New-Object System.Windows.Forms.Label
$networkNoticeLabel.Text = "네트워크 저장은 불가"
$networkNoticeLabel.Dock = "Fill"
$networkNoticeLabel.TextAlign = "MiddleLeft"
$networkNoticeLabel.ForeColor = [System.Drawing.Color]::Firebrick
$networkNoticeLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$outputPanel.Controls.Add($networkNoticeLabel, 1, 2)
$outputPanel.SetColumnSpan($networkNoticeLabel, 2)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Fill"
$buttonPanel.FlowDirection = "LeftToRight"
$buttonPanel.WrapContents = $false
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$main.Controls.Add($buttonPanel, 0, 3)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "PDF 추가"
$addButton.Width = 110
$addButton.Height = 34
$buttonPanel.Controls.Add($addButton)

$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Text = "변환 실행"
$convertButton.Width = 120
$convertButton.Height = 34
$buttonPanel.Controls.Add($convertButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "목록 비우기"
$clearButton.Width = 120
$clearButton.Height = 34
$buttonPanel.Controls.Add($clearButton)

$contentPanel = New-Object System.Windows.Forms.TableLayoutPanel
$contentPanel.Dock = "Fill"
$contentPanel.ColumnCount = 2
$contentPanel.RowCount = 1
$contentPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 66))) | Out-Null
$contentPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 34))) | Out-Null
$main.Controls.Add($contentPanel, 0, 4)

$contentLeftPanel = New-Object System.Windows.Forms.TableLayoutPanel
$contentLeftPanel.Dock = "Fill"
$contentLeftPanel.ColumnCount = 1
$contentLeftPanel.RowCount = 2
$contentLeftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 62))) | Out-Null
$contentLeftPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 38))) | Out-Null
$contentLeftPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$contentPanel.Controls.Add($contentLeftPanel, 0, 0)

$dropGroup = New-Object System.Windows.Forms.GroupBox
$dropGroup.Text = "변환 대상 PDF"
$dropGroup.Dock = "Fill"
$dropGroup.AllowDrop = $true
$contentLeftPanel.Controls.Add($dropGroup, 0, 0)

$pdfList = New-Object System.Windows.Forms.ListBox
$pdfList.Dock = "Fill"
$pdfList.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$pdfList.AllowDrop = $true
$dropGroup.Controls.Add($pdfList)

$errorGroup = New-Object System.Windows.Forms.GroupBox
$errorGroup.Text = "오류 설명"
$errorGroup.Dock = "Fill"
$contentLeftPanel.Controls.Add($errorGroup, 0, 1)

$errorBox = New-Object System.Windows.Forms.TextBox
$errorBox.Dock = "Fill"
$errorBox.Multiline = $true
$errorBox.ScrollBars = "Vertical"
$errorBox.ReadOnly = $true
$errorBox.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$errorGroup.Controls.Add($errorBox)

$laborGroup = New-Object System.Windows.Forms.GroupBox
$laborGroup.Text = ""
$laborGroup.Dock = "Fill"
$laborGroup.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$contentPanel.Controls.Add($laborGroup, 1, 0)

$laborOuterPanel = New-Object System.Windows.Forms.TableLayoutPanel
$laborOuterPanel.Dock = "Fill"
$laborOuterPanel.ColumnCount = 1
$laborOuterPanel.RowCount = 2
$laborOuterPanel.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)
$laborOuterPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
$laborOuterPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$laborGroup.Controls.Add($laborOuterPanel)

$laborHeaderPanel = New-Object System.Windows.Forms.TableLayoutPanel
$laborHeaderPanel.Dock = "Fill"
$laborHeaderPanel.ColumnCount = 3
$laborHeaderPanel.RowCount = 1
$laborHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100))) | Out-Null
$laborHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 82))) | Out-Null
$laborHeaderPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$laborOuterPanel.Controls.Add($laborHeaderPanel, 0, 0)

$laborYearLabel = New-Object System.Windows.Forms.Label
$laborYearLabel.Text = "노임단가 설정"
$laborYearLabel.Dock = "Fill"
$laborYearLabel.TextAlign = "MiddleLeft"
$laborYearLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$laborHeaderPanel.Controls.Add($laborYearLabel, 0, 0)

$laborYearBox = New-Object System.Windows.Forms.NumericUpDown
$laborYearBox.Minimum = 2023
$laborYearBox.Maximum = 2026
$laborYearBox.Value = $script:LaborRateYear
$laborYearBox.Dock = "Fill"
$laborYearBox.TextAlign = "Right"
$laborYearBox.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$laborHeaderPanel.Controls.Add($laborYearBox, 1, 0)

$laborSourceLabel = New-Object System.Windows.Forms.Label
$laborSourceLabel.Text = "년 건설 부문"
$laborSourceLabel.Dock = "Fill"
$laborSourceLabel.TextAlign = "MiddleLeft"
$laborSourceLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 8.5)
$laborSourceLabel.ForeColor = [System.Drawing.Color]::DimGray
$laborHeaderPanel.Controls.Add($laborSourceLabel, 2, 0)

$laborPanel = New-Object System.Windows.Forms.TableLayoutPanel
$laborPanel.Dock = "Fill"
$laborPanel.ColumnCount = 2
$laborPanel.RowCount = $script:LaborRateLabels.Count
$laborPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 112))) | Out-Null
$laborPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
for ($i = 0; $i -lt $script:LaborRateLabels.Count; $i++) {
    $laborPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 12.5))) | Out-Null
}
$laborOuterPanel.Controls.Add($laborPanel, 0, 1)

for ($i = 0; $i -lt $script:LaborRateLabels.Count; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $script:LaborRateLabels[$i]
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleLeft"
    $label.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
    $laborPanel.Controls.Add($label, 0, $i)

    $rateBox = New-Object System.Windows.Forms.TextBox
    $rateBox.Text = $script:LaborRateDefaults[$i]
    $rateBox.Dock = "Fill"
    $rateBox.TextAlign = "Right"
    $rateBox.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
    $rateBox.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    $laborPanel.Controls.Add($rateBox, 1, $i)
    $script:LaborRateBoxes += $rateBox
}

$bottomPanel = New-Object System.Windows.Forms.TableLayoutPanel
$bottomPanel.Dock = "Fill"
$bottomPanel.ColumnCount = 3
$bottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 55))) | Out-Null
$bottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$bottomPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20))) | Out-Null
$main.Controls.Add($bottomPanel, 0, 5)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock = "Fill"
$statusLabel.Text = "대기 중"
$bottomPanel.Controls.Add($statusLabel, 0, 0)

$lastResultLabel = New-Object System.Windows.Forms.Label
$lastResultLabel.Dock = "Fill"
$lastResultLabel.TextAlign = "MiddleRight"
$bottomPanel.Controls.Add($lastResultLabel, 1, 0)

$departmentLabel = New-Object System.Windows.Forms.Label
$departmentLabel.Dock = "Fill"
$departmentLabel.Text = "조경레저부"
$departmentLabel.TextAlign = "MiddleRight"
$departmentLabel.Font = New-Object System.Drawing.Font("맑은 고딕", 9)
$bottomPanel.Controls.Add($departmentLabel, 2, 0)

$dropHandler = {
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
}
$dropCompleteHandler = {
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Add-PdfFiles $files $false
}
$dropGroup.Add_DragEnter($dropHandler)
$dropGroup.Add_DragDrop($dropCompleteHandler)
$pdfList.Add_DragEnter($dropHandler)
$pdfList.Add_DragDrop($dropCompleteHandler)
$form.Add_DragEnter($dropHandler)
$form.Add_DragDrop($dropCompleteHandler)

$addButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
    $dialog.Multiselect = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-PdfFiles $dialog.FileNames $false
    }
})

$outputBrowseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "결과 엑셀을 저장할 폴더를 선택하세요."
    if (-not [string]::IsNullOrWhiteSpace($outputFolderBox.Text) -and (Test-Path -LiteralPath $outputFolderBox.Text)) {
        $dialog.SelectedPath = $outputFolderBox.Text
    } else {
        $dialog.SelectedPath = $script:DefaultOutputFolder
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputFolderBox.Text = $dialog.SelectedPath
        Clear-ErrorText
    }
})

$convertButton.Add_Click({ Invoke-Conversion "수동 변환" })
$clearButton.Add_Click({
    $script:PdfPaths.Clear()
    $script:HasConverted = $false
    $script:LastOutputPath = ""
    $script:OutputNameAuto = $true
    Set-OutputNameFromPdf
    Refresh-PdfList
    Update-Signatures
    $statusLabel.Text = "목록을 비웠습니다."
    Clear-ErrorText
})

$areaBox.Add_TextChanged({ Update-Factors $false })
$baseAreaBox.Add_TextChanged({ Update-Factors $false })
$exponentBox.Add_TextChanged({ Update-Factors $false })
$conversionFactor1Box.Add_TextChanged({ Update-Factors $false })
$conversionFactor2Box.Add_TextChanged({ Update-Factors $false })
$ratioBox.Add_TextChanged({ Update-Factors $false })
$laborYearBox.Add_ValueChanged({ Set-LaborRatesForYear ([int]$laborYearBox.Value) })
$outputNameBox.Add_TextChanged({
    if (-not $script:UpdatingOutputName) {
        $script:OutputNameAuto = $false
    }
    Refresh-OutputLabel
})
$outputFolderBox.Add_TextChanged({ Refresh-OutputLabel })

$watchTimer = New-Object System.Windows.Forms.Timer
$watchTimer.Interval = 3000
$watchTimer.Add_Tick({
    if ($script:PdfPaths.Count -gt 0 -and $script:HasConverted -and -not $script:Busy) {
        if (Signatures-Changed) {
            $statusLabel.Text = "PDF 변경 감지"
            Invoke-Conversion "PDF 변경 자동 변환"
        }
    }
})

Refresh-OutputLabel
Update-Factors $true
Clear-ErrorText
$watchTimer.Start()
[void]$form.ShowDialog()
