param(
    [Parameter(Mandatory = $true)][string]$SourcePptx,
    [Parameter(Mandatory = $true)][string]$OutputPptx,
    [Parameter(Mandatory = $true)][int]$SlideIndex,
    [Parameter(Mandatory = $true)][string]$SentinelText
)

$ErrorActionPreference = "Stop"
$sourcePath = (Get-Item -LiteralPath $SourcePptx -ErrorAction Stop).FullName
$outputPath = [System.IO.Path]::GetFullPath($OutputPptx)

if ([System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant() -ne ".pptx") {
    throw "Source must be a PPTX file"
}
if ([System.IO.Path]::GetExtension($outputPath).ToLowerInvariant() -ne ".pptx") {
    throw "Output must be a PPTX file"
}
if (Test-Path -LiteralPath $outputPath) {
    throw "Refusing to overwrite sentinel output: $outputPath"
}
if ($SlideIndex -lt 1) {
    throw "SlideIndex must be at least 1"
}

$outputDirectory = Split-Path -Parent $outputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
Copy-Item -LiteralPath $sourcePath -Destination $outputPath

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($outputPath, $false, $false, $false)
    if ($SlideIndex -gt $presentation.Slides.Count) {
        throw "SlideIndex exceeds slide count"
    }
    $slide = $presentation.Slides.Item($SlideIndex)
    $shape = $slide.Shapes.AddTextbox(1, 720, 495, 210, 28)
    $shape.Name = "USER_MANUAL_SENTINEL"
    $shape.TextFrame.TextRange.Text = $SentinelText
    $shape.TextFrame.TextRange.Font.Name = "Arial"
    $shape.TextFrame.TextRange.Font.Size = 8
    $shape.TextFrame.TextRange.Font.Color.RGB = 8421504
    $presentation.Save()
    Write-Output ("SENTINEL_SLIDE=" + $SlideIndex)
    Write-Output ("OUTPUT=" + $outputPath)
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
