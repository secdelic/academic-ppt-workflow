param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [Parameter(Mandatory = $true)][string]$PreviewDir
)

$ErrorActionPreference = "Stop"
$inputResolved = [System.IO.Path]::GetFullPath($InputPptx)
$pdfResolved = [System.IO.Path]::GetFullPath($OutputPdf)
$previewResolved = [System.IO.Path]::GetFullPath($PreviewDir)

if (-not (Test-Path -LiteralPath $inputResolved -PathType Leaf)) {
    throw "Input PPTX not found: $inputResolved"
}
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($pdfResolved)) | Out-Null
New-Item -ItemType Directory -Force -Path $previewResolved | Out-Null

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputResolved, $true, $true, $false)
    $presentation.SaveAs($pdfResolved, 32)
    $presentation.Export($previewResolved, "PNG", 1600, 900)
    if (-not (Test-Path -LiteralPath $pdfResolved -PathType Leaf)) {
        throw "PowerPoint returned without producing a PDF"
    }
    $pngCount = @(Get-ChildItem -LiteralPath $previewResolved -Filter "*.PNG" -File).Count
    if ($pngCount -ne $presentation.Slides.Count) {
        throw "Preview count mismatch: expected $($presentation.Slides.Count), found $pngCount"
    }
    "RENDERER=POWERPOINT_COM"
    "SLIDE_COUNT=$($presentation.Slides.Count)"
    "PDF=$pdfResolved"
    "PREVIEW_DIR=$previewResolved"
}
finally {
    if ($presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($powerPoint) {
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
