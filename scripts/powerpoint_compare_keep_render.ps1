param(
    [Parameter(Mandatory = $true)][string]$SourcePptx,
    [Parameter(Mandatory = $true)][string]$UpdatedPptx,
    [Parameter(Mandatory = $true)][string]$PairsJson,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

$ErrorActionPreference = "Stop"

function Resolve-ExistingFile([string]$Value, [string]$Label) {
    $resolved = [System.IO.Path]::GetFullPath($Value)
    if (-not [System.IO.File]::Exists($resolved)) {
        throw "$Label does not exist"
    }
    return $resolved
}

$source = Resolve-ExistingFile $SourcePptx "Source PPTX"
$updated = Resolve-ExistingFile $UpdatedPptx "Updated PPTX"
$pairsPath = Resolve-ExistingFile $PairsJson "Pair manifest"
$destination = [System.IO.Path]::GetFullPath($OutputDir)
if ([System.IO.Directory]::Exists($destination)) {
    throw "KEEP render output directory must be new"
}
[System.IO.Directory]::CreateDirectory($destination) | Out-Null
$pairs = Get-Content -LiteralPath $pairsPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($pairs.Count -lt 1) {
    throw "KEEP render pair manifest is empty"
}

$powerPoint = $null
$sourceDeck = $null
$updatedDeck = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $sourceDeck = $powerPoint.Presentations.Open($source, $true, $false, $false)
    $updatedDeck = $powerPoint.Presentations.Open($updated, $true, $false, $false)
    foreach ($pair in $pairs) {
        $ordinal = [int]$pair.ordinal
        $sourceIndex = [int]$pair.source_index
        $updatedIndex = [int]$pair.updated_index
        if ($sourceIndex -lt 1 -or $sourceIndex -gt $sourceDeck.Slides.Count) {
            throw "Invalid source slide index"
        }
        if ($updatedIndex -lt 1 -or $updatedIndex -gt $updatedDeck.Slides.Count) {
            throw "Invalid updated slide index"
        }
        $sourcePng = Join-Path $destination ("keep_{0:D4}_before.png" -f $ordinal)
        $updatedPng = Join-Path $destination ("keep_{0:D4}_after.png" -f $ordinal)
        $sourceDeck.Slides.Item($sourceIndex).Export($sourcePng, "PNG", 1600, 900)
        $updatedDeck.Slides.Item($updatedIndex).Export($updatedPng, "PNG", 1600, 900)
    }
    Write-Output "KEEP_RENDER_EXPORT_PASS"
}
finally {
    if ($sourceDeck -ne $null) { $sourceDeck.Close() }
    if ($updatedDeck -ne $null) { $updatedDeck.Close() }
    if ($powerPoint -ne $null) { $powerPoint.Quit() }
    if ($sourceDeck -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceDeck) }
    if ($updatedDeck -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($updatedDeck) }
    if ($powerPoint -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
