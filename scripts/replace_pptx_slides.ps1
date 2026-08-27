param(
    [Parameter(Mandatory = $true)][string]$ExistingCopy,
    [Parameter(Mandatory = $true)][string]$Candidate,
    [Parameter(Mandatory = $true)][string]$Mapping
)

$ErrorActionPreference = "Stop"

function Get-AbsoluteFile([string]$PathValue, [string]$Label) {
    $item = Get-Item -LiteralPath $PathValue -ErrorAction Stop
    if ($item.PSIsContainer) {
        throw "$Label must be a file: $PathValue"
    }
    return $item.FullName
}

$existingPath = Get-AbsoluteFile $ExistingCopy "Existing copy"
$candidatePath = Get-AbsoluteFile $Candidate "Candidate"
$mappingPath = Get-AbsoluteFile $Mapping "Mapping"
if ([System.IO.Path]::GetExtension($existingPath).ToLowerInvariant() -ne ".pptx") {
    throw "Existing copy must be a .pptx file"
}
if ([System.IO.Path]::GetExtension($candidatePath).ToLowerInvariant() -ne ".pptx") {
    throw "Candidate must be a .pptx file"
}
if ([System.IO.Path]::GetExtension($mappingPath).ToLowerInvariant() -ne ".json") {
    throw "Mapping must be a .json file"
}
if ([string]::Equals($existingPath, $candidatePath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Existing copy and candidate must be different files"
}
$candidateHashBefore = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash
$mappingHashBefore = (Get-FileHash -LiteralPath $mappingPath -Algorithm SHA256).Hash
$mappingData = Get-Content -LiteralPath $mappingPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $mappingData.targets -or @($mappingData.targets).Count -eq 0) {
    throw "Mapping contains no targets"
}
$seenSlideIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$seenTargetIndexes = [System.Collections.Generic.HashSet[int]]::new()
$seenCandidateIndexes = [System.Collections.Generic.HashSet[int]]::new()
foreach ($target in @($mappingData.targets)) {
    $slideId = [string]$target.slide_id
    $targetIndex = [int]$target.target_index
    $candidateIndex = [int]$target.candidate_index
    $revision = [int]$target.slide_revision
    $contentHash = [string]$target.content_hash
    if ([string]::IsNullOrWhiteSpace($slideId) -or $slideId.Length -gt 120 -or $slideId -match "[\x00-\x1F]") {
        throw "Mapping contains an invalid slide_id"
    }
    if (-not $seenSlideIds.Add($slideId)) {
        throw "Mapping contains duplicate slide_id: $slideId"
    }
    if ($targetIndex -lt 1 -or -not $seenTargetIndexes.Add($targetIndex)) {
        throw "Mapping contains an invalid or duplicate target_index: $targetIndex"
    }
    if ($candidateIndex -lt 1 -or -not $seenCandidateIndexes.Add($candidateIndex)) {
        throw "Mapping contains an invalid or duplicate candidate_index: $candidateIndex"
    }
    if ($revision -lt 1) {
        throw "Mapping contains an invalid slide_revision for $slideId"
    }
    if ($contentHash -notmatch "^[0-9a-fA-F]{64}$") {
        throw "Mapping contains an invalid content_hash for $slideId"
    }
}

$powerPoint = $null
$targetPresentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $targetPresentation = $powerPoint.Presentations.Open($existingPath, $false, $false, $false)

    foreach ($target in @($mappingData.targets)) {
        $targetIndex = [int]$target.target_index
        $candidateIndex = [int]$target.candidate_index
        if ($targetIndex -lt 1 -or $targetIndex -gt $targetPresentation.Slides.Count) {
            throw "Target index out of range: $targetIndex"
        }

        # InsertFromFile avoids the Windows clipboard, which is unreliable in
        # non-interactive COM sessions.  The candidate stays read-only on disk.
        $targetPresentation.Slides.Item($targetIndex).Delete()
        $insertAfter = [Math]::Max(0, $targetIndex - 1)
        $insertedCount = $targetPresentation.Slides.InsertFromFile(
            $candidatePath,
            $insertAfter,
            $candidateIndex,
            $candidateIndex
        )
        if ([int]$insertedCount -ne 1) {
            throw "InsertFromFile did not insert exactly one slide for $($target.slide_id)"
        }
        $insertedSlide = $targetPresentation.Slides.Item($targetIndex)
        $insertedSlide.Name = "AWF_" + [string]$target.slide_id
        $insertedSlide.Tags.Add("ACADEMIC_SLIDE_ID", [string]$target.slide_id)
        $insertedSlide.Tags.Add("ACADEMIC_SLIDE_REVISION", [string]$target.slide_revision)
        $insertedSlide.Tags.Add("ACADEMIC_CONTENT_HASH", [string]$target.content_hash)
    }

    $targetPresentation.Save()
    $candidateHashAfter = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash
    $mappingHashAfter = (Get-FileHash -LiteralPath $mappingPath -Algorithm SHA256).Hash
    if ($candidateHashBefore -ne $candidateHashAfter) {
        throw "Candidate PPTX changed during read-only replacement"
    }
    if ($mappingHashBefore -ne $mappingHashAfter) {
        throw "Mapping JSON changed during replacement"
    }
    Write-Output ("UPDATED_SLIDES=" + @($mappingData.targets).Count)
    Write-Output ("OUTPUT=" + $existingPath)
}
finally {
    if ($targetPresentation -ne $null) {
        $targetPresentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetPresentation)
    }
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
