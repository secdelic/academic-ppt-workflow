param(
    [Parameter(Mandatory = $true)][string]$SourceDeck,
    [Parameter(Mandatory = $true)][string]$OperationPlan,
    [Parameter(Mandatory = $true)][string]$OutputPptx,
    [Parameter(Mandatory = $false)][string]$PreviewDir,
    [Parameter(Mandatory = $false)][string]$ExportPdf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-LocalFile([string]$PathValue, [string]$Label, [string]$Extension) {
    $item = Get-Item -LiteralPath $PathValue -ErrorAction Stop
    if ($item.PSIsContainer) {
        throw "$Label must be a file: $PathValue"
    }
    if ($item.FullName.StartsWith("\\", [System.StringComparison]::Ordinal)) {
        throw "$Label must be a local file, not a UNC path"
    }
    if ([System.IO.Path]::GetExtension($item.FullName).ToLowerInvariant() -ne $Extension) {
        throw "$Label must use the $Extension extension"
    }
    return $item.FullName
}

function Get-LocalOutputPath([string]$PathValue) {
    $fullPath = [System.IO.Path]::GetFullPath($PathValue)
    if ($fullPath.StartsWith("\\", [System.StringComparison]::Ordinal)) {
        throw "Output PPTX must use a local path"
    }
    if ([System.IO.Path]::GetExtension($fullPath).ToLowerInvariant() -ne ".pptx") {
        throw "Output must be a .pptx file"
    }
    if (Test-Path -LiteralPath $fullPath) {
        throw "Refusing to overwrite output: $fullPath"
    }
    return $fullPath
}

function Get-NewLocalDirectoryPath([string]$PathValue, [string]$Label) {
    $fullPath = [System.IO.Path]::GetFullPath($PathValue)
    if ($fullPath.StartsWith("\\", [System.StringComparison]::Ordinal)) {
        throw "$Label must use a local path"
    }
    if (Test-Path -LiteralPath $fullPath) {
        throw "Refusing to reuse $Label path: $fullPath"
    }
    return $fullPath
}

function Get-NewPdfPath([string]$PathValue) {
    $fullPath = [System.IO.Path]::GetFullPath($PathValue)
    if ($fullPath.StartsWith("\\", [System.StringComparison]::Ordinal)) {
        throw "PDF output must use a local path"
    }
    if ([System.IO.Path]::GetExtension($fullPath).ToLowerInvariant() -ne ".pdf") {
        throw "PDF output must use the .pdf extension"
    }
    if (Test-Path -LiteralPath $fullPath) {
        throw "Refusing to overwrite PDF output: $fullPath"
    }
    return $fullPath
}

function Assert-SlideId([string]$SlideId, [string]$Label) {
    if (
        [string]::IsNullOrWhiteSpace($SlideId) -or
        $SlideId.Length -gt 120 -or
        $SlideId -match "[\x00-\x1F\x7F]"
    ) {
        throw "$Label contains an invalid semantic slide_id"
    }
}

function Get-SequenceHash([string[]]$SlideIds) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($SlideIds -join "`n"))
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-Candidate([hashtable]$Candidates, [string]$CandidateId) {
    if (-not $Candidates.ContainsKey($CandidateId)) {
        throw "Unknown candidate_id: $CandidateId"
    }
    return $Candidates[$CandidateId]
}

function Get-OptionalProperty($ObjectValue, [string]$Name) {
    $property = $ObjectValue.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Set-ChangedSlideIdentity($Slide, $Operation) {
    try {
        $slideId = [string]$Operation.slide_id
        $Slide.Name = "AWF_" + $slideId
        $Slide.Tags.Add("ACADEMIC_SLIDE_ID", $slideId)
        $revision = Get-OptionalProperty $Operation "slide_revision"
        if ($null -ne $revision -and [string]$revision) {
            $Slide.Tags.Add("ACADEMIC_SLIDE_REVISION", [string]$revision)
        }
        $contentHash = Get-OptionalProperty $Operation "content_hash"
        if ($null -ne $contentHash -and [string]$contentHash) {
            $Slide.Tags.Add("ACADEMIC_CONTENT_HASH", [string]$contentHash)
        }
    }
    catch {
        throw "Could not apply stable identity to changed slide $($Operation.slide_id): $($_.Exception.Message)"
    }
}

$sourcePath = Resolve-LocalFile $SourceDeck "Source deck" ".pptx"
$planPath = Resolve-LocalFile $OperationPlan "Operation plan" ".json"
$outputPath = Get-LocalOutputPath $OutputPptx
$previewPath = if ([string]::IsNullOrWhiteSpace($PreviewDir)) { $null } else { Get-NewLocalDirectoryPath $PreviewDir "preview directory" }
$pdfPath = if ([string]::IsNullOrWhiteSpace($ExportPdf)) { $null } else { Get-NewPdfPath $ExportPdf }
if (
    [string]::Equals($sourcePath, $planPath, [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($sourcePath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($planPath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)
) {
    throw "Source, plan, and output paths must be distinct"
}
if ($null -ne $pdfPath -and (
    [string]::Equals($pdfPath, $sourcePath, [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($pdfPath, $planPath, [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($pdfPath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)
)) {
    throw "PDF output path must be distinct from every PPTX or plan path"
}

$sourceHashBefore = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
$planHashBefore = (Get-FileHash -LiteralPath $planPath -Algorithm SHA256).Hash.ToLowerInvariant()
$plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$plan.schema_version -ne "1.0" -or [string]$plan.plan_type -ne "fast_enhance_operations") {
    throw "Unsupported operation plan schema or plan_type"
}
if (-not $plan.source_deck -or -not $plan.operations -or -not $plan.expected_final_order) {
    throw "Operation plan is missing required objects"
}
$plannedSourcePath = [System.IO.Path]::GetFullPath([string]$plan.source_deck.path)
if (-not [string]::Equals($sourcePath, $plannedSourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Source deck argument does not match operation plan"
}
if ($sourceHashBefore -ne ([string]$plan.source_deck.sha256).ToLowerInvariant()) {
    throw "Source deck SHA-256 does not match operation plan"
}

$sourceIds = @($plan.source_slide_ids | ForEach-Object { [string]$_ })
$expectedIds = @($plan.expected_final_order | ForEach-Object { [string]$_ })
$operations = @($plan.operations)
if ($sourceIds.Count -ne [int]$plan.source_deck.slide_count) {
    throw "source_slide_ids count does not match source_deck.slide_count"
}
if ($expectedIds.Count -ne [int]$plan.expected_final_slide_count) {
    throw "expected_final_order count does not match expected_final_slide_count"
}
if ($operations.Count -ne $expectedIds.Count) {
    throw "Every final slide requires exactly one operation"
}
$sourceSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($slideId in $sourceIds) {
    Assert-SlideId $slideId "source_slide_ids"
    if (-not $sourceSet.Add($slideId)) {
        throw "Duplicate source slide_id: $slideId"
    }
}
$expectedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($slideId in $expectedIds) {
    Assert-SlideId $slideId "expected_final_order"
    if (-not $expectedSet.Add($slideId)) {
        throw "Duplicate expected slide_id: $slideId"
    }
}
$expectedOrderHash = Get-SequenceHash $expectedIds
if ($expectedOrderHash -ne ([string]$plan.expected_final_order_sha256).ToLowerInvariant()) {
    throw "expected_final_order SHA-256 does not match operation plan"
}

$candidateById = @{}
$protectedCandidateHashes = @{}
foreach ($candidate in @($plan.candidate_decks)) {
    $candidateId = [string]$candidate.candidate_id
    if ([string]::IsNullOrWhiteSpace($candidateId) -or $candidateById.ContainsKey($candidateId)) {
        throw "Invalid or duplicate candidate_id"
    }
    $candidatePath = Resolve-LocalFile ([string]$candidate.path) "Candidate deck" ".pptx"
    if (
        [string]::Equals($candidatePath, $sourcePath, [System.StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($candidatePath, $planPath, [System.StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($candidatePath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        throw "Candidate path must be distinct from source, plan, and output"
    }
    $candidateHash = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($candidateHash -ne ([string]$candidate.sha256).ToLowerInvariant()) {
        throw "Candidate deck SHA-256 does not match operation plan: $candidateId"
    }
    $candidateById[$candidateId] = [pscustomobject]@{
        path = $candidatePath
        slide_count = [int]$candidate.slide_count
    }
    $protectedCandidateHashes[$candidatePath] = $candidateHash
}

$operationById = @{}
foreach ($operation in $operations) {
    $slideId = [string]$operation.slide_id
    Assert-SlideId $slideId "operation"
    if ($operationById.ContainsKey($slideId)) {
        throw "Duplicate operation slide_id: $slideId"
    }
    $action = ([string]$operation.action).ToUpperInvariant()
    if ($action -notin @("KEEP", "REPLACE", "INSERT_AFTER")) {
        throw "Unsupported action for $slideId"
    }
    if ($action -in @("KEEP", "REPLACE") -and -not $sourceSet.Contains($slideId)) {
        throw "$action slide_id is not in source_slide_ids: $slideId"
    }
    if ($action -eq "INSERT_AFTER" -and $sourceSet.Contains($slideId)) {
        throw "INSERT_AFTER must introduce a new slide_id: $slideId"
    }
    if ($action -in @("REPLACE", "INSERT_AFTER")) {
        $candidate = Get-Candidate $candidateById ([string]$operation.candidate_id)
        $candidateIndex = [int]$operation.candidate_index
        if ($candidateIndex -lt 1 -or $candidateIndex -gt $candidate.slide_count) {
            throw "candidate_index out of range for $slideId"
        }
    }
    $contentHash = Get-OptionalProperty $operation "content_hash"
    if ($null -ne $contentHash -and [string]$contentHash) {
        if ([string]$contentHash -notmatch "^[0-9a-fA-F]{64}$") {
            throw "Invalid content_hash for $slideId"
        }
    }
    $revision = Get-OptionalProperty $operation "slide_revision"
    if ($null -ne $revision -and [string]$revision) {
        if ([int]$revision -lt 1) {
            throw "Invalid slide_revision for $slideId"
        }
    }
    $operationById[$slideId] = $operation
}
for ($index = 0; $index -lt $expectedIds.Count; $index += 1) {
    $slideId = $expectedIds[$index]
    if (-not $operationById.ContainsKey($slideId)) {
        throw "No operation for expected slide_id: $slideId"
    }
    $operation = $operationById[$slideId]
    if ([string]$operation.action -eq "INSERT_AFTER") {
        if ($index -eq 0 -or [string]$operation.after_slide_id -ne $expectedIds[$index - 1]) {
            throw "INSERT_AFTER must name the immediate expected-order predecessor for $slideId"
        }
    }
}

$outputDirectory = Split-Path -Parent $outputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
Copy-Item -LiteralPath $sourcePath -Destination $outputPath
if ((Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sourceHashBefore) {
    throw "Working-copy hash mismatch before COM application"
}

$powerPoint = $null
$presentation = $null
$appliedOperations = 0
try {
    # Exactly one PowerPoint application owns the entire incremental operation.
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($outputPath, $false, $false, $false)
    if ($presentation.Slides.Count -ne $sourceIds.Count) {
        throw "PowerPoint source slide count does not match operation plan"
    }

    $currentIds = [System.Collections.Generic.List[string]]::new()
    foreach ($slideId in $sourceIds) {
        $currentIds.Add($slideId)
    }

    # Replacement is count-neutral, so it can be applied before insertions.
    foreach ($operation in $operations) {
        if ([string]$operation.action -ne "REPLACE") { continue }
        $slideId = [string]$operation.slide_id
        $zeroBasedIndex = $currentIds.IndexOf($slideId)
        if ($zeroBasedIndex -lt 0) {
            throw "REPLACE target not found: $slideId"
        }
        $oneBasedIndex = $zeroBasedIndex + 1
        $candidate = Get-Candidate $candidateById ([string]$operation.candidate_id)
        $inserted = $presentation.Slides.InsertFromFile(
            $candidate.path,
            $oneBasedIndex,
            [int]$operation.candidate_index,
            [int]$operation.candidate_index
        )
        if ([int]$inserted -ne 1) {
            throw "InsertFromFile did not insert exactly one replacement for $slideId"
        }
        # The inserted slide follows the old slide; delete the old object so the
        # replacement occupies the same semantic position.
        $presentation.Slides.Item($oneBasedIndex).Delete()
        Set-ChangedSlideIdentity $presentation.Slides.Item($oneBasedIndex) $operation
        $appliedOperations += 1
    }

    # Insert new slides in expected order.  The immediate-predecessor contract
    # prevents order reversal when several slides are inserted near one anchor.
    foreach ($slideId in $expectedIds) {
        if ($currentIds.Contains($slideId)) { continue }
        $operation = $operationById[$slideId]
        if ([string]$operation.action -ne "INSERT_AFTER") {
            throw "Expected new slide does not use INSERT_AFTER: $slideId"
        }
        $anchorId = [string]$operation.after_slide_id
        $anchorZeroBased = $currentIds.IndexOf($anchorId)
        if ($anchorZeroBased -lt 0) {
            throw "INSERT_AFTER anchor not found: $anchorId"
        }
        $candidate = Get-Candidate $candidateById ([string]$operation.candidate_id)
        $inserted = $presentation.Slides.InsertFromFile(
            $candidate.path,
            $anchorZeroBased + 1,
            [int]$operation.candidate_index,
            [int]$operation.candidate_index
        )
        if ([int]$inserted -ne 1) {
            throw "InsertFromFile did not insert exactly one slide for $slideId"
        }
        $newZeroBased = $anchorZeroBased + 1
        $currentIds.Insert($newZeroBased, $slideId)
        Set-ChangedSlideIdentity $presentation.Slides.Item($newZeroBased + 1) $operation
        $appliedOperations += 1
    }

    if ($presentation.Slides.Count -ne $expectedIds.Count -or $currentIds.Count -ne $expectedIds.Count) {
        throw "Final slide count does not match expected_final_order"
    }
    for ($index = 0; $index -lt $expectedIds.Count; $index += 1) {
        if ($currentIds[$index] -ne $expectedIds[$index]) {
            throw "Final semantic slide order mismatch at position $($index + 1)"
        }
    }
    $presentation.Save()
    $changedPreviewCount = 0
    if ($null -ne $previewPath) {
        New-Item -ItemType Directory -Path $previewPath | Out-Null
        $changedIds = @($plan.changed_slide_ids | ForEach-Object { [string]$_ })
        foreach ($slideId in $changedIds) {
            $zeroBasedIndex = $currentIds.IndexOf($slideId)
            if ($zeroBasedIndex -lt 0) {
                throw "Changed slide not found for preview export: $slideId"
            }
            $previewFile = Join-Path $previewPath ("slide_{0:D3}.png" -f ($zeroBasedIndex + 1))
            $presentation.Slides.Item($zeroBasedIndex + 1).Export($previewFile, "PNG", 1600, 900)
            if (-not (Test-Path -LiteralPath $previewFile)) {
                throw "Changed-slide preview was not generated: $slideId"
            }
            $changedPreviewCount += 1
        }
    }
    if ($null -ne $pdfPath) {
        $pdfDirectory = Split-Path -Parent $pdfPath
        if (-not (Test-Path -LiteralPath $pdfDirectory)) {
            New-Item -ItemType Directory -Path $pdfDirectory | Out-Null
        }
        $presentation.SaveAs($pdfPath, 32)
        if (-not (Test-Path -LiteralPath $pdfPath)) {
            throw "PowerPoint did not generate the requested PDF"
        }
        Write-Output "PDF_EXPORT=PASS"
    }
    Write-Output "POWERPOINT_OPEN=PASS"
    Write-Output ("APPLIED_OPERATIONS=" + $appliedOperations)
    Write-Output ("FINAL_SLIDES=" + $presentation.Slides.Count)
    Write-Output ("FINAL_ORDER_SHA256=" + (Get-SequenceHash $currentIds.ToArray()))
    Write-Output ("CHANGED_PREVIEW_COUNT=" + $changedPreviewCount)
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

$sourceHashAfter = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
$planHashAfter = (Get-FileHash -LiteralPath $planPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($sourceHashAfter -ne $sourceHashBefore) {
    throw "Source deck changed during incremental operation"
}
if ($planHashAfter -ne $planHashBefore) {
    throw "Operation plan changed during incremental operation"
}
foreach ($candidatePath in $protectedCandidateHashes.Keys) {
    $after = (Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($after -ne $protectedCandidateHashes[$candidatePath]) {
        throw "Candidate deck changed during incremental operation: $candidatePath"
    }
}
