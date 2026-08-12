param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutputJson,
    [string]$SlideIndexes = "",
    [string]$PreviewDir = "",
    [string]$OutputPdf = ""
)

$ErrorActionPreference = "Stop"
$inputPath = (Get-Item -LiteralPath $InputPptx -ErrorAction Stop).FullName
$outputPath = [System.IO.Path]::GetFullPath($OutputJson)
if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne ".pptx") {
    throw "Input must be a PPTX file"
}
if ([System.IO.Path]::GetExtension($outputPath).ToLowerInvariant() -ne ".json") {
    throw "Output must be a JSON file"
}
if (Test-Path -LiteralPath $outputPath) {
    throw "Refusing to overwrite layout manifest: $outputPath"
}
$outputDirectory = Split-Path -Parent $outputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$previewPath = ""
if (-not [string]::IsNullOrWhiteSpace($PreviewDir)) {
    $previewPath = [System.IO.Path]::GetFullPath($PreviewDir)
    if (Test-Path -LiteralPath $previewPath) {
        if (@(Get-ChildItem -LiteralPath $previewPath -Force).Count -gt 0) {
            throw "Refusing to write previews into a non-empty directory: $previewPath"
        }
    } else {
        New-Item -ItemType Directory -Path $previewPath | Out-Null
    }
}
$pdfPath = ""
if (-not [string]::IsNullOrWhiteSpace($OutputPdf)) {
    $pdfPath = [System.IO.Path]::GetFullPath($OutputPdf)
    if ([System.IO.Path]::GetExtension($pdfPath).ToLowerInvariant() -ne ".pdf") {
        throw "OutputPdf must use the .pdf extension"
    }
    if (Test-Path -LiteralPath $pdfPath) {
        throw "Refusing to overwrite PDF: $pdfPath"
    }
    $pdfParent = Split-Path -Parent $pdfPath
    if (-not (Test-Path -LiteralPath $pdfParent)) {
        New-Item -ItemType Directory -Path $pdfParent | Out-Null
    }
}

$requestedSlideIndexes = [System.Collections.Generic.HashSet[int]]::new()
if (-not [string]::IsNullOrWhiteSpace($SlideIndexes)) {
    foreach ($token in $SlideIndexes.Split(',')) {
        $trimmed = $token.Trim()
        $parsed = 0
        if (-not [int]::TryParse($trimmed, [ref]$parsed) -or $parsed -lt 1) {
            throw "SlideIndexes must be a comma-separated list of positive integers"
        }
        if (-not $requestedSlideIndexes.Add($parsed)) {
            throw "SlideIndexes contains a duplicate index: $parsed"
        }
    }
}

function Get-StringSha256([string]$Value) {
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        return (
            [System.BitConverter]::ToString($algorithm.ComputeHash($bytes))
        ).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
    }
}

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputPath, $true, $false, $false)
    $slides = @()
    foreach ($slide in @($presentation.Slides)) {
        if ($requestedSlideIndexes.Count -gt 0 -and -not $requestedSlideIndexes.Contains([int]$slide.SlideIndex)) {
            continue
        }
        $slideId = ""
        try {
            $candidateId = [string]$slide.Tags.Item("ACADEMIC_SLIDE_ID")
            if ($candidateId -match "^SLD-[0-9A-F]{20}$") {
                $slideId = $candidateId
            }
        } catch {}
        $textShapes = @()
        $allShapes = @()
        foreach ($shape in @($slide.Shapes)) {
            $shapeName = [string]$shape.Name
            if (-not $slideId -and $shapeName -match "(SLD-[0-9A-F]{20})") {
                $slideId = $Matches[1]
            }
            $shapeRole = "content"
            if ($shapeName -match "(?i)page-number") {
                $shapeRole = "page_number"
            } elseif ($shapeName -match "(?i)sources|footer") {
                $shapeRole = "footer"
            } elseif ($shapeName -match "(?i)title") {
                $shapeRole = "title"
            } elseif ($shapeName -match "(?i)chart|forest|reference-line|estimate|ci") {
                $shapeRole = "chart"
            } elseif ($shapeName -match "(?i)source-figure") {
                $shapeRole = "picture"
            }
            $groupChildCount = 0
            try {
                if ([int]$shape.Type -eq 6) {
                    $groupChildCount = [int]$shape.GroupItems.Count
                }
            } catch {}
            $hasChart = $false
            $hasTable = $false
            try { $hasChart = [bool]$shape.HasChart } catch {}
            try { $hasTable = [bool]$shape.HasTable } catch {}
            $allShapes += [PSCustomObject]@{
                object_index = [int]$shape.Id
                object_name_sha256 = Get-StringSha256 $shapeName
                object_type = [int]$shape.Type
                shape_role = $shapeRole
                left_pt = [double]$shape.Left
                top_pt = [double]$shape.Top
                width_pt = [double]$shape.Width
                height_pt = [double]$shape.Height
                rotation_degrees = [double]$shape.Rotation
                has_chart = $hasChart
                has_table = $hasTable
                group_child_count = $groupChildCount
            }
            $hasText = $false
            try {
                $hasText = [bool]($shape.HasTextFrame -and $shape.TextFrame2.HasText)
            } catch {}
            if (-not $hasText) {
                continue
            }
            $text = [string]$shape.TextFrame.TextRange.Text
            $placeholderType = 0
            try {
                if ($shape.Type -eq 14) {
                    $placeholderType = [int]$shape.PlaceholderFormat.Type
                }
            } catch {}
            $boundWidth = $null
            $boundHeight = $null
            $boundLeft = $null
            $boundTop = $null
            $overflowing = $false
            $marginLeft = $null
            $marginRight = $null
            $marginTop = $null
            $marginBottom = $null
            $autoSize = $null
            $wordWrap = $null
            $fontName = $null
            $fontSize = $null
            try {
                $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
                $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
                $boundLeft = [double]$shape.TextFrame2.TextRange.BoundLeft
                $boundTop = [double]$shape.TextFrame2.TextRange.BoundTop
            } catch {}
            try { $overflowing = [bool]$shape.TextFrame2.Overflowing } catch {}
            try { $marginLeft = [double]$shape.TextFrame2.MarginLeft } catch {}
            try { $marginRight = [double]$shape.TextFrame2.MarginRight } catch {}
            try { $marginTop = [double]$shape.TextFrame2.MarginTop } catch {}
            try { $marginBottom = [double]$shape.TextFrame2.MarginBottom } catch {}
            try { $autoSize = [int]$shape.TextFrame2.AutoSize } catch {}
            try { $wordWrap = [int]$shape.TextFrame2.WordWrap } catch {}
            try { $fontName = [string]$shape.TextFrame2.TextRange.Font.Name } catch {}
            try { $fontSize = [double]$shape.TextFrame2.TextRange.Font.Size } catch {}
            $textShapes += [PSCustomObject]@{
                object_index = [int]$shape.Id
                object_name_sha256 = Get-StringSha256 ([string]$shape.Name)
                object_type = [int]$shape.Type
                placeholder_type = $placeholderType
                left_pt = [double]$shape.Left
                top_pt = [double]$shape.Top
                width_pt = [double]$shape.Width
                height_pt = [double]$shape.Height
                rotation_degrees = [double]$shape.Rotation
                shape_role = $shapeRole
                text_length = $text.Length
                text_sha256 = Get-StringSha256 $text
                bound_left_pt = $boundLeft
                bound_top_pt = $boundTop
                bound_width_pt = $boundWidth
                bound_height_pt = $boundHeight
                overflowing = $overflowing
                margin_left_pt = $marginLeft
                margin_right_pt = $marginRight
                margin_top_pt = $marginTop
                margin_bottom_pt = $marginBottom
                auto_size = $autoSize
                word_wrap = $wordWrap
                font_name = $fontName
                font_size_pt = $fontSize
                footer_like = [bool](
                    $shapeRole -in @("footer", "page_number") -or
                    $placeholderType -in @(5, 6, 7, 8, 9, 10, 11, 12, 13) -or
                    (
                        [double]$shape.Top -ge ([double]$presentation.PageSetup.SlideHeight * 0.880) -and
                        ($null -eq $fontSize -or [double]$fontSize -le 11.5)
                    )
                )
            }
        }
        $slides += [PSCustomObject]@{
            slide_index = [int]$slide.SlideIndex
            slide_id = $slideId
            shapes = $allShapes
            text_shapes = $textShapes
        }
        if ($previewPath) {
            $previewFile = Join-Path $previewPath ("slide_{0:D3}.png" -f [int]$slide.SlideIndex)
            $slide.Export($previewFile, "PNG", 1600, 900)
            if (-not (Test-Path -LiteralPath $previewFile -PathType Leaf)) {
                throw "PowerPoint did not export changed-slide preview: $previewFile"
            }
        }
    }
    if ($pdfPath) {
        $presentation.SaveAs($pdfPath, 32)
        if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf)) {
            throw "PowerPoint did not export PDF: $pdfPath"
        }
    }
    $payload = [PSCustomObject]@{
        schema_version = "2.1"
        status = "PASS"
        source_pptx_sha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
        slide_width_pt = [double]$presentation.PageSetup.SlideWidth
        slide_height_pt = [double]$presentation.PageSetup.SlideHeight
        slide_count = [int]$presentation.Slides.Count
        inspected_slide_count = [int]$slides.Count
        requested_slide_indexes = @($requestedSlideIndexes | Sort-Object)
        privacy_contract = [PSCustomObject]@{
            raw_text_returned = $false
            shape_names_returned = $false
            text_hashes_only = $true
        }
        slides = $slides
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    Write-Output ("LAYOUT_SLIDES=" + $presentation.Slides.Count)
    Write-Output ("INSPECTED_SLIDES=" + $slides.Count)
    Write-Output ("PREVIEW_COUNT=" + ($(if ($previewPath) { @(Get-ChildItem -LiteralPath $previewPath -Filter "*.png" -File).Count } else { 0 })))
    Write-Output ("PDF_EXPORTED=" + [bool]$pdfPath)
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
