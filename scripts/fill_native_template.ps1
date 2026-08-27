param(
    [Parameter(Mandatory = $true)][string]$Template,
    [Parameter(Mandatory = $true)][string]$DeckIr,
    [Parameter(Mandatory = $true)][string]$OutputPptx
)

$ErrorActionPreference = "Stop"

function Get-AbsoluteFile([string]$PathValue, [string]$Label) {
    $item = Get-Item -LiteralPath $PathValue -ErrorAction Stop
    if ($item.PSIsContainer) {
        throw "$Label must be a file: $PathValue"
    }
    return $item.FullName
}

$templatePath = Get-AbsoluteFile $Template "Template"
$deckIrPath = Get-AbsoluteFile $DeckIr "Deck IR"
$outputPath = [System.IO.Path]::GetFullPath($OutputPptx)
if ([System.IO.Path]::GetExtension($templatePath).ToLowerInvariant() -notin @(".pptx", ".potx")) {
    throw "Template must be a .pptx or .potx file"
}
if ([System.IO.Path]::GetExtension($deckIrPath).ToLowerInvariant() -ne ".json") {
    throw "Deck IR must be a .json file"
}
if ([System.IO.Path]::GetExtension($outputPath).ToLowerInvariant() -ne ".pptx") {
    throw "Output must be a .pptx file"
}
if (
    [string]::Equals($outputPath, $templatePath, [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($outputPath, $deckIrPath, [System.StringComparison]::OrdinalIgnoreCase)
) {
    throw "Output path must be distinct from every input"
}
if (Test-Path -LiteralPath $outputPath) {
    throw "Refusing to overwrite output: $outputPath"
}
$outputDirectory = Split-Path -Parent $outputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$deck = Get-Content -LiteralPath $deckIrPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $deck.slides -or @($deck.slides).Count -eq 0) {
    throw "Deck IR contains no slides"
}
$seenSlideIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($item in @($deck.slides)) {
    $slideId = [string]$item.slide_id
    $revision = [int]$item.slide_revision
    $contentHash = [string]$item.content_hash
    if ([string]::IsNullOrWhiteSpace($slideId) -or $slideId.Length -gt 120 -or $slideId -match "[\x00-\x1F]") {
        throw "Deck IR contains an invalid slide_id"
    }
    if (-not $seenSlideIds.Add($slideId)) {
        throw "Deck IR contains duplicate slide_id: $slideId"
    }
    if ($revision -lt 1) {
        throw "Deck IR contains an invalid slide_revision for $slideId"
    }
    if ($contentHash -notmatch "^[0-9a-fA-F]{64}$") {
        throw "Deck IR contains an invalid content_hash for $slideId"
    }
}

$templateHashBefore = (Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash
$deckIrHashBefore = (Get-FileHash -LiteralPath $deckIrPath -Algorithm SHA256).Hash

function Has-PlaceholderType($layout, [int[]]$types) {
    foreach ($shape in @($layout.Shapes)) {
        try {
            if ($shape.Type -eq 14 -and $types -contains [int]$shape.PlaceholderFormat.Type) {
                return $true
            }
        } catch {}
    }
    return $false
}

function Select-Layout($layouts, [string]$role, [string]$explicitLayoutId) {
    if ($explicitLayoutId) {
        foreach ($layout in @($layouts)) {
            if ([string]$layout.Index -eq $explicitLayoutId -or $layout.Name -eq $explicitLayoutId) {
                return $layout
            }
        }
        throw "Requested template_layout_id not found: $explicitLayoutId"
    }
    $isCover = $role -in @("cover", "title", "section")
    foreach ($layout in @($layouts)) {
        $hasTitle = Has-PlaceholderType $layout @(1, 3)
        $hasBody = Has-PlaceholderType $layout @(2, 4, 7, 14)
        if ($isCover -and $hasTitle -and (Has-PlaceholderType $layout @(4))) {
            return $layout
        }
        if (-not $isCover -and $hasTitle -and $hasBody) {
            return $layout
        }
    }
    foreach ($layout in @($layouts)) {
        if (Has-PlaceholderType $layout @(1, 3)) {
            return $layout
        }
    }
    return $layouts.Item(1)
}

function Set-Notes($slide, [string]$notesText) {
    try {
        foreach ($shape in @($slide.NotesPage.Shapes)) {
            if (
                $shape.Type -eq 14 -and
                [int]$shape.PlaceholderFormat.Type -eq 2 -and
                $shape.HasTextFrame
            ) {
                $shape.TextFrame.TextRange.Text = $notesText
                return
            }
        }
        $notesShape = $slide.NotesPage.Shapes.AddTextbox(1, 72, 360, 576, 180)
        $notesShape.TextFrame.TextRange.Text = $notesText
    } catch {
        Write-Warning ("Could not set notes for " + $slide.SlideIndex + ": " + $_.Exception.Message)
    }
}

function Test-ReferenceContentShape($shape) {
    try {
        # Placeholders define the layout contract and are populated only with
        # new Deck IR content.  They are not reference-slide content.
        if ($shape.Type -eq 14) {
            return $false
        }
        # Raster/linked pictures are not copied by default, even when they look
        # like logos or decoration.  Authorization cannot be inferred.
        if ($shape.Type -in @(11, 13)) {
            return $true
        }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            if (-not [string]::IsNullOrWhiteSpace([string]$shape.TextFrame.TextRange.Text)) {
                return $true
            }
        }
        # A grouped object can conceal text or pictures.  Inspect its children
        # recursively and remove the whole group if any child is content-bearing.
        if ($shape.Type -eq 6) {
            foreach ($child in @($shape.GroupItems)) {
                if (Test-ReferenceContentShape $child) {
                    return $true
                }
            }
        }
    } catch {
        # Unknown, unreadable non-placeholder objects fail closed.
        return $true
    }
    return $false
}

function Remove-ReferenceContentShapes($shapeCollection) {
    $removed = 0
    for ($index = $shapeCollection.Count; $index -ge 1; $index -= 1) {
        $shape = $shapeCollection.Item($index)
        if (Test-ReferenceContentShape $shape) {
            $shape.Delete()
            $removed += 1
        }
    }
    return $removed
}

$powerPoint = $null
$source = $null
$presentation = $null
$activeLayouts = $null
$templateDesignCount = 1
$referenceContentShapesRemoved = 0
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = 1
    $presentation = $powerPoint.Presentations.Add()

    if ([System.IO.Path]::GetExtension($templatePath).ToLowerInvariant() -eq ".pptx") {
        $source = $powerPoint.Presentations.Open($templatePath, $true, $false, $false)
        if ($source.Designs.Count -lt 1) {
            throw "Template PPTX contains no design"
        }
        $templateDesignCount = $source.Designs.Count
        # Clone only the design (master + layouts), never source slides.
        # This prevents private slide text/data from entering the new deck.
        $clonedDesign = $presentation.Designs.Clone($source.Designs.Item(1))
        $activeLayouts = $clonedDesign.SlideMaster.CustomLayouts
    } else {
        $presentation.ApplyTemplate($templatePath)
        $activeLayouts = $presentation.SlideMaster.CustomLayouts
    }
    if ($activeLayouts.Count -lt 1) {
        throw "Template contains no usable custom layouts"
    }
    $activeMasterShapes = $activeLayouts.Item(1).Design.SlideMaster.Shapes
    $referenceContentShapesRemoved += (
        Remove-ReferenceContentShapes $activeMasterShapes
    )
    foreach ($layoutForSanitization in @($activeLayouts)) {
        $referenceContentShapesRemoved += (
            Remove-ReferenceContentShapes $layoutForSanitization.Shapes
        )
    }

    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    foreach ($item in @($deck.slides)) {
        $role = [string]$item.slide_role
        $layout = Select-Layout $activeLayouts $role ([string]$item.template_layout_id)
        $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layout)
        $titleSet = $false
        $bodySet = $false
        $titleShapeUsed = $null
        $bodyShapeUsed = $null
        foreach ($shape in @($slide.Shapes)) {
            try {
                if ($shape.Type -ne 14) { continue }
                $placeholderType = [int]$shape.PlaceholderFormat.Type
                if ($placeholderType -in @(1, 3) -and -not $titleSet) {
                    $shape.TextFrame.TextRange.Text = [string]$item.slide_title
                    $shape.Name = "awf:" + [string]$item.slide_id + ":title"
                    $titleShapeUsed = $shape
                    $titleSet = $true
                } elseif ($placeholderType -in @(2, 4, 7, 14) -and -not $bodySet) {
                    $shape.TextFrame.TextRange.Text = [string]$item.key_message
                    $shape.Name = "awf:" + [string]$item.slide_id + ":body"
                    $bodyShapeUsed = $shape
                    $bodySet = $true
                }
            } catch {}
        }
        if (-not $titleSet) {
            $titleShape = $slide.Shapes.AddTextbox(
                1,
                $slideWidth * 0.05625,
                $slideHeight * 0.06667,
                $slideWidth * 0.8875,
                $slideHeight * 0.13333
            )
            $titleShape.TextFrame.TextRange.Text = [string]$item.slide_title
            $titleShape.TextFrame.TextRange.Font.Size = 30
            $titleShape.TextFrame.TextRange.Font.Bold = $true
            $titleShape.Name = "awf:" + [string]$item.slide_id + ":title"
            $titleShapeUsed = $titleShape
        }
        if (-not $bodySet) {
            $bodyShape = $slide.Shapes.AddTextbox(
                1,
                $slideWidth * 0.075,
                $slideHeight * 0.27778,
                $slideWidth * 0.85,
                $slideHeight * 0.5
            )
            $bodyShape.TextFrame.TextRange.Text = [string]$item.key_message
            $bodyShape.TextFrame.TextRange.Font.Size = 22
            $bodyShape.Name = "awf:" + [string]$item.slide_id + ":body"
            $bodyShapeUsed = $bodyShape
        }

        # Preserve the inherited theme, alignment, and placeholder semantics,
        # but enforce a collision-safe academic grid and readable type.  The
        # reference fixture deliberately includes a cramped layout; relying on
        # inherited geometry alone can place title and body text on top of each
        # other even though the placeholder boxes themselves do not overlap.
        $titleShapeUsed.Left = $slideWidth * 0.05625
        $titleShapeUsed.Top = $slideHeight * 0.07407
        $titleShapeUsed.Width = $slideWidth * 0.8875
        $titleShapeUsed.Height = $slideHeight * 0.22222
        $titleShapeUsed.TextFrame.TextRange.Font.Size = 30
        $titleShapeUsed.TextFrame.TextRange.Font.Bold = $true
        try {
            $titleShapeUsed.TextFrame2.AutoSize = 2
            $titleShapeUsed.TextFrame2.VerticalAnchor = 3
        } catch {}

        $bodyShapeUsed.Left = $slideWidth * 0.075
        $bodyShapeUsed.Top = $slideHeight * 0.35185
        $bodyShapeUsed.Width = $slideWidth * 0.85
        $bodyShapeUsed.Height = $slideHeight * 0.40741
        $bodyShapeUsed.TextFrame.TextRange.Font.Size = 20
        try {
            $bodyShapeUsed.TextFrame2.AutoSize = 2
            $bodyShapeUsed.TextFrame2.VerticalAnchor = 3
        } catch {}

        $sources = @($item.source_ids | ForEach-Object { [string]$_ })
        $sourceText = if ($sources.Count -gt 0) {
            "Sources: " + ($sources -join ", ")
        } else {
            "Source: INFORMATION_REQUIRED"
        }
        $footer = $slide.Shapes.AddTextbox(
            1,
            $slideWidth * 0.05625,
            $slideHeight * 0.92593,
            $slideWidth * 0.8875,
            $slideHeight * 0.05185
        )
        $footer.TextFrame.TextRange.Text = $sourceText
        $footer.TextFrame.TextRange.Font.Size = 11
        try {
            $footer.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13
        } catch {}
        $footer.Name = "awf:" + [string]$item.slide_id + ":sources"

        $slide.Name = "AWF_" + [string]$item.slide_id
        $slide.Tags.Add("ACADEMIC_SLIDE_ID", [string]$item.slide_id)
        $slide.Tags.Add("ACADEMIC_SLIDE_REVISION", [string]$item.slide_revision)
        $slide.Tags.Add("ACADEMIC_CONTENT_HASH", [string]$item.content_hash)
        $noteText = (
            [string]$item.speaker_notes +
            "`r`n`r`n[Slide-ID]`r`n" +
            [string]$item.slide_id +
            "`r`n[Content-Hash]`r`n" +
            [string]$item.content_hash +
            "`r`n[Sources]`r`n"
        )
        foreach ($sourceId in $sources) {
            $noteText += "- " + $sourceId + "`r`n"
        }
        if ($sources.Count -eq 0) {
            $noteText += "- INFORMATION_REQUIRED`r`n"
        }
        Set-Notes $slide $noteText
    }

    $presentation.SaveAs($outputPath, 24)
    $templateHashAfter = (Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash
    $deckIrHashAfter = (Get-FileHash -LiteralPath $deckIrPath -Algorithm SHA256).Hash
    if ($templateHashBefore -ne $templateHashAfter) {
        throw "Template input changed during native fill"
    }
    if ($deckIrHashBefore -ne $deckIrHashAfter) {
        throw "Deck IR input changed during native fill"
    }
    Write-Output ("NATIVE_TEMPLATE_SLIDES=" + $presentation.Slides.Count)
    Write-Output ("TEMPLATE_DESIGNS_AVAILABLE=" + $templateDesignCount)
    Write-Output ("TEMPLATE_DESIGNS_USED=1")
    Write-Output ("REFERENCE_CONTENT_SHAPES_REMOVED=" + $referenceContentShapesRemoved)
    Write-Output ("OUTPUT=" + $outputPath)
}
finally {
    if ($source -ne $null) {
        $source.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($source)
    }
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
