param(
    [Parameter(Mandatory = $true)][string]$DeckIndexCsv,
    [Parameter(Mandatory = $true)][string]$RenderIndexJson
)

$ErrorActionPreference = "Stop"
$decks = Import-Csv -LiteralPath $DeckIndexCsv
if (-not $decks) { throw "Deck index is empty: $DeckIndexCsv" }
$powerPoint = $null
$results = @()
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    foreach ($row in $decks) {
        $pptx = [System.IO.Path]::GetFullPath($row.pptx_path)
        $outDir = [System.IO.Path]::GetDirectoryName($pptx)
        $pdf = [System.IO.Path]::ChangeExtension($pptx, ".powerpoint.pdf")
        $preview = Join-Path $outDir "preview_powerpoint"
        $geometryPath = Join-Path $outDir "powerpoint_actual_geometry.json"
        if (Test-Path -LiteralPath $pdf) { throw "Refusing to overwrite $pdf" }
        if (Test-Path -LiteralPath $preview) { throw "Refusing to overwrite $preview" }
        New-Item -ItemType Directory -Path $preview | Out-Null
        $presentation = $null
        try {
            $presentation = $powerPoint.Presentations.Open($pptx, $true, $false, $false)
            $slideWidth = [double]$presentation.PageSetup.SlideWidth
            $slideHeight = [double]$presentation.PageSetup.SlideHeight
            $slides = @()
            for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
                $slide = $presentation.Slides.Item($i)
                $png = Join-Path $preview ("slide_{0:D3}.png" -f $i)
                $slide.Export($png, "PNG", 1600, 900)
                $shapes = @()
                foreach ($shape in @($slide.Shapes)) {
                    $entry = [ordered]@{
                        shape_id = [int]$shape.Id
                        name = [string]$shape.Name
                        type = [int]$shape.Type
                        left = [math]::Round([double]$shape.Left, 3)
                        top = [math]::Round([double]$shape.Top, 3)
                        width = [math]::Round([double]$shape.Width, 3)
                        height = [math]::Round([double]$shape.Height, 3)
                        off_slide = ([double]$shape.Left -lt -0.5 -or [double]$shape.Top -lt -0.5 -or
                            ([double]$shape.Left + [double]$shape.Width) -gt ($slideWidth + 0.5) -or
                            ([double]$shape.Top + [double]$shape.Height) -gt ($slideHeight + 0.5))
                        has_text = $false
                        text = ""
                        bound_left = $null
                        bound_top = $null
                        bound_width = $null
                        bound_height = $null
                        overflowing = $null
                        font_size = $null
                        font_name = ""
                        font_name_far_east = ""
                        auto_size = $null
                        margin_left = $null
                        margin_right = $null
                        margin_top = $null
                        margin_bottom = $null
                        group_child_count = 0
                    }
                    try {
                        if ([int]$shape.Type -eq 6) { $entry.group_child_count = [int]$shape.GroupItems.Count }
                    } catch {}
                    try {
                        if ($shape.HasTextFrame -and $shape.TextFrame2.HasText) {
                            $entry.has_text = $true
                            $entry.text = [string]$shape.TextFrame2.TextRange.Text
                            $entry.bound_left = [math]::Round([double]$shape.TextFrame2.TextRange.BoundLeft, 3)
                            $entry.bound_top = [math]::Round([double]$shape.TextFrame2.TextRange.BoundTop, 3)
                            $entry.bound_width = [math]::Round([double]$shape.TextFrame2.TextRange.BoundWidth, 3)
                            $entry.bound_height = [math]::Round([double]$shape.TextFrame2.TextRange.BoundHeight, 3)
                            try { $entry.overflowing = [bool]$shape.TextFrame2.Overflowing } catch { $entry.overflowing = $null }
                            try { $entry.font_size = [math]::Round([double]$shape.TextFrame2.TextRange.Font.Size, 2) } catch { $entry.font_size = $null }
                            try { $entry.font_name = [string]$shape.TextFrame2.TextRange.Font.Name } catch {}
                            try { $entry.font_name_far_east = [string]$shape.TextFrame2.TextRange.Font.NameFarEast } catch {}
                            try { $entry.auto_size = [int]$shape.TextFrame2.AutoSize } catch { $entry.auto_size = $null }
                            try { $entry.margin_left = [math]::Round([double]$shape.TextFrame2.MarginLeft, 3) } catch {}
                            try { $entry.margin_right = [math]::Round([double]$shape.TextFrame2.MarginRight, 3) } catch {}
                            try { $entry.margin_top = [math]::Round([double]$shape.TextFrame2.MarginTop, 3) } catch {}
                            try { $entry.margin_bottom = [math]::Round([double]$shape.TextFrame2.MarginBottom, 3) } catch {}
                        }
                    } catch {}
                    $shapes += [pscustomobject]$entry
                }
                $slides += [pscustomobject]@{
                    slide_number = $i
                    width = $slideWidth
                    height = $slideHeight
                    png = $png
                    shapes = $shapes
                }
            }
            $presentation.SaveAs($pdf, 32)
            if (-not (Test-Path -LiteralPath $pdf)) { throw "PowerPoint PDF was not generated: $pdf" }
            $geometry = [ordered]@{
                source_pptx = $pptx
                slide_width = $slideWidth
                slide_height = $slideHeight
                slide_count = [int]$presentation.Slides.Count
                slides = $slides
            }
            $geometry | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $geometryPath -Encoding utf8
            $pngCount = @(Get-ChildItem -LiteralPath $preview -Filter "slide_*.png").Count
            $results += [pscustomobject]@{
                project_key = $row.project_key
                arm = $row.arm
                pptx = $pptx
                pdf = $pdf
                preview = $preview
                geometry = $geometryPath
                pptx_slide_count = [int]$presentation.Slides.Count
                png_count = $pngCount
                powerpoint_open = $true
                pdf_exists = (Test-Path -LiteralPath $pdf)
            }
        } finally {
            if ($presentation) {
                $presentation.Close()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
            }
        }
    }
} finally {
    if ($powerPoint) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
$parent = Split-Path -Parent $RenderIndexJson
if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
$results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $RenderIndexJson -Encoding utf8
Write-Output ("POWERPOINT_DECKS={0}" -f $results.Count)
