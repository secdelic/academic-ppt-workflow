param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutputPptx
)

$ErrorActionPreference = "Stop"
$inputResolved = [System.IO.Path]::GetFullPath($InputPptx)
$outputResolved = [System.IO.Path]::GetFullPath($OutputPptx)
if (-not [System.IO.File]::Exists($inputResolved)) { throw "Input PPTX does not exist" }
if ([System.IO.File]::Exists($outputResolved)) { throw "Output PPTX already exists" }
[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($outputResolved)) | Out-Null

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputResolved, $true, $false, $false)
    $presentation.SaveAs($outputResolved, 24)
    Write-Output ("POWERPOINT_SAVEAS_PASS version={0} build={1}" -f $powerPoint.Version, $powerPoint.Build)
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($powerPoint -ne $null) { $powerPoint.Quit() }
    if ($presentation -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($powerPoint -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
