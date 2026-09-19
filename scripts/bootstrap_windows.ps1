param(
    [string]$WorkflowHome = $env:PPT_WORKFLOW_HOME,
    [string]$WorkspaceHome = $env:PPT_WORKSPACE_HOME,
    [string]$PythonExecutable = "",
    [string]$NodeExecutable = "",
    [switch]$SkipSmoke
)
$ErrorActionPreference = "Stop"
$repo = if ($WorkflowHome) { [IO.Path]::GetFullPath($WorkflowHome) } else { [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..")) }
if (-not $WorkspaceHome) { $WorkspaceHome = Join-Path (Split-Path $repo -Parent) "academic-ppt-workspace" }
$workspace = [IO.Path]::GetFullPath($WorkspaceHome)
$venv = Join-Path $repo ".venv"
$reportDir = Join-Path $workspace "install_receipts"
New-Item -ItemType Directory -Force -Path $workspace,$reportDir | Out-Null

function Resolve-Python3([string]$requested) {
    $candidates = @($requested, $env:PPT_PYTHON, (Get-Command python.exe -ErrorAction SilentlyContinue).Source, (Get-Command py.exe -ErrorAction SilentlyContinue).Source) | Where-Object { $_ }
    foreach ($candidate in $candidates) {
        $probeArgs = if ((Split-Path $candidate -Leaf) -ieq "py.exe") { @("-3.12", "-c", "import sys; assert (3,12) <= sys.version_info < (3,13); print(sys.executable)") } else { @("-c", "import sys; assert (3,12) <= sys.version_info < (3,13); print(sys.executable)") }
        $resolved = & $candidate @probeArgs 2>$null
        if ($LASTEXITCODE -eq 0 -and $resolved) { return ([string]$resolved).Trim() }
    }
    throw "BLOCKED: Python 3.12 not found. Set -PythonExecutable or PPT_PYTHON."
}
function Resolve-Node([string]$requested) {
    $candidate = if ($requested) { $requested } elseif ($env:PPT_NODE) { $env:PPT_NODE } else { (Get-Command node.exe -ErrorAction SilentlyContinue).Source }
    if (-not $candidate) { throw "BLOCKED: Node.js not found. Set -NodeExecutable or PPT_NODE." }
    $major = [int]((& $candidate -p "process.versions.node.split('.')[0]") 2>$null)
    if ($LASTEXITCODE -ne 0 -or $major -lt 20) { throw "BLOCKED: Node.js >=20 required." }
    return $candidate
}

$pythonCmd = Resolve-Python3 $PythonExecutable
if (-not (Test-Path (Join-Path $venv "Scripts\python.exe"))) {
    & $pythonCmd -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw "BLOCKED: Python virtual environment creation failed" }
}
$venvPython = Join-Path $venv "Scripts\python.exe"
& $venvPython -c "import sys; assert (3,12) <= sys.version_info < (3,13)"
if ($LASTEXITCODE -ne 0) { throw "BLOCKED: Existing virtual environment requires Python 3.12" }
& $venvPython -m pip install --disable-pip-version-check -r (Join-Path $repo "requirements-lock.txt")
if ($LASTEXITCODE -ne 0) { throw "BLOCKED: locked Python dependency installation failed" }

$node = Resolve-Node $NodeExecutable
$npm = Join-Path (Split-Path $node -Parent) "npm.cmd"
if (-not (Test-Path $npm)) { $npm = (Get-Command npm.cmd -ErrorAction SilentlyContinue).Source }
if (-not $npm) { throw "BLOCKED: npm not found next to Node or on PATH" }
Push-Location $repo
try { & $npm ci --ignore-scripts --no-audit --no-fund } finally { Pop-Location }
if ($LASTEXITCODE -ne 0) { throw "BLOCKED: npm ci failed" }

foreach ($name in "input","staging","output","audit","projects") { New-Item -ItemType Directory -Force -Path (Join-Path $workspace $name) | Out-Null }
$local = Join-Path $repo "config\local.yaml"
if (-not (Test-Path $local)) {
    $config = [ordered]@{schema_version="academic-ppt-local-config/1";paths=[ordered]@{workspace_home=$workspace;input_root=(Join-Path $workspace "input");staging_root=(Join-Path $workspace "staging");output_root=(Join-Path $workspace "output");audit_root=(Join-Path $workspace "audit");asset_home=(Join-Path $repo "visual_workspaces");cache_home=$null}}
    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $local -Encoding utf8
}

$powerPoint = $null; $ppVersion = "NOT_FOUND"
try { $powerPoint = New-Object -ComObject PowerPoint.Application; $ppVersion = $powerPoint.Version } catch { }
finally { if ($powerPoint) { $powerPoint.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null } }
if ($ppVersion -eq "NOT_FOUND") { throw "BLOCKED: Microsoft PowerPoint COM is required for the Windows production install" }

$smoke = "SKIPPED"
if (-not $SkipSmoke) {
    $env:PPT_NODE_MODULES = Join-Path $repo "node_modules"
    & $node (Join-Path $repo "scripts\build_visual_workspaces.mjs") (Join-Path $workspace "synthetic_smoke\visual_workspaces")
    if ($LASTEXITCODE -ne 0) { throw "BLOCKED: synthetic workspace smoke failed" }
    $smoke = "PASS"
}
$receipt = [ordered]@{status="PASS";workflow_home=$repo;workspace_home=$workspace;python=(& $venvPython --version);node=(& $node --version);powerpoint=$ppVersion;smoke=$smoke;global_python_modified=$false;global_node_modified=$false;completed_at=(Get-Date).ToString("o")}
$receipt | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $reportDir "bootstrap_report.json") -Encoding utf8
$receipt | ConvertTo-Json -Depth 4
