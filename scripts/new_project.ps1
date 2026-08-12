[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectId,

    [ValidateSet("generate", "enhance", "template-fill", "template-create")]
    [string]$Route = "generate",

    [ValidateSet("quick", "validated", "full")]
    [string]$Quality = "validated",

    [string]$WorkspaceHome = "",
    [string]$ProjectTitle = "INFORMATION_REQUIRED",

    [ValidateSet("zh-CN", "en-US")]
    [string]$Language = "zh-CN",

    [switch]$AllowExisting
)

$ErrorActionPreference = "Stop"
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))

if ($ProjectId -in @(".", "..") -or $ProjectId -notmatch '^[\p{L}\p{N}][\p{L}\p{N}._-]{0,79}$') {
    throw "ProjectId must be 1-80 letters, numbers, dots, underscores, or hyphens and must not contain a path separator."
}

function Resolve-ConfiguredPath([string]$Value, [string]$BasePath) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ([IO.Path]::IsPathRooted($Value)) { return [IO.Path]::GetFullPath($Value) }
    return [IO.Path]::GetFullPath((Join-Path $BasePath $Value))
}

function Resolve-WorkspaceHome([string]$Requested) {
    if (-not [string]::IsNullOrWhiteSpace($Requested)) {
        return (Resolve-ConfiguredPath $Requested (Get-Location).Path)
    }
    if (-not [string]::IsNullOrWhiteSpace($env:PPT_WORKSPACE_HOME)) {
        return (Resolve-ConfiguredPath $env:PPT_WORKSPACE_HOME (Get-Location).Path)
    }

    $localConfig = Join-Path $repoRoot "config\local.yaml"
    if (Test-Path -LiteralPath $localConfig) {
        $raw = Get-Content -LiteralPath $localConfig -Raw -Encoding utf8
        try {
            $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($parsed.paths.workspace_home) {
                return (Resolve-ConfiguredPath ([string]$parsed.paths.workspace_home) $repoRoot)
            }
        }
        catch {
            $match = [regex]::Match(
                $raw,
                '(?m)^\s*workspace_home\s*:\s*["'']?([^"''#\r\n]+)'
            )
            if ($match.Success) {
                return (Resolve-ConfiguredPath $match.Groups[1].Value.Trim() $repoRoot)
            }
        }
    }
    return [IO.Path]::GetFullPath((Join-Path (Split-Path $repoRoot -Parent) "academic-ppt-workspace"))
}

function ConvertTo-YamlScalar([string]$Value) {
    # A JSON string is also a valid YAML scalar and safely preserves punctuation.
    return ($Value | ConvertTo-Json -Compress)
}

function Write-NewUtf8File([string]$Path, [string]$Content) {
    if (Test-Path -LiteralPath $Path) {
        throw "Refusing to overwrite existing project file: $Path"
    }
    Set-Content -LiteralPath $Path -Value $Content -Encoding utf8
}

$workspaceRoot = Resolve-WorkspaceHome $WorkspaceHome
$projectsRoot = [IO.Path]::GetFullPath((Join-Path $workspaceRoot "projects"))
$projectRoot = [IO.Path]::GetFullPath((Join-Path $projectsRoot $ProjectId))
if (-not $projectRoot.StartsWith($projectsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Resolved project path escapes PPT_WORKSPACE_HOME/projects."
}

if ((Test-Path -LiteralPath $projectRoot) -and -not $AllowExisting) {
    throw "Project already exists: $projectRoot. Use -AllowExisting only to add missing scaffold files."
}

$presentationTemplatePath = Join-Path $repoRoot "config\presentation_brief.template.yaml"
$visualTemplatePath = Join-Path $repoRoot "config\visual_brief.template.yaml"
if (-not (Test-Path -LiteralPath $presentationTemplatePath)) {
    throw "Missing repository template: $presentationTemplatePath"
}
if ($Route -ne "enhance" -and -not (Test-Path -LiteralPath $visualTemplatePath)) {
    throw "Missing repository template: $visualTemplatePath"
}

$relativeDirectories = @(
    "input\documents",
    "input\data",
    "input\figures",
    "input\references",
    "input\style_reference",
    "input\template",
    "input\existing",
    "brief",
    "approved_assets",
    "output",
    "audit",
    "cache",
    "private"
)
foreach ($relative in $relativeDirectories) {
    New-Item -ItemType Directory -Force -Path (Join-Path $projectRoot $relative) | Out-Null
}

$templatePath = if ($Route -eq "template-fill") { "input/template/official_template.pptx" } else { "" }
$existingDeckPath = if ($Route -eq "enhance") { "input/existing/existing_deck.pptx" } else { "" }
$inheritExistingStyle = if ($Route -eq "enhance") { "true" } else { "false" }
$preserveTemplateBrand = if ($Route -eq "template-fill") { "true" } else { "false" }
$visualApprovalRequired = if ($Route -eq "enhance") { "false" } else { "true" }

$briefText = Get-Content -LiteralPath $presentationTemplatePath -Raw -Encoding utf8
$briefText = $briefText.Replace("__PROJECT_ID__", $ProjectId)
$briefText = $briefText.Replace("__PROJECT_TITLE_JSON__", (ConvertTo-YamlScalar $ProjectTitle))
$briefText = $briefText.Replace("__ROUTE__", $Route)
$briefText = $briefText.Replace("__QUALITY__", $Quality)
$briefText = $briefText.Replace("__LANGUAGE__", $Language)
$briefText = $briefText.Replace("__TEMPLATE_PATH_JSON__", (ConvertTo-YamlScalar $templatePath))
$briefText = $briefText.Replace("__EXISTING_DECK_PATH_JSON__", (ConvertTo-YamlScalar $existingDeckPath))
$briefText = $briefText.Replace("__INHERIT_EXISTING_STYLE__", $inheritExistingStyle)
$briefText = $briefText.Replace("__PRESERVE_TEMPLATE_BRAND__", $preserveTemplateBrand)
$briefText = $briefText.Replace("__VISUAL_APPROVAL_REQUIRED__", $visualApprovalRequired)

$presentationBriefPath = Join-Path $projectRoot "brief\presentation_brief.yaml"
if (-not (Test-Path -LiteralPath $presentationBriefPath)) {
    Write-NewUtf8File $presentationBriefPath $briefText
}

$visualDraftPath = $null
if ($Route -ne "enhance") {
    $visualDraftPath = Join-Path $projectRoot "brief\visual_brief.DRAFT.yaml"
    if (-not (Test-Path -LiteralPath $visualDraftPath)) {
        $visualDraft = Get-Content -LiteralPath $visualTemplatePath -Raw -Encoding utf8
        try {
            # The repository template is JSON-compatible YAML. Set route without
            # making the user edit an internal or mismatched scope value.
            $visualObject = $visualDraft | ConvertFrom-Json -ErrorAction Stop
            $visualIdPrefix = [regex]::Replace($ProjectId, '[^A-Za-z0-9._-]', '-').Trim('-', '.')
            if ([string]::IsNullOrWhiteSpace($visualIdPrefix) -or $visualIdPrefix -notmatch '^[A-Za-z0-9]') {
                $visualIdPrefix = "project"
            }
            if ($visualIdPrefix.Length -gt 50) { $visualIdPrefix = $visualIdPrefix.Substring(0, 50).TrimEnd('-', '.') }
            $visualObject.visual_brief_id = "$visualIdPrefix-$Route-direction"
            $visualObject.approval.status = "draft"
            $visualObject.approval.approved_by = $null
            $visualObject.approval.approved_at = $null
            if (-not $visualObject.scope) {
                $visualObject | Add-Member -MemberType NoteProperty -Name scope -Value ([pscustomobject]@{ route = $Route })
            }
            elseif ($visualObject.scope.PSObject.Properties.Name -contains "route") {
                $visualObject.scope.route = $Route
            }
            else {
                $visualObject.scope | Add-Member -MemberType NoteProperty -Name route -Value $Route
            }
            $visualDraft = $visualObject | ConvertTo-Json -Depth 20
        }
        catch {
            $visualDraft = $visualDraft.Replace("__PROJECT_ID__", $ProjectId).Replace("__ROUTE__", $Route)
        }
        Write-NewUtf8File $visualDraftPath $visualDraft
    }
}

$assetRegistryPath = Join-Path $projectRoot "approved_assets\asset_registry.yaml"
if (-not (Test-Path -LiteralPath $assetRegistryPath)) {
    $assetRegistry = @"
schema_version: academic-ppt-approved-assets/1
project_id: "$ProjectId"
policy:
  user_approval_required: true
  non_data_assets_only: true
  scientific_evidence_must_be_false: true
assets: []
# Each approved entry must contain:
# - asset_id: stable-local-id
#   file: <filename> # resolved inside approved_assets/
#   role: background | divider_motif | concept_visual | decorative_accent
#   source: user_provided | authorized_reference | user_approved_generated
#   scientific_evidence: false
#   approved_by: INFORMATION_REQUIRED
#   approved_at: INFORMATION_REQUIRED
"@
    Write-NewUtf8File $assetRegistryPath $assetRegistry
}

$nextCommand = "python run_ppt_workflow.py --project `"$projectRoot`""
$result = [ordered]@{
    status = "PROJECT_SCAFFOLD_READY"
    project_id = $ProjectId
    project_root = $projectRoot
    route = $Route
    quality = $Quality
    presentation_brief = $presentationBriefPath
    visual_brief_draft = if ($visualDraftPath) { $visualDraftPath } else { "NOT_CREATED_EXISTING_STYLE_INHERITED" }
    approved_assets_registry = $assetRegistryPath
    next_command = $nextCommand
}
$result | ConvertTo-Json -Depth 4
