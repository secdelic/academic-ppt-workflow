param([string]$WorkflowHome=$env:PPT_WORKFLOW_HOME,[string]$Destination=(Join-Path $env:USERPROFILE ".codex\skills\academic-ppt-production"))
$repo=if($WorkflowHome){[IO.Path]::GetFullPath($WorkflowHome)}else{[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))}
$source=Join-Path $repo "skills\academic-ppt-production"
if(-not(Test-Path (Join-Path $source "SKILL.md"))){throw "BLOCKED: Canonical Skill candidate is not present because Case B, Case C, and second-device gates have not passed."}
if(Test-Path $Destination){throw "BLOCKED: destination already exists; refusing overwrite: $Destination"}
Copy-Item -LiteralPath $source -Destination $Destination -Recurse
"PASS: installed thin skill to $Destination"
