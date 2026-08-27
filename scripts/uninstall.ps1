param([string]$WorkflowHome=$env:PPT_WORKFLOW_HOME,[switch]$RemoveVenv,[switch]$RemoveNodeModules)
$repo=if($WorkflowHome){[IO.Path]::GetFullPath($WorkflowHome)}else{[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))}
$targets=@();if($RemoveVenv){$targets+=Join-Path $repo ".venv"};if($RemoveNodeModules){$targets+=Join-Path $repo "node_modules"}
foreach($target in $targets){$resolved=[IO.Path]::GetFullPath($target);if(-not $resolved.StartsWith($repo,[StringComparison]::OrdinalIgnoreCase)){throw "BLOCKED: target escapes workflow home"};if(Test-Path $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}}
"PASS: removed only explicitly selected project-local dependencies. Workspace/cache/output were preserved."
