param([string]$WorkflowHome=$env:PPT_WORKFLOW_HOME)
$repo=if($WorkflowHome){[IO.Path]::GetFullPath($WorkflowHome)}else{[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))}
Push-Location $repo
try {
  if(git status --porcelain){throw "BLOCKED: worktree is dirty; update must not overwrite local work"}
  git fetch --tags --prune
  if($LASTEXITCODE -ne 0){throw "BLOCKED: git fetch failed"}
  "PASS: fetched remote metadata only. Review and fast-forward manually."
} finally {Pop-Location}
