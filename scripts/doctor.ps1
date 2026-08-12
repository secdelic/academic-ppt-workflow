param([string]$WorkflowHome = $env:PPT_WORKFLOW_HOME,[string]$WorkspaceHome = $env:PPT_WORKSPACE_HOME)
$ErrorActionPreference="Stop"
$repo=if($WorkflowHome){[IO.Path]::GetFullPath($WorkflowHome)}else{[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))}
$workspace=if($WorkspaceHome){[IO.Path]::GetFullPath($WorkspaceHome)}else{[IO.Path]::GetFullPath((Join-Path (Split-Path $repo -Parent) "academic-ppt-workspace"))}
$checks=@()
function Add-Check($name,$status,$detail,$fix=""){$script:checks += [pscustomobject]@{check=$name;status=$status;detail=$detail;fix=$fix}}
$py=Join-Path $repo ".venv\Scripts\python.exe"; Add-Check "python_venv" ($(if(Test-Path $py){"PASS"}else{"BLOCKED"})) $py "Run scripts/bootstrap_windows.ps1"
$nodeModules=Join-Path $repo "node_modules\pptxgenjs"; Add-Check "pptxgenjs_lock" ($(if((Test-Path $nodeModules)-and(Test-Path (Join-Path $repo "package-lock.json"))){"PASS"}else{"BLOCKED"})) $nodeModules "Run bootstrap"
Add-Check "local_config" ($(if(Test-Path (Join-Path $repo "config\local.yaml")){"PASS"}else{"WARNING"})) "config/local.yaml" "Run bootstrap or copy local.example.yaml"
Add-Check "workspace" ($(if(Test-Path $workspace){"PASS"}else{"BLOCKED"})) $workspace "Run bootstrap"
$pp=$null;try{$pp=New-Object -ComObject PowerPoint.Application;Add-Check "powerpoint" "PASS" $pp.Version}catch{Add-Check "powerpoint" "BLOCKED" $_.Exception.Message "Install/repair Microsoft PowerPoint"}finally{if($pp){$pp.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($pp)|Out-Null}}
$soffice=(Get-Command soffice.com -ErrorAction SilentlyContinue);if(-not $soffice){$soffice=(Get-Command soffice.exe -ErrorAction SilentlyContinue)};Add-Check "libreoffice_optional" ($(if($soffice){"PASS"}else{"WARNING"})) ($(if($soffice){$soffice.Source}else{"not found"})) "Optional: install LibreOffice or set PPT_SOFFICE"
$fonts=@("Microsoft YaHei","Aptos","Arial");$fontKeys=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -ErrorAction SilentlyContinue).PSObject.Properties.Name -join "|";foreach($font in $fonts){Add-Check "font_$font" ($(if($fontKeys -match [regex]::Escape($font)){"PASS"}else{"WARNING"})) $font "Install font or accept configured fallback"}
$overall=if($checks.status -contains "BLOCKED"){"BLOCKED"}elseif($checks.status -contains "WARNING"){"WARNING"}else{"PASS"}
$result=[ordered]@{status=$overall;workflow_home=$repo;workspace_home=$workspace;checks=$checks;checked_at=(Get-Date).ToString("o")};$result|ConvertTo-Json -Depth 5
if($overall -eq "BLOCKED"){exit 2}
