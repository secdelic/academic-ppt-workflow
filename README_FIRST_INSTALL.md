# Academic PPT Workflow v2.7.0-rc4 — First Install

This is a private, local-first release candidate. The repository remains the only canonical authority. Native PptxGenJS is the default and only production generation backend; no external PPT Skill is installed or invoked.

## Windows install

```powershell
$env:PPT_WORKFLOW_HOME = (Resolve-Path .).Path
$env:PPT_WORKSPACE_HOME = Join-Path $env:LOCALAPPDATA "AcademicPPTWorkspace"
powershell -ExecutionPolicy Bypass -File scripts/bootstrap_windows.ps1
powershell -ExecutionPolicy Bypass -File scripts/doctor.ps1
```

Bootstrap creates only project-local `.venv`, `node_modules`, ignored `config/local.yaml`, and a separate local workspace. It does not modify global Python or Node dependencies.

## Create and run a project

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new_project.ps1 `
  -ProjectId "my-study" `
  -Route generate `
  -Quality validated

.\.venv\Scripts\python.exe run_ppt_workflow.py --project "<project-directory>"
```

The four public routes are `generate`, `enhance`, `template-fill`, and
`template-create`; quality is `quick`, `validated`, or `full`. Internal hashes,
IR objects, operation plans, and checkpoints are not user inputs.

Never place patient material inside the repository. Use the configured external workspace and complete manual scientific review before delivery.

This candidate is for private-repository review only. Do not publish a tag or
release until the privacy gate and synthetic CI have passed and a human has
confirmed the GitHub repository is private.
