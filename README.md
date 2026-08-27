# Academic PPT Workflow

Academic PPT Workflow is a local-first, source-bound production system for
scientific PowerPoint decks. The repository remains the single canonical
authority; native PptxGenJS is the default and only production rendering
backend.

This checkout is a private release candidate. It is not approved for public
distribution or unattended scientific delivery. Human scientific and visual
review remain mandatory.

## User interface

Users choose one route and one quality level:

- Routes: `generate`, `enhance`, `template-fill`, `template-create`
- Quality: `quick`, `validated`, `full`

Internal evidence objects, hashes, checkpoints, and change-impact state are
managed by the workflow and are not user inputs.

## First install on Windows

```powershell
$env:PPT_WORKFLOW_HOME = (Resolve-Path .).Path
$env:PPT_WORKSPACE_HOME = Join-Path $env:LOCALAPPDATA "AcademicPPTWorkspace"
powershell -ExecutionPolicy Bypass -File scripts/bootstrap_windows.ps1
powershell -ExecutionPolicy Bypass -File scripts/doctor.ps1
```

The bootstrap uses project-local Python and Node dependencies. It does not
modify global Python or Node installations.

The application and user workspace are separate authorities. Production cache
defaults to `<PPT_WORKSPACE_HOME>/projects/<PROJECT_ID>/cache/`; an explicitly
configured `PPT_CACHE_HOME` provides a shared parent with opaque, isolated
project keys. Repository `.cache` is reserved for development and tests.

## Start a project

```powershell
powershell -ExecutionPolicy Bypass -File scripts/new_project.ps1 `
  -ProjectId "my-study" `
  -Route generate `
  -Quality validated
```

Then run the project through the single public entry point:

```powershell
.\.venv\Scripts\python.exe run_ppt_workflow.py --project "<project-directory>"
```

Do not place patient material, real clinical decks, private caches, runtime
logs, or credentials inside the repository. Use the configured external
workspace.

## Documentation

- [Quick guide (Chinese)](docs/README_%E5%BF%AB%E9%80%9F%E4%BD%BF%E7%94%A8.md)
- [User manual (Chinese)](docs/%E7%94%A8%E6%88%B7%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E%E4%B9%A6.md)
- [First install](README_FIRST_INSTALL.md)
- [Architecture](docs/architecture_v2.md)

## Release boundary

The private GitHub release gate must pass privacy, synthetic CI, dependency
lock, and bundle validation before a tag or release is created. A local commit
or successful automated QA does not by itself authorize publication.
