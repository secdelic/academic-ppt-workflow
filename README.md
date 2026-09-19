# Academic PPT Workflow

Local, source-bound academic PowerPoint generation and assisted visual review.
Scientific claims, numbers, figures and citations remain traceable to registered inputs;
final scientific and visual decisions remain with the author.

Routes: **generate**, **enhance**, **template-fill**, **template-create**.
The native PptxGenJS backend preserves editable text, semantic shapes and chart data.
Approved visual plans can use role-aware real PowerPoint Masters/Layouts, a process
quality floor, and optional manually generated presentation-only illustrations.

| Quality | Delivery contract |
|---|---|
| quick | Draft; no visual-floor certification |
| validated | Scientific QA, full render, roles/layouts, Floor and human gates |
| full | validated plus provenance/privacy audit and Anchor/STRETCH opportunity review |

## Start on Windows

Requires Python 3.12, Node.js 20+ and Microsoft PowerPoint. From the repository root:

```powershell
$env:PPT_WORKFLOW_HOME = (Get-Location).Path
$env:PPT_WORKSPACE_HOME = Join-Path (Split-Path $env:PPT_WORKFLOW_HOME -Parent) 'academic-ppt-workspace'
& .\scripts\bootstrap_windows.ps1 -WorkflowHome $env:PPT_WORKFLOW_HOME -WorkspaceHome $env:PPT_WORKSPACE_HOME
& .\scripts\new_project.ps1 -ProjectId 'academic-demo' -Route generate -Quality validated -WorkspaceHome $env:PPT_WORKSPACE_HOME
```

Complete the project brief, registered sources and approved visual brief before running.
Follow the manual to prepare the source-bound visual plan and complete review.
Unresolved science or visual evidence blocks delivery; unsigned review remains a draft.

Keep project inputs, candidates, credentials and run outputs in the external private
workspace. No image API, automatic image generation or external Skill is used.
Candidates require explicit human approval and slide-specific authorization.

**Lifecycle:** assisted use; this consolidation is **UNRELEASED / NEXT**. Existing
version metadata is aligned to `2.7.0-rc4`, not a new release. A second independent
real project remains unvalidated. No aesthetic guarantee or stable-release claim.

- [中文操作说明书](docs/中文操作说明书.md)
- [Architecture](docs/architecture.md)
- [KEEP semantics](docs/KEEP_语义保留合同.md)
- [Changelog](CHANGELOG.md)
