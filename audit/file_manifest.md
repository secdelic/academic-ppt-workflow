# File Manifest

## Read during Phase 0

- Empty workspace root `E:\PPT`
- Local runtime/package metadata
- Installed-font names
- Applicable presentation, PDF, spreadsheet, and document skill instructions

## Created - governance and documentation

- `AGENTS.md`
- `README_使用说明.md`
- `docs/architecture.md`
- `audit/00_preflight_report.md`
- `audit/runtime_discovery_report.md`
- `audit/existing_file_inventory.csv`
- `audit/source_manifest.csv`
- `audit/unresolved_items.md`
- `audit/validation_summary.md`
- `audit/file_manifest.md`

## Created - configuration and contracts

- `config/workflow.yaml`
- `config/design_system.yaml`
- `config/profiles.yaml`
- `config/qa_rules.yaml`
- `config/dependency_lock.json`
- `brief/presentation_brief.yaml`
- `package.json`
- `requirements-lock.txt`

## Created - implementation

- `run_ppt_workflow.py`
- `scripts/render_pptx.ps1`
- `scripts/generate_deck_artifact.mjs`
- `scripts/generate_deck_pptxgen.mjs`
- `scripts/academic_ppt/` Python package

## Created - tests

- `tests/test_workflow.py`
- `tests/fixtures/synthetic_brief.yaml`
- `tests/fixtures/synthetic_project/` fully synthetic source package

## Created - formal synthetic outputs

- Default PptxGenJS run:
  `output/synthetic_academic_workflow_demo/20260728_104100_pptxgen_final/`
- Internal Artifact Tool validation run:
  `output/synthetic_academic_workflow_demo/20260728_104000_artifact_final/`

Each formal run contains PPTX, PDF, per-slide PNG previews, contact sheet,
outline, storyboard, speaker notes, source and slide manifests, traceability
maps, QA report, manual checklist, generation log, runtime manifest, and
dependency manifest.

## Modified source files

None. Files under the synthetic fixture were read and hash-verified but not
modified during workflow execution.
