# Architecture and Contracts

## Pipeline

`preflight -> inventory -> extraction -> evidence -> outline -> storyboard ->
PPTX -> render -> visual QA -> scientific QA -> package -> manual review`

Every stage writes a completion marker to the run-specific staging directory.
`--resume` reuses a completed stage only when its recorded input fingerprint
matches; `--from-stage` starts at a named stage and preserves earlier products.

## Path contract

Resolution order is CLI argument, environment variable, project configuration,
then repository-relative default. No implementation path depends on a fixed
drive letter.

## Input contract

- `input/` is immutable for a run.
- Supported extensions are declared in `config/workflow.yaml`.
- Each file receives a SHA-256 `source_id` and source-manifest record.
- OCR is disabled by default and is never silently invoked.

## Evidence contract

The deterministic core never asks a language model to invent or complete
content. It registers exact source excerpts as candidate claims. Explicit
`CLAIM|...` records may provide stronger metadata; all other candidates require
manual review. Missing required content becomes `INFORMATION_REQUIRED`.

## Slide contract

No PPTX generation can occur without both `deck_outline.md` and
`storyboard.csv`. Each storyboard row has one purpose and one key message,
source IDs, layout, visual type, citation requirement, notes, confidence, and
manual-review state.

## Output contract

Formal results are written to a new
`output/<project_name>/<YYYYMMDD_HHMMSS>/` directory. A run never overwrites an
existing formal output. The status is one of:

- `READY_FOR_ASSISTED_USE`
- `PARTIALLY_READY`
- `BLOCKED`

`READY_FOR_ASSISTED_USE` means the automated gates passed. It never means final
scientific approval.
