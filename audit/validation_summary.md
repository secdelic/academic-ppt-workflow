# Validation Summary

Validation date: 2026-07-28

## Unit and syntax checks

- Python compile check: pass.
- Node syntax checks for both deck generators: pass.
- Python unit tests: 5/5 pass.
- Path-priority, read-only hashing, text/CSV extraction, configuration loading,
  and no-completion claim registration were exercised.

## End-to-end synthetic tests

Two independent six-slide runs used only the files under
`tests/fixtures/synthetic_project/`.

### Default PptxGenJS backend

- Output run: `20260728_104100_pptxgen_final`
- PPTX SHA-256:
  `33F9A1E043822382DCE3CEFCE8A3981D2EEE2970F5C3579EB7B88211241628A3`
- PDF SHA-256:
  `EDBECADB0AADD5BB2BC160E2396D4D03E5AA09E7EB2CFD2EE1674496ABA4B59B`
- PowerPoint COM open/export: pass.
- PPTX slides / PDF pages / PNG previews: 6 / 6 / 6.
- Editable native chart part: present (`ppt/charts/chart1.xml`).
- Speaker notes with `[Sources]`: 6/6.
- External relationships: 0.
- Official overflow test: pass.
- Individual 1600 x 900 slide inspection: pass.

### Internal Artifact Tool validation backend

- Output run: `20260728_104000_artifact_final`
- PowerPoint COM open/export: pass.
- PPTX slides / PDF pages / PNG previews: 6 / 6 / 6.
- Speaker notes with `[Sources]`: 6/6.
- External relationships: 0.
- Official overflow test: pass.
- Individual render/contact-sheet inspection: pass.

## Resume test

A synthetic run was intentionally stopped at `generate` by supplying a missing
Node executable. It returned canonical blocked exit code 2. The same run ID was
then resumed with `--resume --from-stage generate`; prior inventory, extraction,
evidence, outline, and storyboard artifacts were reused and the run completed
successfully.

## Scientific boundary

All test values are explicitly labeled synthetic. No external literature,
patient-level data, investigator identity, ethics approval, clinical result, or
mechanism claim was introduced. Automated readiness does not replace human
scientific approval.
