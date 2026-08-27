# KEEP semantic preservation fixture

- `fixture_provenance`: `SYNTHETIC`
- Purpose: exercise editable text, table, image, shape, mixed-content, and
  native-chart KEEP semantics without any patient or real research content.
- Generation: `generate_fixture.cjs` creates the PPTX at test runtime in a
  temporary directory. No historical audit, benchmark, cache, or private
  workspace is required.
- The PowerPoint COM regression uses the same generated fixture and performs a
  real `SaveAs`; CI unit tests use deterministic OOXML mutations only because
  GitHub-hosted runners are not the PowerPoint authority.
