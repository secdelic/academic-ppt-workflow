# Changelog

## v2.7.0-rc4

- Replaced raw shape-tree equality as the sole KEEP gate with three fail-closed
  checks: presentation identity, editable semantic identity, and render identity.
- Added one canonical semantic-manifest helper and an explicit normalization
  allowlist for PowerPoint SaveAs serialization behavior. Raw OOXML hashes remain
  diagnostic and true text, table, image, geometry, z-order, Master/Layout, and
  chart-data mutations still fail.
- Added the fully synthetic `SECOND_DEVICE_COM_KEEP_TABLE_NORMALIZATION`
  regression with real local PowerPoint COM SaveAs and exact render comparison.
- Scientific semantics, StoryGraph, ArtDirection, renderer drawing semantics,
  cache/lineage contracts, native PptxGenJS, and `NO_GO_EXTERNAL_SKILL` remain
  unchanged. RC1, RC2, and RC3 remain immutable historical candidates.

## v2.7.0-rc3

- Separated enhance presentation lineage from scientific claim lineage.
  Existing decks remain registered and fingerprinted as presentation baselines
  but no longer enter the scientific claim-coverage denominator by default.
- Added one canonical source-role classifier, explicit inherited-content
  boundaries, and fail-closed changed-slide scientific bindings.
- Added the synthetic `SECOND_DEVICE_EXISTING_DECK_LINEAGE` regression for
  separate application/workspace roots. Scientific semantics, visual design,
  rendering, cache schema, and external-Skill policy remain unchanged.
- `v2.7.0-rc1` and `v2.7.0-rc2` remain immutable historical candidates.

## v2.7.0-rc2

- Corrected the production project-cache authority for deployments where the
  application and user workspace are separate.
- Production now resolves cache by explicit project cache root, configured
  `PPT_CACHE_HOME`, then `<ProjectRoot>/cache`; repository `.cache` remains
  available only for development, tests, and legacy compatibility.
- Added synthetic cross-device, containment, traversal, project-isolation, and
  public-route regression coverage. Scientific, visual, renderer, source-binding,
  and external-Skill contracts are unchanged.
- `v2.7.0-rc1` remains preserved as the historical candidate whose local RC
  validation passed but second-device deployment exposed the cache conflict.

## v2.7.0-rc1

- Standardized the user-facing interface around four routes (`generate`,
  `enhance`, `template-fill`, and `template-create`) and three quality modes.
- Added the approved `visual_brief.yaml` art-direction contract without adding
  a second scientific or visual authority.
- Kept the scientific core, privacy controls, native PptxGenJS backend, and
  `NO_GO_EXTERNAL_SKILL` contract frozen.
- Expanded the private release allowlist to include the runtime's checked-in
  extension registry, prompts, visual template registry, and sanitized template
  documentation while continuing to exclude local audit, benchmark, private,
  cache, staging, and output material.
- The candidate remains unpublished until the private GitHub gate and synthetic
  CI are complete.

## v2.6.0-rc1

- Added six synthetic Canonical Visual Workspace candidates with empty human-review gates.
- Added composition grammar and role/variant contracts without introducing a new IR.
- Added portable `PPT_*` home resolution and ignored machine-local configuration.
- Added project-local Windows bootstrap, doctor, update and uninstall scripts.
- Added privacy allowlist, release exclusions, synthetic GitHub Actions and local release validation.
- Kept native PptxGenJS as the only production backend and preserved `NO_GO_EXTERNAL_SKILL`.
- Real Case B, Case C, second-device install and human workspace scores remain pending; the release is not production-ready.
