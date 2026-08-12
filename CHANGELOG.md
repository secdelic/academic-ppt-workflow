# Changelog

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
