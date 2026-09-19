# Architecture and contracts

## Authority and flow

Registered immutable input -> evidence and source/claim bindings -> presentation
brief -> Deck IR -> approved visual brief and separate visual plan -> native
semantic composition -> real role-aware Master/Layout -> native scientific content
-> render and scientific/file/geometry QA -> Visual Quality Floor -> optional
Ambition and manual asset handoff -> human approval -> image-only assembly -> new
render/QA -> scientific and visual Human Gate.

`run_ppt_workflow.py` is the public entry point. It resolves the project and route,
runs `academic_ppt.runner`, then applies `visual_workflow.finalize_run` before
`project_interface` can publish to final. Each run has a new directory. BLOCK and
HUMAN_REVIEW remain drafts. Template-create creates only a style profile.

## Scientific authority

Deck IR, evidence/source contracts and KEEP remain the scientific authority.
Stable slide IDs derive from deck identity and logical slide keys, not page order.
Normalized content hashes bind scientific planning; binary PPTX hashes bind exact
reviewed bytes. Missing facts are INFORMATION_REQUIRED. The deterministic core
registers exact excerpts as candidate claims and never invents conclusions.

An approved `assisted-visual-plan/1` sidecar binds the IR canonical hash and the
complete ordered slide ID/index set. It cannot add source text or scientific
meaning. Native semantic composition checks nodes, labels, relations, direction,
grouping, ordering and protected objects before rendering. Native Art Direction
can apply explicitly approved typography/palette operations while preserving
text, notes, chart data, media, source bindings and unchanged slide objects.

Source roles distinguish SCIENTIFIC_SOURCE, PRESENTATION_BASELINE,
STYLE_REFERENCE, FORMAL_TEMPLATE and APPROVED_VISUAL_ASSET. An existing deck is a
presentation baseline by default; inherited unverified content needs human review.
Only explicit content-reference use may register its content as scientific source.
Style-only extraction excludes slide text, notes, alt text, chart values and media.
External OOXML relationships are detected and never followed.

## Real Master/Layout

`master_layout.py` creates actual shared parts and slide-to-layout relationships.
Seven layout roles share a theme, safe zones, font roles and footer treatment;
Master elements contain no page-specific science. Original source master/theme
parts remain intact. Shared IDs are globally unique to satisfy PowerPoint.
Native title, source and page-number placeholders inherit typography; original
scientific geometry and other native objects remain at slide level.

Official template authority wins: template-fill preserves its Master/Layout,
brand and mandatory zones. Enhance preserves existing layouts through explicit
per-slide exception receipts. A role label alone is not evidence of materialized
layout. Master scene reuse requires approval for every bound slide. Crop, scale,
position and opacity retain original asset bytes; evidence layouts reject scenes.
The current family requires a wide canvas and already validated fonts.

## Quality modes and evidence

Quick makes no Floor claim. Validated/full use complete scientific processing and
rendering; full adds Ambition plus an OOXML/private-path/external-relationship
scan and provenance/geometry audit. Validated Enhance uses full validation;
quick Enhance can use the bounded fast path.

`visual_quality_floor.py` is the sole evaluator. `config/visual_quality` is its
versioned contract authority; contract hashes cover its own generic files only.
Missing required evidence blocks, warnings remain visible, and mandatory human
review is never synthesized. Declared Art Direction, execution receipts and
render observations are distinct. Family repetition is advisory, not an aesthetic
score. Review signatures bind exact deck, render set and ordered slide IDs.

`visual_ambition.py` independently evaluates narrative Anchor roles and Floor,
Target or Stretch opportunities from approved controlled contexts. It cannot
clear Floor failures, change science or force illustrations. Full-mode opportunity
review binds the Ambition plan as well as deck/render hashes.

## Manual visual assets

`illustration_requests.py` creates closed conceptual requests; arbitrary source
text, medical evidence imitations and private paths cannot enter exported packets.
`manual_image_handoff.py` writes copy/paste packets only with an explicit enable
flag. A separate import flag is required to scan the authorized external inbox.
No API, network image generation, diffusion model or credential handler exists.

Intake checks containment, links, filename/request identity, MIME, dimensions,
size budgets and SHA256. Imported files remain unapproved CANDIDATE records.
`candidate_preview.py` renders the immutable baseline locally and creates clearly
labelled raster comparison artifacts. Those artifacts do not replace the native
scientific authority and cannot be used as scientific source figures.

`asset_approval.py` accepts explicit human CSV decisions and creates a new immutable
approved batch. Refinement or file presence never grants approval. Authorization
binds image, exact baseline, slide, focal region and reviewed preview; integrity
hashes detect later modification. Native assembly only adds approved pictures into
empty authorized slots and preserves original scientific parts. New rendering and
final human asset/scientific/visual review remain mandatory after assembly.
`visual_workflow --audit-assembly` verifies image-only changes and rerenders;
`--review-run` re-evaluates the exact reviewed Deck without regeneration.

## Incremental Enhance and cache

Full Enhance requires previous Deck IR and exactly one stable slide, section or
figure selector. Managed object conflicts block without an explicit override.
Quick fast Enhance requires a valid project cache, explicit operation plan and
prepared changed-slide candidate. Supported operations are KEEP, REPLACE and
INSERT_AFTER. Candidate scientific content remains source-bound.

Cache is a validated snapshot, not a second evidence authority. It uses UTF-8
JSON, opaque project keys, SHA256 delta, immutable generations and atomic promotion.
A failed run cannot promote a partial generation. Source removal, missing impact
edges, changed scientific definitions or global Master/theme changes escalate to
full validation. Clinical cache containment is checked before cache creation.
A different directory on the same host is not second-device validation.

Internal automation parameters remain available: `--existing-pptx`,
`--previous-deck-ir`, `--update-slide` / `--update-section` / `--update-figure`,
`--operation-plan`, `--changed-candidate-pptx`, `--workflow-mode`, `--resume`,
`--run-id`, `--retry-slide` and `--retry-stage`. Retry selects one allowed affected
stage; resume requires matching recorded input fingerprints. These mechanisms do
not bypass the public validated/full visual gate.

## File, privacy and release boundaries

Application code and external project workspaces are separate. Inputs are read-only;
intermediates and outputs are run-scoped. No formal output is silently overwritten.
The release builder and privacy scanner use `release_exclusion_rules.yaml` as the
single file-selection authority. Runtime audits, private inputs, generated Decks,
manual candidates, approvals and local configuration are excluded. Generic
synthetic fixtures and source templates remain included.

Automated privacy scans cannot decide whether a clinical image is safe; human
review remains necessary. Native text/shapes/charts stay editable, while imported
raster figures stay replaceable images. A second independent real-project
validation remains required before cross-project visual generalization or stable
release claims. See the [Chinese manual](中文操作说明书.md) and
[KEEP contract](KEEP_语义保留合同.md).
