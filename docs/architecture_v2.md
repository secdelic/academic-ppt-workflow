# Academic PPT Workflow v2 Architecture

Status: controlled-fusion implementation on branch
`workflow/academic-ppt-v2-upstream-benchmark`.

The v1 PptxGenJS path remains the default and rollback backend. v2 adds small,
repository-owned modules around that path; it does not install or embed an
upstream Skill.

```text
Source Intake
    ↓
Evidence Registry
    ↓
Presentation Brief
    ↓
Narrative Router
    ↓
Deck Intermediate Representation
    ↓
Slide Planner
    ↓
Backend Router
    ├─ PptxGenJS Stable Backend (default)
    ├─ Native PowerPoint Shape/SVG Visual Path
    └─ Native Template Fill Path (PowerPoint COM, Windows)
    ↓
Render
    ↓
Scientific QA
    ↓
Visual QA
    ↓
OOXML QA
    ↓
Author Review
```

## Top-level routes

Exactly one route is active per run:

- `generate`: source-bound deck generation through the stable PptxGenJS backend.
- `create-style-profile`: read-only PPTX/POTX style extraction.
- `fill-template`: new presentation using a reference deck's native
  master/layout contract; reference slide content is not copied into evidence.
- `enhance-existing`: a new output copy with explicitly selected semantic
  slides replaced; the supplied deck is never overwritten.

Task route and generation backend are separate decisions. `--backend` cannot
select a task lifecycle.

## Deck IR contract

`deck_ir.json` is the canonical machine-readable presentation plan. Required
deck and slide fields include:

- `deck_id`, `deck_version`;
- `slide_id`, `slide_revision`, `slide_role`, `narrative_mode`;
- `key_message`, `claim_ids`, `source_ids`;
- `layout_id`, `backend`, `template_layout_id`;
- `visual_assets`, `editable_object_requirements`;
- `speaker_notes`, `manual_review_required`;
- `content_hash`, `render_hash`.

Internal fields `logical_slide_key`, `section_id`, `figure_ids`,
`managed_object_ids`, and `powerpoint_slide_id` support stable identity and
incremental selection. A stable `slide_id` is derived from the deck identity
and logical slide key, never the page number. Reordering therefore does not
change semantic identity.

Canonical hashes are calculated over normalized JSON and object manifests.
Binary PPTX hashes are recorded for integrity but are not treated as the only
determinism measure.

## Reference boundaries

Reference modes are:

- `style-only` (default): structural and style metadata only;
- `template-fill`: authorized use of native masters/layouts for new content;
- `content-reference`: explicit content use, subject to the evidence registry;
- `protected-reference`: hash, package, and safety metadata only.

`style-only` output must not contain slide text, speaker notes, alt text, chart
values, media bytes, patient data, private institutional content, or resolved
external links. External OOXML relationships are registered as warnings and
are never followed.

## Incremental update boundary

An update:

1. resolves exactly one selector (`slide_id`, `section_id`, or `figure_id`);
2. records and hashes the source deck;
3. creates a run-specific pre-update backup;
4. builds a candidate from the canonical Deck IR;
5. replaces only selected semantic slides in a new copy;
6. writes a slide-level diff;
7. reruns affected-slide rendering and global scientific, relationship, page
   count, and consistency checks.

Workflow-managed objects use stable names such as
`awf:<slide_id>:key-message`. Unmanaged user objects on untouched pages are
preserved. A detected managed-object conflict blocks unless the user explicitly
allows that specific overwrite. Cross-platform raw OOXML slide splicing is
deferred because chart, media, relationship, and content-type collisions can
silently corrupt a deck.

## Scientific visual boundary

Repository-owned declarative templates cover cohort, CONSORT, STROBE, DAG,
target-trial, bioinformatics, single-cell, multi-omics, mechanism, timeline,
and study-design schematics. PowerPoint shapes and local SVG are the stable
paths. Graphviz and Mermaid are optional, disabled-by-default local renderers
that require capability detection and no network access.

Mechanism visuals require one of:

- `demonstrated_mechanism`;
- `supported_interpretation`;
- `proposed_mechanism`;
- `hypothesis_only`.

The label controls wording but never upgrades the evidence.

Native OMML and a general SVG-to-DrawingML compiler remain benchmark-only until
they pass package safety, PowerPoint-open, render-fidelity, and editability
tests. Plain-text or local SVG fallbacks must be disclosed.

## Rollback

The immutable rollback point is local tag `academic-ppt-v1-ready` at commit
`834163c324727f3a782c12c710b1c5c44d2a8427`. No remote, merge, or push is part
of this upgrade. The default backend remains `pptxgenjs`.
