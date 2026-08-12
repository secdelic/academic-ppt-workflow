# `academic-ppt-production` Canonical Skill Plan

Current decision: `DEFER`.

The repository remains the canonical authority. No production Skill is
installed or created in this round.

If later promoted, the single Skill will be a thin layer:

```text
academic-ppt-production/
├─ SKILL.md
├─ workflows/
│  ├─ generate-deck.md
│  ├─ create-style-profile.md
│  ├─ fill-template.md
│  ├─ enhance-existing.md
│  └─ audit-deck.md
├─ references/
└─ scripts/
```

It may perform only task recognition, one-route selection, repository workflow
invocation, input/output contract enforcement, scientific guardrails, QA, and
final-state reporting. Detailed parsers, generators, renderers, tests, and
configuration remain in `$PPT_WORKFLOW_HOME`; they are not duplicated into `SKILL.md`.

Promotion requires:

- controlled benchmark evidence for at least one adopted capability;
- unchanged canonical scientific and traceability gates;
- effective rollback to the v1 PptxGenJS path;
- no unapproved external service or global dependency;
- explicit real-world validation status;
- user approval of the final scientific boundary.
