# v2 Attribution and Clean-room Manifest

Audit date: 2026-07-28

## Current implementation statement

No upstream code, prompt, schema bundle, binary asset, example deck, or Skill
package was copied into the canonical workflow during this round. Selected
capabilities are implemented independently from the user-approved requirements
and repository-owned tests.

The six upstream checkouts remain isolated under the Git-ignored
`audit/upstream_review/sources/` directory. No upstream setup, installer,
generator, bridge, service, updater, API client, or Skill was executed during
Phase 1.

## Audited sources

| Project | Commit | License boundary | Current use |
|---|---|---|---|
| hugohe3/ppt-master | `a7ee83c75f60016c8a09f257d06424fed49b1c00` | MIT | behavior benchmark and clean-room requirements |
| kdnsna/ultimate-ppt-master-skill | `1749ad4587e78972493b1978816092210215072a` | MIT plus composite third-party notices | behavior benchmark and clean-room requirements |
| Gabberflast/academic-pptx-skill | `9f2b703ffe8d1449851617665ab1ffb3516d54ac` | root MIT conflicts with Skill proprietary metadata | academic narrative benchmark only |
| Noi1r/powerpoint-skill | `a39cd8cfba332741a96d52bebd2bb378b638364e` | MIT; inspected notice does not name the holder | isolated formula/diagram/OOXML behavior benchmark |
| MiniMax-AI/skills `pptx-generator` | `60aaae52bb2af8162732751a4332f62a5fef518b` | MIT | PptxGenJS design-pattern benchmark only |
| anthropics/skills `pptx` | `b29e7cf65e5cb78a5ac33d582270551bc74a14eb` | proprietary source-available | behavior reference only; copying and derivatives rejected |

If direct reuse is proposed later, it requires a new file-level adoption
decision, license verification, exact provenance, dependency review, and the
applicable copyright and license notice. This manifest is an engineering
record, not legal advice.
