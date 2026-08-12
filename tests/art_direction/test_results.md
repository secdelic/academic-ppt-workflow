# Academic PPT Workflow v2.5 art-direction contract tests

- Executed: 2026-08-11 (Asia/Shanghai)
- Production scope: `scripts/v2_5/`
- Contract scope: `tests/art_direction/`
- Art-direction result: **44/44 PASS; 0 skip, 0 failure, 0 error**
- Full workflow regression result: **177/177 PASS; 0 failure, 0 error**
- Full-suite runtime: **13.735 s**
- Evidence: `audit/v2_5/runs/20260811_200637_v25_final/regression_test_raw.log`

The connected tests exercise the production `ArtDirectionSpec`,
`DeckRhythmPlan`, layout variants, hero geometry, annotation limits, hybrid
render routing, native-chart privacy, figure-rebuild policy, frozen scientific
identity, Unicode/CJK line breaking, source-bound semantic overlays, graph and
network topology, and the short-lived offline batch renderer. No contract was
treated as an expected failure or precondition skip.

All fixtures are project-neutral and use semantic slide IDs. No test depends
on a fixed project name, slide number, title, or an existing generated deck.
