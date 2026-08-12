# Academic PPT Workflow v2.4 presentation-design regression tests

- Executed: 2026-08-11 (Asia/Shanghai)
- Scope: `tests/presentation_design/`
- Runner: `python -m unittest discover -s tests/presentation_design -p "test_*.py" -v`
- Result: **36/36 PASS**
- Runtime: 0.026 s (final v2.4 candidate)
- Full repository regression: **133/133 PASS** in 10.418 s

The suite uses project-neutral synthetic objects and does not depend on a fixed
project name, slide number, title, or page count. It calls the canonical v2.4
model/building interfaces and reads the versioned design-system contracts.

Covered contracts:

1. slide-specific and field-level source bindings;
2. zh-CN language compliance;
3. title/entity scope, facet, and sign preservation;
4. independent networks are not sequences;
5. expected/rendered row equality and declared omission handling;
6. brief `must_include` and slide-budget preservation;
7. true risk-of-bias matrix, risk matrix, and DAG semantics;
8. forest axes/reference and prediction-interval labeling;
9. native chart data bindings;
10. typography, layout diversity, card ratio, and generic-title thresholds;
11. duplicate slide semantics and source-figure callouts;
12. protocol no-results boundary.

Latest contract-strengthening coverage:

13. pre-render `rendered_row_count=None` remains unknown and cannot be counted
    as a render PASS;
14. `source_row_count` is kept distinct from the expected displayed visual-row
    count;
15. appendix conflict and unresolved items contribute non-empty source
    bindings;
16. source and claim maps list only the `visual_id` values matching the exact
    source/field binding;
17. the renderer emits independent `renderEvidence` and
    `created_object_count`, without the former payload-only
    `actualRenderedRowCount` path;
18. grouped-composition group names are data-driven rather than hard-coded;
19. DAG rendering and validation accept arbitrary node identifiers.
20. `data_contract_id` changes when fields, row filters, aggregation,
    entity scope, or outcome semantics change;
21. claim/source visual mapping requires the complete expected-field set to be
    a subset of the visual binding, not merely an arbitrary field overlap;
22. English `High` and Chinese `高` risk severity both use the warning token;
23. protocol assessment timelines sort by time and allocate adjacent labels to
    collision-safe lanes, failing explicitly when the lane budget is exceeded;
24. faceted volcano plots retain all points while limiting labels to the
    declared top-two-per-facet budget with collision adjustment.
25. canonical source-category values are registered from input rather than
    fabricated by the renderer.

No input data, global environment, or external dependency was modified by the
test run.
