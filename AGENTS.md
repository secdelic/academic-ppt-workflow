# Academic PPT Workflow - Mandatory Rules

1. Never invent scientific facts, results, numbers, citations, methods, authors,
   ethics approvals, software versions, or mechanism conclusions.
2. Treat every file under `input/` as read-only. Never overwrite an existing
   formal PPTX/PDF or delete user files.
3. Every claim, number, chart, figure, and citation must trace to a registered
   input source. Missing information is `INFORMATION_REQUIRED`.
4. Put intermediates only in `staging/`; put formal results only in a new
   `output/<project_name>/<timestamp>/` directory.
5. Use a consistent grid, readable typography, sufficient contrast, one key
   message per slide, and no text smaller than configured minimums.
6. Render every slide and run scientific, visual, and file QA before delivery.
7. Never report completion while a blocking issue remains; use only
   `READY_FOR_ASSISTED_USE`, `PARTIALLY_READY`, or `BLOCKED`.
8. The user retains final scientific approval and must complete manual review.
