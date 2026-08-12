# Visual backend regression suite

This runner discovers exactly four isolated synthetic projects. It never treats
the parent suite as one presentation project, and excludes
`expected_ground_truth/` from discovery, manifests, narrative construction,
Visual IR, claims, and initial QA.

Each project and arm receives an independent staging directory, source manifest,
Deck IR, Visual IR, claim-source map, output directory, and QA report.

Project names select only isolation and output paths. Visual routing is based on
CSV field contracts, content types, data dimensionality, and protocol status.
