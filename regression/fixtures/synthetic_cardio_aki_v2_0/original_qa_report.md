# QA Report

- Final status: `PARTIALLY_READY`
- Renderer: PowerPoint COM
- PPTX slides: 18
- PDF pages: 18
- Preview pages: 18

## Scientific QA

- PASS: all registered claims map to known input sources; no automated claim completion was used

## Visual QA

- FAIL: Slide 2: text may overflow vertically (60.0pt content in 30.2pt box)
- FAIL: Slide 17: text may overflow vertically (60.0pt content in 30.2pt box)
- FAIL: Slide 18: text may overflow vertically (36.0pt content in 30.2pt box)

## File QA

- PASS: PPTX ZIP, PDF page count, preview count, and immutable-input hash checks passed

## OOXML QA

- WARN: EMBEDDED_OR_OLE_CONTENT_PRESENT: Embedded/OLE content is present and was not opened.

## PowerPoint Text Geometry QA

- FAIL: Slide 2: text may overflow vertically (60.0pt content in 30.2pt box)
- FAIL: Slide 17: text may overflow vertically (60.0pt content in 30.2pt box)
- FAIL: Slide 18: text may overflow vertically (36.0pt content in 30.2pt box)

## Boundary

Automated QA does not constitute final scientific approval. Manual review remains mandatory.