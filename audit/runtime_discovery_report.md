# Runtime Discovery Report

Audit time: 2026-07-28T10:12:58+08:00

| Capability | Result | Detail |
|---|---|---|
| Windows PowerShell | available | 5.1 host |
| PowerPoint COM | available | Microsoft PowerPoint 16.0 |
| LibreOffice | available | `soffice.com` discovered on PATH |
| Poppler | available | `pdftoppm` available in bundled runtime |
| Bundled Node.js | available | 24.14.0 |
| Bundled Python | available | 3.12.13 |
| `@oai/artifact-tool` | available | 2.8.33; proprietary, internal evaluation/testing only |
| PptxGenJS | available | 4.0.1; MIT |
| Sharp | available | 0.34.5; Apache-2.0 |
| pypdf | available | 6.10.0 |
| python-docx | available | 1.2.0; MIT |
| openpyxl | available | 3.1.5; MIT |
| Pillow | available | 12.2.0 |
| pandas | available | 3.0.1; BSD-3-Clause |
| lxml | available | 6.0.2; BSD-3-Clause |
| PyYAML | unavailable | configuration files use JSON-compatible YAML |

## Renderer priority

1. Explicit CLI configuration.
2. Environment variables.
3. PowerPoint COM auto-detection.
4. LibreOffice with an isolated temporary user profile and timeout.
5. Explicit `RENDER_UNAVAILABLE` report.

PowerPoint COM is the primary renderer in this environment. LibreOffice and
Poppler provide independent fallbacks.

## Font discovery

Available safe fallbacks include Microsoft YaHei, SimHei, Arial, and Calibri.
Font files are not copied, embedded, packaged, or redistributed.

## Dependency and licensing boundary

The installed Artifact Tool license permits internal evaluation and testing
only. The maintainable project backend therefore remains PptxGenJS 4.0.1 (MIT).
Artifact Tool may be selected explicitly for controlled internal validation in
this Codex environment, but is not declared as a distributable production
dependency.
