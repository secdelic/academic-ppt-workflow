# Synthetic CI contract fixtures

All files in this directory are minimal, deterministic test contracts.

```text
fixture_provenance: SYNTHETIC
```

They contain no patient data, real research results, real names, private paths,
historical run directories, or authorized templates.  They replace ignored
developer-workspace evidence that is unavailable in a clean Git checkout.

| Fixture | Contract under test |
| --- | --- |
| `source_isolation.json` | Four project source registries remain disjoint and do not carry hidden answer material. |
| `style_cache_status.json` | Projects without a reference style report no style cache reuse. |
| `powerpoint_render_index.json` | A 16-deck structural result contract records successful open and equal slide/preview counts. |
| `libreoffice_render_index.json` | A 16-deck cross-render result contract records successful exit and non-zero page counts. |
| `input_hash_integrity.csv` | Synthetic registered inputs retain the same before/after hash. |

These fixtures validate parsing and gate semantics only. They are not evidence
that PowerPoint or LibreOffice ran in the current environment; executable smoke
and application validation remain separate runtime gates.
