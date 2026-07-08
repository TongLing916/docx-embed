# Embedded Document Parser (EDP)

A pre-processing layer that recovers embedded OLE attachments (XLSX / PDF / DOCX / PPT) from DOCX **before** your main parser runs — then reattaches them as inline links in a clean, auditable document package.

> 中文版：[README.zh-CN.md](README.zh-CN.md)

---

## Why EDP exists

Real-world DOCX files carry embedded attachments inside `word/embeddings/` — an Excel model, a PDF brief, a nested DOCX. Every mainstream parser we tested (MarkItDown, MinerU, Docling, Pandoc, RAGFlow, Dify, WeKnora) **silently drops them**. The attachment vanishes from the output with no error, no warning, no link.

EDP closes exactly that gap. It is a **resource layer**, not a competing parser: it extracts embedded objects, registers them in a manifest (hash / mime / anchor / parse_status), generates shallow previews, and reattaches them at their original anchor positions. The main parser (MarkItDown / MinerU / Docling / Pandoc / RAGFlow) is a **swappable component** that only parses the cleaned main body. EDP never rewrites the main parser's markdown.

---

## Headline: every pure framework drops embeds

On our test corpus (`llm-social-impact-fixture.docx`, 5 embedded attachments: txt / csv / xlsx / pdf / docx):

| Tier | `embedded_object_recall` |
|---|:---:|
| All standalone frameworks (`pure-*`) | **0.0000** |
| All EDP-enhanced pipelines (`pipeline-*`) | **1.0000** |

Embedded attachments are the isolated, decisive gap EDP closes. Newer dimensions also track checkboxes, chart text, SmartArt text, nested tables, nested-table assets, formulas, numbering, hyperlinks, TOC, and portability. Full 14-method × 23-dimension scoreboard: [INSIGHTS.md](INSIGHTS.md).

---

## What EDP does

- **Extracts** embedded OLE objects from DOCX via `python-docx` + `olefile` / `oleobj`, including orphan files in `word/embeddings/` not referenced by any OLE relationship.
- **Registers** each resource in `embedded_resources.jsonl` with `sha256` / mime / anchor / `parse_status` (`shallow_preview` / `extracted_only` / `preserved_image`).
- **Previews** allowlisted types: `txt` / `csv` → fenced text; `xlsx` → sheet-by-sheet markdown table + typed `preview.json`; Word charts / SmartArt → shallow OOXML text/data previews. PDF / DOCX / PPTX / unknown binaries are preserved and linked, not previewed.
- **Reattaches** each resource as an inline markdown link at its original anchor position in `content.md`.
- **Emits** a self-contained, auditable package: `{doc_id}/{raw,structured,manifest.json}`.

---

## Quick start

EDP targets Windows, macOS, and Linux on Python 3.12. The core library and
`markitdown` / `ragflow` local parser paths are pure Python. Optional parser
backends keep their own platform requirements: `pandoc` must be installed on
`PATH`, and MinerU / Docling require reachable HTTP services.

Before running, check that inputs are modern `.docx` files. Legacy `.doc` files are not supported directly; convert them to `.docx` first. Avoid writing output back into the source document folder, and skip Office lock files such as `~$report.docx`.

```bash
# Convert one DOCX file with the default pipeline
uv run edp pipeline input.docx output_package

# Switch the main parser
uv run edp pipeline input.docx output_package --main-parser pandoc

# Allow EDP to unwrap Ole10Native .bin shells into the real payload
uv run edp pipeline input.docx output_package --unsafe-unwrap-embedded

# Pure-parser control: no attachment extraction, no manifest — for comparison only
uv run edp parser pandoc input.docx output_pandoc
```

Convert every `.docx` file under a folder, including nested directories, while preserving the relative output layout:

```bash
uv run edp batch input_docs output/packages --main-parser markitdown --unsafe-unwrap-embedded
```

The older `examples/run_pipeline.py` and `examples/run_parser.py` entrypoints
remain available for compatibility.

Use `--unsafe-unwrap-embedded` only when you want Ole10Native `.bin` wrappers unpacked into their real payloads. For public datasets, review generated `manifest.json` files before publishing.

`--main-parser` choices: `markitdown` (default), `mineru-pipeline`, `mineru-vlm-engine`, `mineru-hybrid-engine`, `docling`, `pandoc`, `ragflow` (`mineru` is a compat alias for `mineru-pipeline`). `ragflow` uses the Markdown export path.

### Environment variables

| Parser | Required env | Notes |
|---|---|---|
| `markitdown` | none | Local, no API key. Strong default. |
| `pandoc` | none | Requires system `pandoc` binary on PATH. |
| `ragflow` | none | Local mammoth→markdownify export-style parser; no RAGFlow service. |
| `mineru-pipeline` / `mineru-vlm-engine` / `mineru-hybrid-engine` | `MINERU_API_KEY`, `MINERU_BASE_URL` | Black-box HTTP service. Compat aliases: `MINERU_AUTHORIZATION`, `DOCUMENT_CONVERTER_AUTHORIZATION`, `MINERU_API_ENDPOINT`, `MINERU_REQUEST_TIMEOUT`, `MINERU_BACKEND`. |
| `docling` | `DOCLING_AUTHORIZATION`, `DOCLING_API_BASE` | Local Docling Serve (`http://127.0.0.1:5001` by default). Compat aliases: `DOCLING_API_ENDPOINT`, `DOCLING_REQUEST_TIMEOUT`. |

---

## Choose a parser

No single parser wins every dimension. Pick by what your downstream needs.

### By need

| Your need | Recommended | Why |
|---|---|---|
| Recover embedded attachments | EDP `pipeline-*` (any) | Only EDP = `1.0`; all `pure-*` = `0.0` |
| Highest overall fidelity | `pipeline-pandoc` | `doc_score` 0.9790; strongest overall on this fixture |
| Long / wide / cross-page tables | `pipeline-docling` | 326-row dual-column case: `table_cell_match` 1.0000 |
| Word chart labels | any EDP `pipeline-*` | EDP extracts chart OOXML into `structured/resources/chart_*` |
| SmartArt text | any EDP `pipeline-*` | EDP extracts diagram OOXML into `structured/resources/diagram_*` |
| Nested tables | `pipeline-pandoc` on this fixture | only Pandoc restores the sampled nested table structure |
| RAG chunking (clean semantic text) | `ragflow` / `WeKnora` / `Dify` / `mineru` | verify formulas, links, and numbering on your own documents |

### By content

| Your DOCX content | Recommended |
|---|---|
| Simple text + ordinary images | `pipeline-markitdown` or another local parser is usually enough |
| Embedded Excel / PDF / DOCX attachments | **must** use EDP `pipeline-*` |
| Long / wide / cross-page tables | `pipeline-docling` |
| Word charts / SmartArt / nested tables | charts + SmartArt: any EDP `pipeline-*`; sampled nested table: `pipeline-pandoc` |
| Formula / TOC / numbering heavy | run the evaluation on your own fixture; these dimensions are implemented but not populated in the current public GT |

**Default recommendation:** `pipeline-markitdown` (local, no API key, strong all-round). Upgrade to `pipeline-pandoc` when fidelity / numbering / table-row numbering matter and `pandoc` is installed. Use `pipeline-docling` for table-dense docs. **Never ship `pure-*` for attachment-bearing DOCX** — the attachments will be gone.

Full decision matrix and per-framework trade-offs: [INSIGHTS.md](INSIGHTS.md).

---

## Document package output

```
{doc_id}/
├── raw/                  # original.docx + extracted embedded/ binaries
├── structured/           # content.md, child_files.md, position_map.csv,
│                         # embedded_resources.jsonl, assets/, resources/
└── manifest.json         # package manifest (parse_status, content_map, …)
```

`content.md` stays clean — only inline markdown links at original anchor positions. Full package tree and data structures: [docs/DESIGN.md](docs/DESIGN.md).

---

## Status

**V0.2 — evaluation system shipped.** 14-method × 23-dimension scoreboard covering MarkItDown / MinerU / Docling / Pandoc / RAGFlow (pipeline-enhanced vs pure). Chart and SmartArt OOXML text/data previews are already included in the DOCX resource layer. See [INSIGHTS.md](INSIGHTS.md).

**Next:** image understanding for DOCX assets: image metadata, optional OCR/VLM descriptions, classification, and caption anchoring. PDF / PPTX embedded files remain preserved as linked resources, not expanded into a separate parser roadmap.

---

## Documentation

- **[INSIGHTS.md](INSIGHTS.md)** — concise 14×23 scoreboard, TL DR, and parser selection guide.
- **[docs/DESIGN.md](docs/DESIGN.md)** — v1.0 technical design (pipeline, package structure, data structures, evaluation system, module layout, roadmap).
- **[docs/examples/converted-markdown/](docs/examples/converted-markdown/)** — curated converted Markdown files for direct reading on GitHub.
- **Per-framework research notes** (Chinese): [RAGFlow](docs/ragflow-docx-parsing-research.md) · [Dify](docs/dify-docx-parsing-research.md) · [WeKnora](docs/weknora-docx-parsing-research.md) · [MinerU backend](docs/mineru-backend-evaluation.md) · [evaluation report](docs/docx-markdown-evaluation-report.md)

---

## License

MIT — commercial use permitted. See [LICENSE](LICENSE).

中文版：[README.zh-CN.md](README.zh-CN.md) | [INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)
