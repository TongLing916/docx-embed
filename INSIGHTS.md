# EDP Insights

Evidence behind the README claims: DOCX parser fidelity across **7 parser families × 14 method variants**, measured with a rule-based evaluation over **23 dimensions**.

> 中文版：[INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)

## TL DR

- Standalone parsers still drop embedded DOCX attachments: every `pure-*` method scores `embedded_object_recall = 0.0000`.
- Every EDP pipeline recovers those attachments: every `pipeline-*` method scores `embedded_object_recall = 1.0000`.
- EDP also surfaces Word chart and SmartArt text as shallow resource previews; that is why pipeline rows score `chart_text_recall = 1.0000` and `smartart_text_recall = 1.0000`.
- Parser choice still matters for the main body: Pandoc is strongest overall on this fixture because it is the only method here that restores the sampled nested table.
- `ragflow` now means the mammoth→markdownify Markdown export path. The older naive RAGFlow-style body-order parser is no longer a public method.

## Scoreboard

Source: `evaluation/results_summary.tsv`. `doc_score` is the unweighted mean over applicable dimensions. `mineru-*` groups the three MinerU backends because they score identically on this fixture.

| Method | doc_score | embed | image | table | nested_tbl | nested_asset | heading | key_text | checkbox | chart | smartart |
|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| `pipeline-pandoc` | **0.9790** | 1.00 | 1.00 | 0.77 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 |
| `pipeline-markitdown` | 0.9091 | 1.00 | 1.00 | 1.00 | 0.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 |
| `pipeline-docling` | 0.9091 | 1.00 | 1.00 | 1.00 | 0.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 |
| `pipeline-ragflow` | 0.9091 | 1.00 | 1.00 | 1.00 | 0.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 |
| `pipeline-mineru-*` | 0.9021 | 1.00 | 1.00 | 0.92 | 0.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 | 1.00 |
| `pure-pandoc` | 0.6608 | 0.00 | 1.00 | 0.77 | 1.00 | 0.50 | 1.00 | 1.00 | 1.00 | 0.00 | 0.00 |
| `pure-mineru-*` | 0.6180 | 0.00 | 1.00 | 0.92 | 0.00 | 0.50 | 1.00 | 1.00 | 1.00 | 0.38 | 0.00 |
| `pure-markitdown` | 0.5909 | 0.00 | 1.00 | 1.00 | 0.00 | 0.50 | 1.00 | 1.00 | 1.00 | 0.00 | 0.00 |
| `pure-docling` | 0.5909 | 0.00 | 1.00 | 1.00 | 0.00 | 0.50 | 1.00 | 1.00 | 1.00 | 0.00 | 0.00 |
| `pure-ragflow` | 0.5818 | 0.00 | 0.90 | 1.00 | 0.00 | 0.50 | 1.00 | 1.00 | 1.00 | 0.00 | 0.00 |

Implemented but not currently populated by this fixture: `footnote_recall`, `footnote_anchor_accuracy`, `hyperlink_recall`, `toc_recall`, `textbox_recall`, `numbering_recall`, `table_numbering_recall`, `formula_recall`, `table_cell_match`, `asset_anchor_accuracy`, `embedded_content_hit`, `markdown_portability`.

## Key Takeaways

- **Embedded objects:** EDP is the differentiator. Use `pipeline-*` for any DOCX that may contain embedded Excel, PDF, DOCX, PPTX, or Ole10Native payloads.
- **Charts and SmartArt:** EDP extracts OOXML chart/diagram resources into `structured/resources/chart_*` and `structured/resources/diagram_*`, then links previews from `content.md`.
- **Nested tables:** Pandoc is the only method in this fixture that restores the sampled nested table structure.
- **Long tables:** A separate 326-row cross-page table case shows Docling and Pandoc recover sampled cell text best; MinerU drops the body of that table.
- **RAG use:** High structural fidelity is not always embedding-friendly. Strip noisy Markdown/HTML/LaTeX in a downstream chunking step when building vector indexes.

## Selection Guide

| Need | Recommended |
|---|---|
| Attachment recovery | Any EDP `pipeline-*` |
| Local default | `pipeline-markitdown` |
| Best overall on this fixture | `pipeline-pandoc` |
| Table-dense documents | `pipeline-docling`, with Pandoc as comparison |
| Nested tables | `pipeline-pandoc` on this fixture |
| RAG-oriented clean Markdown | `pipeline-ragflow` can be useful, but verify formulas, links, and numbering on your own documents |

## Reproduce

```bash
scripts/batch_run_all.sh
uv run python evaluation/run_eval.py
```

Curated Markdown outputs are available in [`docs/examples/converted-markdown/`](docs/examples/converted-markdown/).

Detailed research notes: [RAGFlow](docs/ragflow-docx-parsing-research.md), [Dify](docs/dify-docx-parsing-research.md), [WeKnora](docs/weknora-docx-parsing-research.md), [MinerU](docs/mineru-backend-evaluation.md), and [evaluation report](docs/docx-markdown-evaluation-report.md).
