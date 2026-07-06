# DOCX to Markdown Evaluation Report

This report summarizes the current public evaluation shape. For the concise reader-facing conclusions, start with [`../INSIGHTS.md`](../INSIGHTS.md).

## Current Scope

- Methods: 14 rows, covering pipeline and pure variants for MarkItDown, MinerU pipeline/VLM/hybrid, Docling, Pandoc, and RAGFlow.
- Dimensions: 23 rule-based dimensions in `evaluation/layer_fidelity.py`.
- Primary fixture: `data/llm-social-impact-fixture.docx`.
- Output artifacts: `evaluation/results.tsv`, `evaluation/results_summary.tsv`, `evaluation/misses.tsv`, and `evaluation/misses_report.md`.

`ragflow` now means the Markdown export path. The previous naive RAGFlow-style body-order parser is not a public evaluation method.

## Key Dimensions

| Dimension | Purpose |
|---|---|
| `embedded_object_recall` | Embedded DOCX/OLE attachments registered in the manifest. |
| `image_recall` | Images referenced inline or registered in the manifest. |
| `table_recall` / `table_cell_match` | Table shape and sampled cell text recovery. |
| `nested_table_recall` | Nested table shape plus sampled text. |
| `nested_table_asset_recall` | Image/SmartArt evidence inside nested tables. |
| `checkbox_recall` | Visible checkbox labels/symbols. |
| `chart_text_recall` | Text from Word chart OOXML or chart previews. |
| `smartart_text_recall` | Text from SmartArt/diagram OOXML or previews. |
| `heading_recall` / `heading_tree_score` | Heading text and relative hierarchy. |
| `hyperlink_recall` | URL target survival, not just anchor text. |
| `toc_recall` | TOC entry text inside the TOC region. |
| `numbering_recall` / `table_numbering_recall` | Word auto-numbering in lists and table rows. |
| `formula_recall` | Token-level OMML/formula content recall. |

## Current Findings

- All `pure-*` methods drop embedded attachments.
- All `pipeline-*` methods recover embedded attachments.
- EDP pipeline methods recover chart and SmartArt text through shallow resource previews.
- Pandoc is strongest overall on the current fixture and uniquely recovers the sampled table-row numbering sequence.
- Docling remains the safer first comparison for long table cell text.
- Formula, hyperlink, TOC, numbering, and portability scorers are implemented but not currently populated by the public fixture's ground truth.

## Reproduction

```bash
scripts/batch_run_all.sh
uv run python evaluation/run_eval.py
```

The batch outputs are intentionally ignored by git. Curated Markdown examples are committed under [`examples/converted-markdown`](examples/converted-markdown/).
