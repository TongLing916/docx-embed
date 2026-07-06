# EDP 洞察

这是 README 结论背后的证据：**7 类解析器 × 14 种方法变体**，用不依赖 LLM 的规则评测衡量 **23 个维度**。

> English: [INSIGHTS.md](INSIGHTS.md)

## TL DR

- 独立解析器仍会丢 DOCX 嵌入附件：所有 `pure-*` 的 `embedded_object_recall = 0.0000`。
- EDP 管线能回收这些附件：所有 `pipeline-*` 的 `embedded_object_recall = 1.0000`。
- EDP 还会把 Word chart 和 SmartArt 文本抽成浅预览资源，所以 pipeline 行的 `chart_text_recall` 与 `smartart_text_recall` 都是 `1.0000`。
- 主正文质量仍取决于主解析器：Pandoc 在本 fixture 上总体最高，因为只有它恢复了抽样嵌套表。
- `ragflow` 现在指 mammoth→markdownify 的 Markdown export 路径。旧的 naive RAGFlow-style body-order parser 不再是公开方法。

## 记分板

来源：`evaluation/results_summary.tsv`。`doc_score` 是全部适用维度的等权均值。`mineru-*` 合并展示三个 MinerU backend，因为它们在该 fixture 上得分相同。

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

已实现但当前 fixture 未填充：`footnote_recall`、`footnote_anchor_accuracy`、`hyperlink_recall`、`toc_recall`、`textbox_recall`、`numbering_recall`、`table_numbering_recall`、`formula_recall`、`table_cell_match`、`asset_anchor_accuracy`、`embedded_content_hit`、`markdown_portability`。

## 关键结论

- **嵌入对象：** EDP 是决定性差异。任何可能含 Excel、PDF、DOCX、PPTX、Ole10Native 的 DOCX 都应使用 `pipeline-*`。
- **图表与 SmartArt：** EDP 会抽取 OOXML chart/diagram 资源到 `structured/resources/chart_*` 与 `structured/resources/diagram_*`，并从 `content.md` 链接预览。
- **嵌套表：** 本 fixture 中只有 Pandoc 还原了抽样嵌套表结构。
- **长表：** 另一个 326 行跨页表个案显示 Docling 和 Pandoc 的抽样单元格文本最好；MinerU 会漏掉该表正文。
- **RAG 使用：** 高结构保真不等于 embedding 友好。做向量库时，应在下游 chunking 步骤清理 Markdown/HTML/LaTeX 噪声。

## 选择建议

| 需求 | 推荐 |
|---|---|
| 回收附件 | 任意 EDP `pipeline-*` |
| 本地默认 | `pipeline-markitdown` |
| 本 fixture 综合最高 | `pipeline-pandoc` |
| 表格密集文档 | `pipeline-docling`，用 Pandoc 对照 |
| 嵌套表 | 本 fixture 上为 `pipeline-pandoc` |
| RAG 取向的干净 Markdown | `pipeline-ragflow` 可用，但需在你的文档上检查公式、链接和编号 |

## 复现

```bash
scripts/batch_run_all.sh
uv run python evaluation/run_eval.py
```

可直接阅读的 Markdown 样例在 [`docs/examples/converted-markdown/`](docs/examples/converted-markdown/)。

详细资料：[RAGFlow](docs/ragflow-docx-parsing-research.md)、[Dify](docs/dify-docx-parsing-research.md)、[WeKnora](docs/weknora-docx-parsing-research.md)、[MinerU](docs/mineru-backend-evaluation.md)、[评测报告](docs/docx-markdown-evaluation-report.md)。
