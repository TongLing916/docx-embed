# Converted Markdown Samples

These files are curated outputs from `llm-social-impact-fixture.docx` so readers can inspect parser behavior directly on GitHub without running the full batch pipeline.

## Files

| File | Source method | Notes |
|---|---|---|
| `pipeline-markitdown.md` | EDP + MarkItDown | Default local pipeline. |
| `pipeline-mineru-pipeline.md` | EDP + MinerU pipeline backend | Cloud/service-parser comparison with EDP resource recovery. |
| `pipeline-pandoc.md` | EDP + Pandoc | Strong numbering and formula baseline. |
| `pipeline-docling.md` | EDP + Docling | Useful table-dense comparison. |
| `pipeline-ragflow.md` | EDP + RAGFlow markdown export | Current `ragflow` behavior. |

Selected resource previews live under `resources/`: one XLSX preview, one Word chart preview, and two SmartArt previews. Full reproducible packages are intentionally not committed; regenerate them with:

```bash
scripts/batch_run_all.sh
uv run python evaluation/run_eval.py
```
