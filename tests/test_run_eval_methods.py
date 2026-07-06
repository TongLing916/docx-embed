from __future__ import annotations

from pathlib import Path

from evaluation.layer_fidelity import DIMENSIONS
from evaluation.run_eval import METHODS


def test_evaluation_methods_include_mineru_backend_variants() -> None:
    assert "pipeline-mineru-pipeline" in METHODS
    assert "pipeline-mineru-vlm-engine" in METHODS
    assert "pipeline-mineru-hybrid-engine" in METHODS
    assert "pure-mineru-pipeline" in METHODS
    assert "pure-mineru-vlm-engine" in METHODS
    assert "pure-mineru-hybrid-engine" in METHODS


def test_evaluation_dimensions_include_extended_quality_metrics() -> None:
    for dimension in (
        "footnote_recall",
        "footnote_anchor_accuracy",
        "heading_tree_score",
        "table_cell_match",
        "asset_anchor_accuracy",
        "embedded_content_hit",
        "markdown_portability",
        "nested_table_recall",
        "checkbox_recall",
        "chart_text_recall",
        "smartart_text_recall",
        "nested_table_asset_recall",
    ):
        assert dimension in DIMENSIONS


def test_ragflow_markdown_is_folded_into_ragflow_methods() -> None:
    assert "pipeline-ragflow" in METHODS
    assert "pure-ragflow" in METHODS
    assert "pipeline-ragflow-markdown" not in METHODS
    assert "pure-ragflow-markdown" not in METHODS


def test_batch_eval_pipeline_methods_use_unsafe_unwrap() -> None:
    script = Path("scripts/batch_run_all.sh").read_text(encoding="utf-8")
    for line in script.splitlines():
        if "uv run examples/run_pipeline.py" in line:
            assert "--unsafe-unwrap-embedded" in line
