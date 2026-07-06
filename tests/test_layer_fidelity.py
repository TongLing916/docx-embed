from __future__ import annotations

import importlib.util
import json
import shutil
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from evaluation.layer_fidelity import (  # noqa: E402
    DIMENSIONS,
    _parse_markdown_tables,
    score,
)


def _load_build_groundtruth() -> callable:
    spec = importlib.util.spec_from_file_location(
        "build_groundtruth", REPO_ROOT / "scripts" / "build_groundtruth.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.build_groundtruth


build_groundtruth = _load_build_groundtruth()


def _write_package(
    tmp_path: Path,
    content_md: str,
    *,
    embedded_objects: list[dict] | None = None,
    images: list[dict] | None = None,
) -> Path:
    package = tmp_path / "pkg"
    if package.exists():
        shutil.rmtree(package)
    (package / "structured").mkdir(parents=True)
    (package / "structured" / "content.md").write_text(content_md, encoding="utf-8")
    manifest = {"content_map": {"embedded_objects": embedded_objects or [], "images": images or []}}
    (package / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    return package


def test_parse_pipe_table():
    content = "| A | B |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |\n"
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    assert rows == 3
    assert cols == 2
    assert cells[1] == ["1", "2"]


def test_parse_pandoc_grid_table():
    # Pandoc emits grid tables whose rows are separated by ``+---+`` borders
    # (``+===+`` under the header). The parser must keep them as one table,
    # not split into one-row fragments, and drop border rows from the count.
    content = (
        "| A | B |\n"
        "+===+===+\n"
        "| 1 | 2 |\n"
        "+---+---+\n"
        "| 3 | 4 |\n"
        "+---+---+\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    assert rows == 3
    assert cols == 2
    assert cells[2] == ["3", "4"]


def test_parse_pandoc_grid_table_multiline_cells():
    # Pandoc grid tables expand merged cells across multiple physical ``|``
    # lines, with partial borders (``|  |  +---``) marking cell boundaries.
    # The logical row count must follow the border count, not the physical
    # ``|``-line count, or merged-cell tables over-count rows ~2x.
    content = (
        "| 一级指标 | 二级指标 | 指标项 |\n"
        "+----------+----------+--------+\n"
        "| A.1      | A.1.1    | A.1.1.1|\n"
        "| 基础指标 |          |        |\n"
        "|          |          +--------+\n"
        "|          |          | A.1.1.2|\n"
        "|          +----------+--------+\n"
        "|          | A.1.2    | A.1.2.1|\n"
        "+----------+----------+--------+\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    # 1 header + 3 logical data rows = 4, NOT the 6 physical ``|`` lines.
    assert rows == 4
    assert cols == 3


def test_parse_pandoc_simple_table_multiline_rows():
    # Pandoc simple/multiline tables use dashed rule lines (no ``|``/``+``):
    # a full-width top border, a header, a gapped column-rule (column count),
    # blank-separated data rows, and a closing border. The eval must recognize
    # this format -- otherwise pandoc's simple tables are invisible and every
    # one is a false miss.
    content = (
        "  ---------------------------------------------------------------\n"
        "  序号        名称             A型车      B型车\n"
        "  ------ ---------------------------- ------------ --------------\n"
        "  1     车体长度（mm）        21880      19000\n"
        "\n"
        "  2     车辆长度（mm）        22800      19520\n"
        "  ---------------------------------------------------------------\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    # 1 header + 2 data rows = 3, NOT 4 physical content lines.
    assert rows == 3
    assert cols == 4


def test_parse_pandoc_simple_table_no_blanks():
    # A simple table whose data rows are consecutive single lines (no blank
    # separators) must count each line as a row.
    content = (
        "  Name   Value\n"
        "  ------ ------\n"
        "  a      1\n"
        "  b      2\n"
        "  c      3\n"
        "  ------ ------\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), _ = tables[0]
    assert rows == 4  # 1 header + 3 data
    assert cols == 2


def test_parse_pandoc_multiline_table_wrapped_cell():
    # A multiline row may span several physical lines; the blank line is the
    # true row separator, so two data rows even though one wraps.
    content = (
        "  -------------------------------------------\n"
        "  Header1     Header2\n"
        "  ------      ------\n"
        "  row1 cell   wraps onto\n"
        "  a second line\n"
        "\n"
        "  row2 cell   short\n"
        "  -------------------------------------------\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), _ = tables[0]
    assert rows == 3  # 1 header + 2 data (row1 wraps but is one row)
    assert cols == 2


def test_simple_table_rule_not_confused_with_pipe_or_grid():
    # A horizontal rule / thematic break (``---``) must not be parsed as a
    # 1-column simple table, and a pipe table nearby must still parse.
    content = (
        "Some intro text.\n"
        "\n"
        "---\n"
        "\n"
        "| A | B |\n"
        "| --- | --- |\n"
        "| 1 | 2 |\n"
    )
    tables = _parse_markdown_tables(content)
    # Only the pipe table counts; the ``---`` thematic break is not a table.
    shapes = [shape for shape, _ in tables]
    assert ((2, 2)) in shapes
    assert (1, 1) not in shapes


def test_parse_html_table():
    content = "<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>\n"
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    assert rows == 2
    assert cols == 2
    assert cells[1] == ["1", "2"]


def test_parse_html_table_expands_colspan():
    # A colspan="2" header occupies two grid columns; raw <td> counting would
    # report 2 columns instead of 3.
    content = (
        "<table>"
        "<tr><td>一级</td><td colspan=\"2\">指标项</td></tr>"
        "<tr><td>x</td><td>1</td><td>2</td></tr>"
        "</table>\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    assert rows == 2
    assert cols == 3
    assert cells[0] == ["一级", "指标项", "指标项"]


def test_parse_html_table_expands_rowspan():
    # rowspan is vertical and does not affect column count; only colspan
    # drives the width. The grid is 3 columns wide (colspan="2" header).
    content = (
        "<table>"
        "<tr><td rowspan=\"2\">a</td><td colspan=\"2\">b</td></tr>"
        "<tr><td>1</td><td>2</td></tr>"
        "</table>\n"
    )
    tables = _parse_markdown_tables(content)
    assert len(tables) == 1
    (rows, cols), cells = tables[0]
    assert rows == 2
    assert cols == 3
    assert cells[0] == ["a", "b", "b"]


def test_table_recall_tolerates_one_column_drift(tmp_path: Path):
    # Merged cells commonly shift the parsed column count by one versus the
    # Word-grid ground truth; a strict shape match would miss the table.
    content = "<table><tr><td>a</td><td>b</td><td>c</td><td>d</td></tr></table>\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "images": [],
            "tables": [{"index": 0, "rows": 1, "cols": 5}],
            "headings": [],
            "embedded_objects": [],
            "key_texts": [],
        }
    }
    assert score(package, gt)["table_recall"] == 1.0


def test_footnote_recall_and_anchor_accuracy_are_separate(tmp_path: Path):
    content = "Intro [^1] sentence.\n\nLater text.\n\n[^1]: Foot note text\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "footnotes": [
                {
                    "id": "1",
                    "text": "Foot note text",
                    "anchor_before": "Intro",
                    "anchor_after": "sentence",
                }
            ]
        }
    }
    scored = score(package, gt)
    assert scored["footnote_recall"] == 1.0
    assert scored["footnote_anchor_accuracy"] == 1.0

    drifted = _write_package(
        tmp_path,
        "[^1]\n\nIntro sentence.\n\n[^1]: Foot note text\n",
    )
    drifted_score = score(drifted, gt)
    assert drifted_score["footnote_recall"] == 1.0
    assert drifted_score["footnote_anchor_accuracy"] == 0.0


def test_heading_tree_score_tolerates_uniform_level_shift(tmp_path: Path):
    # docling-style output: every Word heading shifted down by one (# -> ##).
    # The relative tree is correct, so a uniform offset should be tolerated.
    package = _write_package(tmp_path, "## Title\n\n### Section\n\nbody\n")
    gt = {
        "expected": {
            "headings": [
                {"level": 1, "text": "Title"},
                {"level": 2, "text": "Section"},
            ]
        }
    }

    scored = score(package, gt)

    assert scored["heading_recall"] == 1.0
    assert scored["heading_tree_score"] == 1.0


def test_heading_tree_score_penalizes_nonuniform_level_shift(tmp_path: Path):
    # A non-uniform shift (Section and Sub both at ###) is a real structural
    # error and must still be penalized.
    package = _write_package(tmp_path, "## Title\n\n### Section\n\n### Sub\n\nbody\n")
    gt = {
        "expected": {
            "headings": [
                {"level": 1, "text": "Title"},
                {"level": 2, "text": "Section"},
                {"level": 3, "text": "Sub"},
            ]
        }
    }

    scored = score(package, gt)

    assert scored["heading_recall"] == 1.0
    assert scored["heading_tree_score"] < 1.0


def test_heading_tree_score_strict_for_single_heading(tmp_path: Path):
    # A single heading gives no offset evidence, so comparison stays strict.
    package = _write_package(tmp_path, "## Title\n\nbody\n")
    gt = {
        "expected": {
            "headings": [{"level": 1, "text": "Title"}]
        }
    }

    scored = score(package, gt)

    assert scored["heading_recall"] == 1.0
    assert scored["heading_tree_score"] < 1.0


def test_table_cell_match_penalizes_wrong_cells_even_when_shape_matches(tmp_path: Path):
    content = "| h1 | h2 |\n| --- | --- |\n| ok | wrong |\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "tables": [{"index": 0, "rows": 2, "cols": 2}],
            "table_cell_checks": [
                {
                    "table_index": 0,
                    "label": "sample table",
                    "cells": [
                        {"row": 1, "col": 0, "text": "ok"},
                        {"row": 1, "col": 1, "text": "expected"},
                    ],
                }
            ],
        }
    }

    scored = score(package, gt)

    assert scored["table_recall"] == 1.0
    assert scored["table_cell_match"] == 0.5


def test_asset_anchor_accuracy_requires_assets_near_expected_context(tmp_path: Path):
    content = "Intro\n\n![](assets/images/image_01.png)\n\nOutro\n"
    package = _write_package(
        tmp_path,
        content,
        images=[{"ref": "image_01", "path": "structured/assets/images/image_01.png"}],
    )
    gt = {
        "expected": {
            "asset_anchors": [
                {"ref": "image_01", "kind": "image", "before": "Intro", "after": "Outro"}
            ]
        }
    }
    assert score(package, gt)["asset_anchor_accuracy"] == 1.0

    drifted = _write_package(
        tmp_path,
        "![](assets/images/image_01.png)\n\nIntro\n\nOutro\n",
        images=[{"ref": "image_01", "path": "structured/assets/images/image_01.png"}],
    )
    assert score(drifted, gt)["asset_anchor_accuracy"] == 0.0


def test_embedded_content_hit_searches_attachment_content_by_ref(tmp_path: Path):
    package = _write_package(tmp_path, "parent body only\n")
    attachment = package / "structured" / "attachments" / "attachment_01"
    attachment.mkdir(parents=True)
    (attachment / "content.md").write_text("Sheet1\nRevenue\nProfit\n", encoding="utf-8")
    gt = {
        "expected": {
            "embedded_content": [{"ref": "attachment_01", "texts": ["Revenue", "Profit"]}]
        }
    }

    assert score(package, gt)["embedded_content_hit"] == 1.0

    pure = _write_package(tmp_path, "parent body only\n")
    assert score(pure, gt)["embedded_content_hit"] == 0.0


def test_embedded_content_hit_searches_resource_preview_by_ref(tmp_path: Path):
    package = _write_package(tmp_path, "parent body only\n")
    resource = package / "structured" / "resources" / "attachment_01"
    resource.mkdir(parents=True)
    (resource / "preview.md").write_text(
        "# Resource Preview attachment_01\n\n"
        "| 字段 | 类型 |\n"
        "| --- | --- |\n"
        "| 单是否为临时策略 | 单项选择 |\n",
        encoding="utf-8",
    )
    gt = {
        "expected": {
            "embedded_content": [{"ref": "attachment_01", "texts": ["单是否为临时策略"]}]
        }
    }

    assert score(package, gt)["embedded_content_hit"] == 1.0


def test_markdown_portability_penalizes_nonportable_markdown(tmp_path: Path):
    content = (
        "Title\n=====\n\n"
        "![inline](data:image/png;base64,AAAA)\n\n"
        "![abs](/tmp/image.png)\n\n"
        "![missing](assets/images/missing.png)\n\n"
        "<table><tr><td>x</td></tr></table>\n"
    )
    package = _write_package(tmp_path, content)

    scored = score(package, {"expected": {"markdown_portability": True}})

    assert scored["markdown_portability"] < 1.0


def test_new_optional_dimensions_are_skipped_without_groundtruth(tmp_path: Path):
    package = _write_package(tmp_path, "# Only Heading\nbody text\n")
    scored = score(package, {"expected": {"headings": [{"level": 1, "text": "Only Heading"}]}})
    for dimension in (
        "footnote_recall",
        "footnote_anchor_accuracy",
        "table_cell_match",
        "asset_anchor_accuracy",
        "embedded_content_hit",
    ):
        assert scored[dimension] is None
    assert scored["markdown_portability"] is None


def test_score_perfect_package(tmp_path: Path):
    content = (
        "# Title\n\n"
        "| h1 | h2 |\n| --- | --- |\n| a | b |\n\n"
        "![alt](img1.png) ![alt2](img2.png)\n\n"
        "key phrase one and key phrase two\n"
    )
    package = _write_package(
        tmp_path,
        content,
        embedded_objects=[{"ref": "attachment_01", "type": "xlsx"}],
    )
    gt = {
        "expected": {
            "embedded_objects": [{"ref": "attachment_01", "type": "xlsx"}],
            "images": [{"ref": "image_01"}, {"ref": "image_02"}],
            "tables": [{"index": 0, "rows": 2, "cols": 2}],
            "headings": [{"level": 1, "text": "Title"}],
            "key_texts": ["key phrase one", "key phrase two"],
        }
    }
    scored = score(package, gt)
    assert scored["embedded_object_recall"] == 1.0
    assert scored["image_recall"] == 1.0
    assert scored["table_recall"] == 1.0
    assert scored["heading_recall"] == 1.0
    assert scored["heading_tree_score"] == 1.0
    assert scored["key_text_hit"] == 1.0
    assert scored["doc_score"] == 1.0


def test_score_missing_dimensions_are_skipped(tmp_path: Path):
    package = _write_package(tmp_path, "# Only Heading\nbody text\n")
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [{"level": 1, "text": "Only Heading"}],
            "key_texts": [],
        }
    }
    scored = score(package, gt)
    assert scored["embedded_object_recall"] is None
    assert scored["image_recall"] is None
    assert scored["table_recall"] is None
    assert scored["heading_recall"] == 1.0
    assert scored["heading_tree_score"] == 1.0
    assert scored["key_text_hit"] is None
    assert scored["doc_score"] == 1.0


def test_new_dimensions_are_registered() -> None:
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


def test_chart_text_recall_scores_chart_title_series_and_categories(tmp_path: Path):
    gt = {
        "expected": {
            "charts": [
                {
                    "index": 0,
                    "title": "I am a diagram",
                    "series": ["系列 1", "系列 2", "系列 3"],
                    "categories": ["类别1", "类别2", "类别3", "类别4"],
                    "texts": [
                        "I am a diagram",
                        "系列 1",
                        "系列 2",
                        "系列 3",
                        "类别1",
                        "类别2",
                        "类别3",
                        "类别4",
                    ],
                }
            ]
        }
    }
    docling_like = _write_package(tmp_path, "Intro\n\n<!-- image -->\n\nOutro\n")
    assert score(docling_like, gt)["chart_text_recall"] == 0.0

    recovered = _write_package(
        tmp_path,
        "I am a diagram\n\n系列 1 系列 2 系列 3\n\n类别1 类别2 类别3 类别4\n",
    )
    assert score(recovered, gt)["chart_text_recall"] == 1.0


def test_chart_text_recall_counts_pipeline_chart_preview(tmp_path: Path):
    gt = {
        "expected": {
            "charts": [
                {
                    "title": "I am a diagram",
                    "series": ["系列 1", "系列 2"],
                    "categories": ["类别1", "类别2"],
                }
            ]
        }
    }
    package = _write_package(
        tmp_path,
        "Intro\n\n[Chart 01: I am a diagram](resources/chart_01/preview.md)\n",
    )
    preview_dir = package / "structured" / "resources" / "chart_01"
    preview_dir.mkdir(parents=True)
    (preview_dir / "preview.md").write_text(
        "# Chart Preview chart_01\n\n"
        "Title: I am a diagram\n\n"
        "| Category | 系列 1 | 系列 2 |\n"
        "| --- | --- | --- |\n"
        "| 类别1 | 4.3 | 2.4 |\n"
        "| 类别2 | 2.5 | 4.4 |\n",
        encoding="utf-8",
    )

    assert score(package, gt)["chart_text_recall"] == 1.0


def test_smartart_text_recall_scores_smartart_text_separately(tmp_path: Path):
    gt = {
        "expected": {
            "smartarts": [
                {
                    "index": 0,
                    "source_path": "word/diagrams/drawing1.xml",
                    "texts": ["one", "two", "three", "four"],
                }
            ]
        }
    }

    image_only = _write_package(tmp_path, "Intro\n\n![diagram](assets/images/smartart.png)\n")
    assert score(image_only, gt)["smartart_text_recall"] == 0.0

    recovered = _write_package(tmp_path, "SmartArt nodes: one, two, three, four\n")
    assert score(recovered, gt)["smartart_text_recall"] == 1.0


def test_smartart_text_recall_counts_pipeline_diagram_preview(tmp_path: Path):
    gt = {
        "expected": {
            "smartarts": [
                {
                    "source_path": "word/diagrams/data1.xml",
                    "texts": ["one", "two", "three", "four"],
                }
            ]
        }
    }
    package = _write_package(
        tmp_path,
        "Intro\n\n[SmartArt 01](resources/diagram_01/preview.md)\n",
    )
    preview_dir = package / "structured" / "resources" / "diagram_01"
    preview_dir.mkdir(parents=True)
    (preview_dir / "preview.md").write_text(
        "# SmartArt Preview diagram_01\n\n"
        "Source part: `word/diagrams/data1.xml`\n\n"
        "- one\n"
        "- two\n"
        "- three\n"
        "- four\n",
        encoding="utf-8",
    )

    assert score(package, gt)["smartart_text_recall"] == 1.0


def test_checkbox_recall_counts_repeated_checkbox_lines(tmp_path: Path):
    gt = {
        "expected": {
            "checkboxes": [
                {"text": "□ 常规  ■ 加急"},
                {"text": "□ 通过  □ 需修订"},
                {"text": "□ 通过  □ 需修订"},
            ]
        }
    }
    package = _write_package(
        tmp_path,
        "| 字段 | 选项 |\n"
        "| --- | --- |\n"
        "| 紧急程度 | □ 常规 ■ 加急 |\n"
        "| 结果 | □ 通过 □ 需修订 |\n",
    )

    assert score(package, gt)["checkbox_recall"] == 2 / 3


def test_nested_table_asset_recall_scores_assets_inside_nested_tables(tmp_path: Path):
    gt = {
        "expected": {
            "nested_table_assets": [
                {
                    "table_index": 1,
                    "kind": "image",
                    "texts": ["Nested cell image"],
                },
                {
                    "table_index": 1,
                    "kind": "smartart",
                    "texts": ["one", "two"],
                },
            ]
        }
    }
    image_only = _write_package(tmp_path, "Nested cell image\n")
    assert score(image_only, gt)["nested_table_asset_recall"] == 0.5

    recovered = _write_package(tmp_path, "Nested cell image\n\nSmartArt: one two\n")
    assert score(recovered, gt)["nested_table_asset_recall"] == 1.0


def test_nested_table_asset_recall_counts_pipeline_diagram_preview(tmp_path: Path):
    gt = {
        "expected": {
            "nested_table_assets": [
                {
                    "table_index": 1,
                    "kind": "smartart",
                    "texts": ["eins", "zwei", "drei", "vier"],
                }
            ]
        }
    }
    package = _write_package(
        tmp_path,
        "Nested table contains [SmartArt 02](resources/diagram_02/preview.md)\n",
    )
    preview_dir = package / "structured" / "resources" / "diagram_02"
    preview_dir.mkdir(parents=True)
    (preview_dir / "preview.md").write_text(
        "# SmartArt Preview diagram_02\n\n"
        "- eins\n"
        "- zwei\n"
        "- drei\n"
        "- vier\n",
        encoding="utf-8",
    )

    assert score(package, gt)["nested_table_asset_recall"] == 1.0


def test_nested_table_recall_matches_nested_table_shape_and_text(tmp_path: Path):
    gt = {
        "expected": {
            "nested_tables": [
                {
                    "index": 1,
                    "depth": 1,
                    "rows": 2,
                    "cols": 2,
                    "texts": ["Nested Header", "Nested Value"],
                }
            ]
        }
    }
    recovered = _write_package(
        tmp_path,
        "<table>"
        "<tr><td>Outer</td><td>"
        "<table><tr><td>Nested Header</td><td>Status</td></tr>"
        "<tr><td>Nested Value</td><td>OK</td></tr></table>"
        "</td></tr>"
        "</table>\n",
    )
    assert score(recovered, gt)["nested_table_recall"] == 1.0

    missing_text = _write_package(
        tmp_path,
        "<table><tr><td>Outer</td><td><table><tr><td>Nested Header</td><td>Status</td></tr>"
        "<tr><td>Wrong</td><td>OK</td></tr></table></td></tr></table>\n",
    )
    assert score(missing_text, gt)["nested_table_recall"] == 0.0


def test_nested_table_recall_matches_pandoc_grid_table_inside_outer_cell(tmp_path: Path):
    gt = {
        "expected": {
            "nested_tables": [
                {
                    "index": 13,
                    "depth": 1,
                    "rows": 5,
                    "cols": 7,
                    "texts": ["index", "单元格公式", "公平性", "跨群体误差差异", "< 0.08", "效率"],
                }
            ]
        }
    }
    recovered = _write_package(
        tmp_path,
        "+-----------------+--------------------------------------------------------------------------------------------------------------+\n"
        "| Verwaltung      | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "|                 | | index      | 单元格公式                 | 解释             | 阈值    |         |         |         |       |\n"
        "|                 | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "|                 | | 1. 公平性  | $$E = mc^{2}\\ $$          | 跨群体误差差异   | \\< 0.08 |         |         |         |       |\n"
        "|                 | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "|                 | | 2. 信任度  | ![](assets/images/image.png) | 申诉与审计覆盖 | \\> 0.75 |         |         |         |       |\n"
        "|                 | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "|                 | | 3. 效率    |                            | 收益扣除复核成本 | \\> 0    |         |         |         |       |\n"
        "|                 | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "|                 | |            |                            |                  |         |         |         |         |       |\n"
        "|                 | +------------+----------------------------+------------------+---------+---------+---------+---------+       |\n"
        "+-----------------+--------------------------------------------------------------------------------------------------------------+\n",
    )

    assert score(recovered, gt)["nested_table_recall"] == 1.0


def test_build_groundtruth_extracts_nested_tables_checkboxes_and_charts(tmp_path: Path):
    docx = tmp_path / "rich.docx"
    _write_minimal_docx_with_nested_table_checkbox_and_chart(docx)

    gt = build_groundtruth(docx)
    expected = gt["expected"]

    assert expected["tables"] == [{"index": 0, "rows": 1, "cols": 2}]
    assert expected["nested_tables"] == [
        {
            "index": 1,
            "depth": 1,
            "rows": 2,
            "cols": 2,
            "texts": ["Nested Header", "Status", "Nested Value", "OK"],
        }
    ]
    assert expected["checkboxes"] == [{"text": "□ 常规  ■ 加急"}]
    assert expected["images"] == [
        {
            "ref": "image_01",
            "filename": "image1.png",
            "source_path": "word/media/image1.png",
        }
    ]
    assert expected["charts"] == [
        {
            "index": 0,
            "source_path": "word/charts/chart1.xml",
            "title": "I am a diagram",
            "series": ["系列 1"],
            "categories": ["类别1", "类别2"],
            "texts": ["I am a diagram", "系列 1", "类别1", "类别2"],
        }
    ]
    assert expected["smartarts"] == [
        {
            "index": 0,
            "source_path": "word/diagrams/data1.xml",
            "texts": ["one", "two"],
        }
    ]
    assert expected["nested_table_assets"] == [
        {"table_index": 1, "kind": "image", "rel_id": "rIdImage1", "texts": []},
        {"table_index": 1, "kind": "smartart", "rel_id": "rIdDiagramData1", "texts": ["one", "two"]},
    ]


def test_heading_recall_accepts_setext_markdown_headings(tmp_path: Path):
    package = _write_package(
        tmp_path,
        "Title\n=====\n\nSection\n-------\n\nbody text\n",
    )
    gt = {
        "expected": {
            "headings": [
                {"level": 1, "text": "Title"},
                {"level": 2, "text": "Section"},
            ],
        }
    }

    assert score(package, gt)["heading_recall"] == 1.0


def test_score_partial_image_and_table(tmp_path: Path):
    content = "![a](img1.png)\n\n| h |\n| --- |\n| x |\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "images": [{"ref": "image_01"}, {"ref": "image_02"}],
            "tables": [{"index": 0, "rows": 3, "cols": 2}],
            "headings": [],
            "embedded_objects": [],
            "key_texts": [],
        }
    }
    scored = score(package, gt)
    assert scored["image_recall"] == 0.5
    # table shape (3 rows x 2 cols) not matched by (2 rows x 1 col)
    assert scored["table_recall"] == 0.0


def test_build_groundtruth_extracts_embedded_xlsx(embedded_xlsx_docx: Path):
    gt = build_groundtruth(embedded_xlsx_docx)
    expected = gt["expected"]
    assert len(expected["embedded_objects"]) == 1
    assert expected["embedded_objects"][0]["type"] == "xlsx"
    assert expected["key_texts"] == []


def test_key_text_hit_searches_attachment_content(tmp_path: Path):
    package = _write_package(tmp_path, "parent body only\n")
    attachment = package / "structured" / "attachments" / "attachment_01"
    attachment.mkdir(parents=True)
    (attachment / "content.md").write_text("Revenue,Profit\n1200,320\n", encoding="utf-8")
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["Revenue"],  # only present in the attachment, not the parent body
        }
    }
    scored = score(package, gt)
    assert scored["key_text_hit"] == 1.0


def test_key_text_hit_searches_resource_preview_content(tmp_path: Path):
    package = _write_package(tmp_path, "parent body only\n")
    resource = package / "structured" / "resources" / "attachment_02"
    resource.mkdir(parents=True)
    (resource / "preview.md").write_text(
        "# Resource Preview attachment_02\n\n"
        "| 源ip | 单是否为临时策略 |\n"
        "| --- | --- |\n"
        "| 10.0.0.1 | 是 |\n",
        encoding="utf-8",
    )
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["单是否为临时策略"],
        }
    }

    assert score(package, gt)["key_text_hit"] == 1.0


def test_key_text_hit_normalizes_html_entities_and_superscripts(tmp_path: Path):
    # content rendered by MinerU (HTML-escaped >, <sup> tag) must match a
    # plain key_text written by a human.
    content = "threshold &gt;132 and limit 应小于0. 083m/s<sup>2</sup> here\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": [">132", "应小于0. 083m/s^2", "应小于0. 083m/s²"],
        }
    }
    scored = score(package, gt)
    assert scored["key_text_hit"] == 1.0


def test_key_text_hit_stray_lt_in_table_cell(tmp_path: Path):
    # markitdown emits literal "<3800" / "<3850" in table cells (not HTML-escaped).
    # The tag-stripping normalizer must not treat the bare "<" as the start of a
    # tag and swallow everything up to the ">" of ">2100".
    content = (
        "| 名称 | A型车 |\n| --- | --- |\n"
        "| 车体高度(mm) | <3800 |\n"
        "| 车内净高(mm) | >2100 |\n"
    )
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": [">2100", "<3800"],
        }
    }
    scored = score(package, gt)
    assert scored["key_text_hit"] == 1.0


def test_key_text_hit_ignores_markdown_emphasis_and_spacing(tmp_path: Path):
    content = "**表A.1** **基础统计指标汇总表(续)**\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["表A.1  基础统计指标汇总表(续)"],
        }
    }

    assert score(package, gt)["key_text_hit"] == 1.0


def test_key_text_hit_ignores_markdown_backslash_escapes(tmp_path: Path):
    content = "客流量【5\\]= ∑车站进站量【4a】\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["客流量【5]= ∑车站进站量【4a】"],
        }
    }

    assert score(package, gt)["key_text_hit"] == 1.0


def test_key_text_hit_ignores_pandoc_escaped_quotes_in_paths(tmp_path: Path):
    content = (
        "| \\--SavePath\\' \\'./result/benchmark_bs{bs}\\_in"
        "{input_length}\\_out{output_length}/\\' |\n"
    )
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["'./result/benchmark_bs{bs}_in{input_length}_out{output_length}/'"],
        }
    }

    assert score(package, gt)["key_text_hit"] == 1.0


def test_key_text_hit_normalizes_standard_number_dashes_and_spaces(tmp_path: Path):
    content = "DB31/T 1122---2018\n"
    package = _write_package(tmp_path, content)
    gt = {
        "expected": {
            "embedded_objects": [],
            "images": [],
            "tables": [],
            "headings": [],
            "key_texts": ["DB31/T  1122—2018"],
        }
    }

    assert score(package, gt)["key_text_hit"] == 1.0


def _write_minimal_docx_with_nested_table_checkbox_and_chart(path: Path) -> None:
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p><w:r><w:t>□ 常规  ■ 加急</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
    <w:tbl>
      <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Outer</w:t></w:r></w:p></w:tc>
        <w:tc>
          <w:tbl>
            <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>Nested Header</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Status</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Nested Value</w:t></w:r></w:p></w:tc>
              <w:tc>
                <w:p><w:r><w:t>OK</w:t></w:r></w:p>
                <w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><a:blip r:embed="rIdImage1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                <w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdDiagramData1" r:lo="rIdDiagramLayout1" r:qs="rIdDiagramStyle1" r:cs="rIdDiagramColors1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>
"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rIdDiagramData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
  <Relationship Id="rIdDiagramLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>
  <Relationship Id="rIdDiagramStyle1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>
  <Relationship Id="rIdDiagramColors1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="diagrams/colors1.xml"/>
</Relationships>
"""
    chart_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>I am a diagram</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>系列 1</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>类别1</c:v></c:pt><c:pt idx="1"><c:v>类别2</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"""
    smartart_data_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt><dgm:t><a:p><a:r><a:t>one</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt><dgm:t><a:p><a:r><a:t>two</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>
"""
    smartart_drawing_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dsp:spTree>
    <dsp:sp><dsp:txBody><a:p><a:r><a:t>one</a:t></a:r></a:p></dsp:txBody></dsp:sp>
    <dsp:sp><dsp:txBody><a:p><a:r><a:t>two</a:t></a:r></a:p></dsp:txBody></dsp:sp>
  </dsp:spTree>
</dsp:drawing>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
  <Override PartName="/word/diagrams/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramDrawing+xml"/>
</Types>
"""
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", rels_xml)
        docx.writestr("word/charts/chart1.xml", chart_xml)
        docx.writestr("word/diagrams/data1.xml", smartart_data_xml)
        docx.writestr("word/diagrams/drawing1.xml", smartart_drawing_xml)
        docx.writestr("word/diagrams/layout1.xml", "<xml/>")
        docx.writestr("word/diagrams/quickStyle1.xml", "<xml/>")
        docx.writestr("word/diagrams/colors1.xml", "<xml/>")
        docx.writestr("word/media/image1.png", b"png")
