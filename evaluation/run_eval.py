"""Run the V0.2 structural-fidelity evaluation across all parsed packages.

For each ground-truth file in ``evaluation/groundtruth/<doc_id>.json`` and each
configured method under ``output_batch/<doc_id>/<method>/``, compute the
``layer_fidelity`` dimension scores and write:

* ``evaluation/results.tsv``        - one row per (doc, method)
* ``evaluation/results_summary.tsv`` - per-method cross-document means
* ``evaluation/misses.tsv``         - one row per missing item (doc, method, dimension, detail)
* ``evaluation/misses_report.md``   - human-readable report grouped by method
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from evaluation.layer_fidelity import DIMENSIONS, diagnose, score, score_gold


METHODS = (
    "pipeline-markitdown",
    "pipeline-mineru-pipeline",
    "pipeline-mineru-vlm-engine",
    "pipeline-mineru-hybrid-engine",
    "pipeline-docling",
    "pipeline-pandoc",
    "pipeline-ragflow",
    "pure-markitdown",
    "pure-mineru-pipeline",
    "pure-mineru-vlm-engine",
    "pure-mineru-hybrid-engine",
    "pure-docling",
    "pure-pandoc",
    "pure-ragflow",
)
NULL = "-"


def run_eval(
    output_batch: Path = REPO_ROOT / "output_batch",
    gt_root: Path = REPO_ROOT / "evaluation" / "groundtruth",
    gold_root: Path = REPO_ROOT / "evaluation" / "gold",
    results_path: Path = REPO_ROOT / "evaluation" / "results.tsv",
    summary_path: Path = REPO_ROOT / "evaluation" / "results_summary.tsv",
    misses_path: Path = REPO_ROOT / "evaluation" / "misses.tsv",
    misses_report_path: Path = REPO_ROOT / "evaluation" / "misses_report.md",
) -> int:
    rows: list[dict[str, Any]] = []
    misses: list[dict[str, str]] = []
    gold_dirs = {p.name for p in gold_root.glob("*") if p.is_dir()} if gold_root.exists() else set()

    for gt_file in sorted(gt_root.glob("*.json")):
        gt = json.loads(gt_file.read_text(encoding="utf-8"))
        doc_id = gt.get("doc_id") or gt_file.stem
        for method in METHODS:
            package_dir = output_batch / doc_id / method
            if not package_dir.exists():
                continue
            diagnosed = diagnose(package_dir, gt)
            row = {"doc_id": doc_id, "method": method, "doc_score": _doc_score(diagnosed)}
            for dim in DIMENSIONS:
                row[dim] = diagnosed[dim].score
                for detail in diagnosed[dim].misses:
                    misses.append(
                        {"doc_id": doc_id, "method": method, "dimension": dim, "missing": detail}
                    )
            if doc_id in gold_dirs:
                gold_scored = score_gold(package_dir, gold_root / doc_id)
                for key, value in gold_scored.items():
                    row[f"gold_{key}"] = value
            rows.append(row)

    if not rows:
        print("No matching packages found.", file=sys.stderr)
        return 1

    _write_results(rows, results_path)
    _write_summary(rows, summary_path)
    _write_misses(misses, misses_path)
    _write_misses_report(misses, misses_report_path)
    print(f"wrote {results_path.relative_to(REPO_ROOT)} ({len(rows)} rows)")
    print(f"wrote {summary_path.relative_to(REPO_ROOT)}")
    print(f"wrote {misses_path.relative_to(REPO_ROOT)} ({len(misses)} missing items)")
    print(f"wrote {misses_report_path.relative_to(REPO_ROOT)}")
    return 0


def _doc_score(diagnosed: dict[str, Any]) -> float:
    applicable = [result.score for result in diagnosed.values() if result.score is not None]
    return sum(applicable) / len(applicable) if applicable else 0.0


def _write_results(rows: list[dict[str, Any]], path: Path) -> None:
    gold_cols = [key for key in rows[0] if key.startswith("gold_")] if rows else []
    text_headers = {"doc_id", "method"}
    headers = ["doc_id", "method", *DIMENSIONS, "doc_score", *gold_cols]
    lines = ["\t".join(headers)]
    for row in rows:
        cells = []
        for header in headers:
            value = row.get(header)
            if header in text_headers:
                cells.append(str(value))
            else:
                cells.append(_format_score(value))
        lines.append("\t".join(cells))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _write_summary(rows: list[dict[str, Any]], path: Path) -> None:
    metric_keys = [*DIMENSIONS, "doc_score"]
    gold_keys = sorted({key for row in rows for key in row if key.startswith("gold_")})
    headers = ["method", "n_docs", *metric_keys, *gold_keys]
    lines = ["\t".join(headers)]
    for method in METHODS:
        method_rows = [row for row in rows if row["method"] == method]
        if not method_rows:
            continue
        cells = [method, str(len(method_rows))]
        for key in metric_keys:
            cells.append(_format_score(_mean(method_rows, key)))
        for key in gold_keys:
            cells.append(_format_score(_mean(method_rows, key)))
        lines.append("\t".join(cells))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _write_misses(misses: list[dict[str, str]], path: Path) -> None:
    headers = ["doc_id", "method", "dimension", "missing"]
    lines = ["\t".join(headers)]
    for miss in misses:
        lines.append("\t".join(str(miss[h]) for h in headers))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _write_misses_report(misses: list[dict[str, str]], path: Path) -> None:
    by_method: dict[str, list[dict[str, str]]] = defaultdict(list)
    for miss in misses:
        by_method[miss["method"]].append(miss)
    lines = ["# Missing-item report", ""]
    lines.append(f"Total missing items: **{len(misses)}** across {len(by_method)} methods.")
    lines.append("")
    for method in METHODS:
        method_misses = by_method.get(method, [])
        lines.append(f"## {method} ({len(method_misses)} missing)")
        if not method_misses:
            lines.append("")
            lines.append("_(no missing items)_")
            lines.append("")
            continue
        by_doc: dict[str, list[dict[str, str]]] = defaultdict(list)
        for miss in method_misses:
            by_doc[miss["doc_id"]].append(miss)
        lines.append("")
        for doc_id in sorted(by_doc):
            lines.append(f"- **{doc_id}**")
            for miss in sorted(by_doc[doc_id], key=lambda m: (m["dimension"], m["missing"])):
                lines.append(f"  - {miss['dimension']}: {miss['missing']}")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def _mean(rows: list[dict[str, Any]], key: str) -> float | None:
    values = [row[key] for row in rows if row.get(key) is not None]
    return statistics.fmean(values) if values else None


def _format_score(value: float | None) -> str:
    if value is None:
        return NULL
    return f"{value:.4f}"
    if value is None:
        return NULL
    return f"{value:.4f}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the V0.2 structural-fidelity evaluation.")
    parser.add_argument("--output-batch", type=Path, default=REPO_ROOT / "output_batch")
    parser.add_argument("--gt-root", type=Path, default=REPO_ROOT / "evaluation" / "groundtruth")
    parser.add_argument("--gold-root", type=Path, default=REPO_ROOT / "evaluation" / "gold")
    parser.add_argument("--results", type=Path, default=REPO_ROOT / "evaluation" / "results.tsv")
    parser.add_argument("--summary", type=Path, default=REPO_ROOT / "evaluation" / "results_summary.tsv")
    parser.add_argument("--misses", type=Path, default=REPO_ROOT / "evaluation" / "misses.tsv")
    parser.add_argument("--misses-report", type=Path, default=REPO_ROOT / "evaluation" / "misses_report.md")
    args = parser.parse_args(argv)
    return run_eval(
        args.output_batch,
        args.gt_root,
        args.gold_root,
        args.results,
        args.summary,
        args.misses,
        args.misses_report,
    )


if __name__ == "__main__":
    raise SystemExit(main())
