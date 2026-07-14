# XLSX Parsing Guide

The `docx-embed` project provides deep XLSX workbook parsing that goes beyond
simple cell-value extraction. This guide covers the parsing strategy, output
structure, asset extraction, and best practices learned from real-world
workbook processing.

## Quick Start

```python
from edp.xlsx.parser import parse_xlsx_package

package = parse_xlsx_package("report.xlsx", "output/", "workbook_01")

# content.md  — human-readable Markdown preview of all tables
# tables/     — table_001.csv, table_001.json (one pair per sub-table)
# assets/     — extracted images, charts, embedded objects
# workbook.json — workbook-level metadata
```

## Parsing Strategy

### Dual-Workbook Loading

Openpyxl is used in two modes simultaneously:

| Mode | Purpose |
| ---- | ------- |
| `data_only=False` | Captures formulas, hyperlinks, comments, number formats |
| `data_only=True`  | Captures computed cell values |

Each cell record combines both views:

```json
{
  "D2": {
    "value": 25,
    "data_type": "n",
    "formula": "=SUM(B2:C2)",
    "display_value": 25
  }
}
```

### Sub-Table Splitting

Real-world workbooks often pack multiple logical tables into a single sheet,
separated by blank rows or repeated headers. The parser detects boundaries
using three signals:

1. **Repeated header rows** — A row whose non-empty value signature appears
   ≥3 times across the sheet and contains ≥2 genuinely non-numeric text cells
   is treated as a recurring header. Each occurrence starts a new sub-table.

2. **Blank-row gaps** — Two or more consecutive fully-empty rows signal a
   table boundary.

3. **Image anchors** — When an image or chart is anchored at a row that
   already coincides with another boundary signal, it confirms the split.
   Standalone images inside a continuous table do **not** force a split.

```json
{
  "sheet": "Inspection",
  "start_row": 1,
  "end_row": 3,
  "boundary_signals": ["header_repeat"]
}
```

### Row and Column Pruning

Within each sub-table, entirely blank rows and columns are removed to reduce
noise. This is particularly important for real-world workbooks that allocate
wide empty columns as visual separators.

## Output Files

### Per-table JSON (`tables/table_NNN.json`)

Each table JSON record contains:

- `sheet` — worksheet name
- `state` — `visible` / `hidden` / `veryHidden`
- `start_row` / `end_row` — 1-indexed row range
- `boundary_signals` — which signals triggered the split
- `rows` — pruned 2D array of cell values
- `cells` — sparse dict of `A1` → cell metadata (value, formula, hyperlink, comment, data_type, number_format)
- `merged_ranges` — merged-cell ranges intersecting this sub-table
- `hidden_rows` / `hidden_columns` — hidden row/column annotations
- `chart_count` / `image_count` — assets anchored in this sub-table
- `assets` — filtered asset records (images, charts, attachments, equations)
- `framework_result` — sidecar results from optional framework parsers

### Per-table CSV (`tables/table_NNN.csv`)

Plain UTF-8 CSV with pruned rows and columns. Empty cells are written as
empty strings.

### Workbook Metadata (`workbook.json`)

```json
{
  "source_file": "report.xlsx",
  "sheet_count": 3,
  "sheets": [
    {
      "name": "Financials",
      "state": "visible",
      "max_row": 50,
      "max_column": 8,
      "merged_ranges": ["A1:C1"],
      "hidden_rows": [5],
      "hidden_columns": ["D"],
      "assets": { ... }
    }
  ],
  "assets": {
    "images": [ ... ],
    "charts": [ ... ],
    "attachments": [ ... ],
    "equations": [ ... ]
  }
}
```

### Content Markdown (`content.md`)

A human-readable Markdown preview that:

- Renders **all** rows (no 10-row truncation)
- Handles intra-cell newlines as `<br>` to preserve GFM table structure
- Lists extracted assets with anchor positions
- Adapts labels to the content language (English / Chinese)

## Asset Extraction

Images, charts, OLE objects, and equations are extracted via the XLSX drawing
relationship chain:

```
xl/workbook.xml
  └── sheet → xl/worksheets/sheetN.xml
        └── xl/drawings/drawingM.xml
              └── xl/media/*       (images)
              └── xl/charts/*     (chart XML)
              └── xl/embeddings/*  (OLE objects)
```

### Image Recognition

Small images are detected as likely logos and skipped for VLM captioning:

1. **Filename heuristics**: `logo*`, `brand*`, `icon*`, `favicon*`, `symbol*`,
   `mark*`, `badge*`, `emblem*` → always treated as logo.
2. **Dimension threshold**: area < `EDP_LOGO_MAX_AREA` (default 40 000 px²) →
   likely logo.

Set `EDP_LOGO_MAX_AREA=0` to disable dimension-based detection.

### Semantic Enrichment (Optional)

When configured, extracted images receive OCR and VLM captions:

```bash
export EDP_PADDLEOCR_AUTHORIZATION="Bearer xxx"
export EDP_PADDLEOCR_URL="https://ocr.example.com/v1/ocr"
export EDP_VLM_AUTHORIZATION="Bearer xxx"
export EDP_VLM_URL="https://vlm.example.com/v1/chat/completions"
```

Image descriptions are written to `assets/resources/{image_ref}/description.md`
and `description.json`, with labels adapting to the dominant content language.

## Language-Aware Output

The parser detects when workbook content is predominantly Chinese and switches
label language accordingly:

| English | Chinese |
| ------- | ------- |
| `Caption: ` | `标题: ` |
| `OCR: ` | `OCR文字: ` |
| `Equation` | `公式` |
| `Embedded` | `嵌入对象` |
| `see ` | `详见 ` |
| `(untitled)` | `(无标题)` |

Language detection requires ≥30% Chinese characters among alphabetic text.

## Framework Sidecars (Optional)

When [`unstructured`](https://github.com/Unstructured-IO/unstructured) is
installed, its xlsx partitioner runs as a sidecar comparison:

```bash
pip install "unstructured[xlsx]"
```

Results are written to each table's `framework_result.unstructured` field
and to `workbook.json → framework_results`. Failures are silent — no
exceptions, no warnings. This is purely for comparison and evaluation.

## Performance Notes

- **Large workbooks**: Openpyxl loads the entire workbook into memory. For
  very large files (>100MB), consider pre-filtering sheets.
- **Sub-table splitting**: O(n²) in the worst case for signature counting.
  Real-world workbooks rarely exceed a few hundred rows per sheet.
- **Asset extraction**: Charts and OLE objects are decompressed from the ZIP
  on demand. SHA256 deduplication prevents writing identical payloads twice.

## Comparison with Simple Approaches

| Feature | This Parser | `openpyxl` alone | `pandas.read_excel` |
| ------- | ----------- | ---------------- | -------------------- |
| Formulas captured | ✓ | ✓ (with on_demand) | ✗ |
| Hyperlinks/Comments | ✓ | ✓ | ✗ |
| Sub-table splitting | ✓ | ✗ | ✗ |
| Merged ranges | ✓ | ✓ | ✗ |
| Image/Chart extraction | ✓ | ✗ | ✗ |
| Language-aware output | ✓ | ✗ | ✗ |
| Framework sidecars | ✓ | ✗ | ✗ |
| Hidden row/col tracking | ✓ | ✓ | ✗ |

## Real-World Examples

### Multi-table Sheet Splitting

A maintenance inspection sheet with 50+ sections, each starting with the same
header row:

```
| 部件名称 | 检修部位 | 标准或限度 |
| 受电弓   | 表面清洁 | 无破损     |
| ... 48 more rows ... |
| 部件名称 | 检修部位 | 标准或限度 |   ← detected as boundary
| 牵引电机 | 定子绕组 | 绝缘良好   |
| ... etc ... |
```

The parser detects 10+ sub-tables from a single worksheet, each with its own
`table_NNN.csv` and `table_NNN.json`.

### Formula Preservation

Financial workbooks with computed columns:

```
A1: "Region"   B1: "Q1"   C1: "Q2"   D1: "Total"
A2: "North"    B2: 10      C2: 15      D2: "=SUM(B2:C2)"
```

The D2 cell record captures both the formula (`=SUM(B2:C2)`) and the
computed value (25), making it suitable for both formula auditing and
data extraction workflows.
