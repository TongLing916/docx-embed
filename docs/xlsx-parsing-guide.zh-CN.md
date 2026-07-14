# XLSX 解析指南

> English: [XLSX Parsing Guide](xlsx-parsing-guide.md)

`docx-embed` 项目提供深度 XLSX 工作簿解析，远超简单的单元格值提取。本文档涵盖解析策略、输出结构、资产提取及从真实工作簿处理中总结的最佳实践。

## 快速开始

```python
from edp.xlsx.parser import parse_xlsx_package

package = parse_xlsx_package("report.xlsx", "output/", "workbook_01")

# content.md     — 所有表格的可读 Markdown 预览
# tables/        — 每个子表的 table_001.csv + table_001.json
# assets/        — 提取的图片、图表、嵌入对象
# workbook.json  — 工作簿级元数据
```

## 解析策略

### 双工作簿加载

Openpyxl 以两种模式同时加载：

| 模式 | 用途 |
| ---- | ---- |
| `data_only=False` | 捕获公式、超链接、批注、数字格式 |
| `data_only=True`  | 捕获计算结果 |

每个单元格记录合并两种视图：

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

### 子表拆分

真实工作簿经常把多个逻辑表格塞进同一个工作表，用空行或重复表头分隔。解析器通过三种信号检测边界：

1. **重复表头行**——某行的非空值签名在整个工作表中出现 ≥3 次，且包含 ≥2 个真正的非数字文本单元格时，判定为重复表头。每次出现都开启一个新子表。

2. **空行间隙**——连续 ≥2 行完全为空时，视为表边界。

3. **图片锚点**——当图片或图表锚定在已与其他边界信号重合的行时，确认该分割。孤立在连续表格内的图片**不会**强制分割。

```json
{
  "sheet": "Inspection",
  "start_row": 1,
  "end_row": 3,
  "boundary_signals": ["header_repeat"]
}
```

### 行列裁剪

每个子表中，完全为空的行和列会被移除以减少噪声。这对真实工作簿尤为重要——它们常常用大段空列作视觉分隔。

## 输出文件

### 每表 JSON（`tables/table_NNN.json`）

每条表 JSON 记录包含：

- `sheet` — 工作表名称
- `state` — `visible` / `hidden` / `veryHidden`
- `start_row` / `end_row` — 1-indexed 行范围
- `boundary_signals` — 触发拆分的信号列表
- `rows` — 裁剪后的二维单元格值数组
- `cells` — 稀疏字典，`A1` → 单元格元数据（value、formula、hyperlink、comment、data_type、number_format）
- `merged_ranges` — 与此子表相交的合并单元格范围
- `hidden_rows` / `hidden_columns` — 隐藏行列标注
- `chart_count` / `image_count` — 锚定在此子表的资产数量
- `assets` — 过滤后的资产记录（图片、图表、嵌入对象、公式）
- `framework_result` — 可选框架解析器的侧车结果

### 每表 CSV（`tables/table_NNN.csv`）

纯 UTF-8 CSV，已裁剪行列。空单元格以空字符串写出。

### 工作簿元数据（`workbook.json`）

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

### 内容 Markdown（`content.md`）

可读的 Markdown 预览：

- 渲染**所有**行（不做 10 行截断）
- 将单元格内换行转为 `<br>` 以保持 GFM 表结构
- 列出提取的资产及其锚点位置
- 根据内容语言自适应标签（中文 / 英文）

## 资产提取

图片、图表、OLE 对象和公式通过 XLSX drawing 关系链提取：

```
xl/workbook.xml
  └── sheet → xl/worksheets/sheetN.xml
        └── xl/drawings/drawingM.xml
              └── xl/media/*       (图片)
              └── xl/charts/*     (图表 XML)
              └── xl/embeddings/*  (OLE 对象)
```

### 图片识别

小型图片被检测为可能的 logo 并跳过 VLM 描述：

1. **文件名启发式**：`logo*`、`brand*`、`icon*`、`favicon*`、`symbol*`、
   `mark*`、`badge*`、`emblem*` → 始终视为 logo。
2. **尺寸阈值**：面积 < `EDP_LOGO_MAX_AREA`（默认 40 000 px²）→ 可能为 logo。

设置 `EDP_LOGO_MAX_AREA=0` 可禁用基于尺寸的检测。

### 语义增强（可选）

配置后，提取的图片可进行 OCR 和 VLM 描述：

```bash
export EDP_PADDLEOCR_AUTHORIZATION="Bearer xxx"
export EDP_PADDLEOCR_URL="https://ocr.example.com/v1/ocr"
export EDP_VLM_AUTHORIZATION="Bearer xxx"
export EDP_VLM_URL="https://vlm.example.com/v1/chat/completions"
```

图片描述写入 `assets/resources/{image_ref}/description.md` 和 `description.json`，标签根据内容主导语言自适应。

## 语言自适应输出

解析器检测工作簿内容是否以中文为主，并据此切换标签语言：

| 英文 | 中文 |
| ---- | ---- |
| `Caption: ` | `标题: ` |
| `OCR: ` | `OCR文字: ` |
| `Equation` | `公式` |
| `Embedded` | `嵌入对象` |
| `see ` | `详见 ` |
| `(untitled)` | `(无标题)` |

语言检测要求中文字符占字母文本 ≥30%。

## 框架侧车（可选）

安装 [`unstructured`](https://github.com/Unstructured-IO/unstructured) 后，其 xlsx 解析器作为侧车对比运行：

```bash
pip install "unstructured[xlsx]"
```

结果写入每个表的 `framework_result.unstructured` 字段和 `workbook.json → framework_results`。失败静默处理——不抛异常、不告警。仅用于对比和评测。

## 性能注意事项

- **大型工作簿**：Openpyxl 将整个工作簿加载到内存。对于超大文件（>100MB），建议预过滤工作表。
- **子表拆分**：签名计数的最坏情况为 O(n²)。真实工作簿每个工作表通常不超过几百行。
- **资产提取**：图表和 OLE 对象按需从 ZIP 解压。SHA256 去重避免写入相同载荷两次。

## 与简单方案的对比

| 功能 | 本解析器 | 仅 `openpyxl` | `pandas.read_excel` |
| ---- | -------- | ------------- | -------------------- |
| 公式捕获 | ✓ | ✓ (with on_demand) | ✗ |
| 超链接/批注 | ✓ | ✓ | ✗ |
| 子表拆分 | ✓ | ✗ | ✗ |
| 合并范围 | ✓ | ✓ | ✗ |
| 图片/图表提取 | ✓ | ✗ | ✗ |
| 语言自适应输出 | ✓ | ✗ | ✗ |
| 框架侧车 | ✓ | ✗ | ✗ |
| 隐藏行列追踪 | ✓ | ✓ | ✗ |

## 实例

### 多表工作表拆分

一张包含 50+ 个工段的检修工作表，每个工段以相同表头行开头：

```
| 部件名称 | 检修部位 | 标准或限度 |
| 受电弓   | 表面清洁 | 无破损     |
| ... 48 more rows ... |
| 部件名称 | 检修部位 | 标准或限度 |   ← 检测为边界
| 牵引电机 | 定子绕组 | 绝缘良好   |
| ... etc ... |
```

解析器从单个工作表中检测出 10+ 个子表，每个子表有独立的 `table_NNN.csv` 和 `table_NNN.json`。

### 公式保留

带计算列的财务工作簿：

```
A1: "Region"   B1: "Q1"   C1: "Q2"   D1: "Total"
A2: "North"    B2: 10      C2: 15      D2: "=SUM(B2:C2)"
```

D2 单元格记录同时捕获公式（`=SUM(B2:C2)`）和计算结果（25），适用于公式审计和数据提取两种场景。

### 多类型数据

工作簿中常见的数据类型全部正确处理：

| 类型 | 示例值 | 导出格式 |
| ---- | ------ | -------- |
| 文本 | "Hello World" | 字符串 |
| 整数 | 42 | 整数 |
| 浮点 | 3.14159 | 浮点数 |
| 布尔 | TRUE | 布尔 |
| 日期 | 2025-12-31 | ISO 日期 |
| 日期时间 | 2025-06-15 14:30:00 | ISO 8601 |
| 百分比 | 85% | 小数 0.85 |
| 货币 | $1,999.99 | 浮点数 |

运行以下命令生成示例工作簿并查看完整输出：

```bash
python tests/generate_demo_xlsx.py tests/fixtures/demo_workbook.xlsx
uv run python -c "
from edp.xlsx.parser import parse_xlsx_package
pkg = parse_xlsx_package('tests/fixtures/demo_workbook.xlsx', 'demo_output/', 'demo')
print(f'Tables: {len(pkg.tables)}')
print(open(pkg.content_path).read())
"
```
