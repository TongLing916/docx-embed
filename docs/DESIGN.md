# EDP 技术设计文档 v1.0

> 一个站在 MinerU/Docling 肩膀上的增强模块，专门收拾它们没解决的烂摊子——**文档嵌入对象（OLE）的提取、递归解析与按位拼装**。

---

## 1. 项目定位与核心主张

现有主流开源文档解析框架（MinerU、Docling、Unstructured 等）普遍存在一个致命短板：**对 DOCX 中嵌入的 XLSX、PDF、PPT 等 OLE 对象视而不见，或直接丢弃。**  
EDP 不做重复造轮子，只补那一块最短的板。它作为预处理增强层，与基座框架解耦，输入原始文件，输出**标准化的、自包含的文档包**。

**核心价值**：
- **完整性**：嵌入对象不再丢失，100% 提取并纳入文档包结构。
- **可递归**：支持嵌入文件内再嵌文件的深度遍历与解析。
- **可评测**：自带分层评测脚本，用数据量化基座框架的缺失与本模块的增益。
- **可消费**：输出结构化的文档包（含 `manifest.json`），下游 Agent 无需任何改动即可直接消费。

---

## 2. 处理流水线

```
[原始 DOCX]
       │
       ▼
┌──────────────────────────┐
│ 阶段0: 预检与提取          │  核心自研
│ - python-docx 定位 OLE    │
│ - olefile/oleobj 提取     │
│ - 产出: 干净主文档 + 子文件列表 + 位置映射表
└──────────┬───────────────┘
           │
     ┌─────┴─────┐
     ▼           ▼
[主文档]     [子文件列表]
     │           │
     ▼           ▼
┌──────────┐ ┌──────────────────┐
│ 阶段1a:  │ │ 阶段1b: 浅 preview │
│ MinerU/  │ │ txt/log/csv 前段文本│
│ Docling  │ │ xlsx sheet/前几行  │
│ 解析主   │ │ 其他资源只登记      │
│ 文档正文 │ └────────┬─────────┘
└────┬─────┘          │
     │                │
     ▼                ▼
[主文档结构化包]  [资源登记表与预览]
     │                │
     └──────┬─────────┘
            ▼
┌──────────────────────────┐
│ 阶段2: 结果拼装           │  核心自研
│ - 在 content.md 中插入    │
│   行内 Markdown 引用       │
│ - 合并 resources/assets   │
│ - 写入 embedded_resources │
│ - 生成 manifest.json      │
└────────────┬─────────────┘
             │
             ▼
┌──────────────────────────┐
│ 输出: 标准文档包          │
│ {doc_id}/                │
│ ├── raw/                 │
│ ├── structured/          │
│ └── manifest.json        │
└──────────────────────────┘
```

---

## 3. 标准文档包结构

```
{document_id}/
├── raw/                         # 原始文件及提取的嵌入对象备份
│   ├── original.docx
│   └── embedded/
│       ├── attachment_01.xlsx
│       └── attachment_02.pdf
├── structured/                  # 结构化产出，Agent 直接消费
│   ├── content.md               # 干净主文档正文，含行内 Markdown 引用
│   ├── child_files.md           # 图片与附件子文件列表
│   ├── position_map.csv         # 原文位置到子文件 ref/path 的映射
│   ├── embedded_resources.jsonl # 资源登记表，含 hash/mime/anchor/parse_status
│   ├── assets/                  # 图片、图表等二进制资源
│   │   ├── images/
│   │   │   ├── image_01.png
│   │   │   └── image_01_desc.txt # OCR 或视觉模型描述
│   │   └── other/
│   │       └── chart_001_data.json # 图表提取出的数据表
│   ├── resources/
│   │   └── attachment_01/
│   │       ├── preview.md       # 浅 preview，Agent 默认入口
│   │       └── preview.json     # XLSX 等结构化 preview 元数据
│   ├── _meta/                   # 日志、评测结果与调试痕迹
│   │   ├── parse_log.txt
│   │   ├── work/
│   │   │   └── clean/
│   │   │       └── main.docx    # 替换附件/图片为哨兵位的干净主文档
│   │   └── parsers/
│   │       ├── markitdown/
│   │       ├── mineru/
│   │       ├── docling/
│   │       ├── pandoc/
│   │       └── ragflow/
│   │   └── first_layer_eval.json
│   ├── chunks/
│   │   └── chunks.jsonl         # 用于 RAG 索引的标准化 chunks
│   └── attachments/             # 显式完整解析时的子文档结构
│       └── attachment_01/
│           ├── content.md
│           ├── tables/
│           └── chunks/
└── manifest.json                # 文档包清单
```

---

## 4. 核心数据结构

### 4.1 Manifest

```python
@dataclass
class Manifest:
    doc_id: str
    original_file: str
    parse_timestamp: str
    pipeline_version: str
    parse_status: str           # "success" | "partial" | "failed"
    content_stats: ContentStats
    parse_warnings: List[str]
    content_map: Dict[str, Any]
```

### 4.2 主文档、子文件列表与位置映射

主文档 `structured/content.md` 保持干净正文，只在原始位置插入行内 Markdown 引用；不使用底部引用定义块：

```markdown
项目利润说明

[2024年Q3利润明细表.xlsx](resources/attachment_01/preview.md)

现场照片如下：

![](assets/images/image_01.png)

原始说明附件：

[notes.txt](resources/attachment_02/preview.md)
```

子文件信息放在 `structured/child_files.md`，位置关系放在 `structured/position_map.csv`，完整资源登记放在 `structured/embedded_resources.jsonl`。`child_files.md` 同时包含适合人工扫描的 Markdown 表格和适合程序读取的 `yaml` fenced block。Agent 先读干净主文档，再按 ref 查询子文件列表、位置映射表或资源登记表，必要时进入 `structured/resources/{ref}/preview.md`、`preview.json`、原始附件或图片资源。

```yaml
child_files:
  - ref: "attachment_01"
    kind: "attachment"
    filename: "2024年Q3利润明细表.xlsx"
    markdown_reference: "[2024年Q3利润明细表.xlsx](resources/attachment_01/preview.md)"
    path: "structured/resources/attachment_01/preview.md"
    entry_point: "structured/resources/attachment_01/preview.md"
    tables: []
  - ref: "image_01"
    kind: "image"
    filename: "现场照片.png"
    markdown_reference: "![](assets/images/image_01.png)"
    path: "structured/assets/images/image_01.png"
    entry_point: ""
    tables: []
```

> 处理流程入口、运行命令与环境变量见 [README.md § Quick start](../README.md#quick-start)。

---

## 5. 评测体系

### 5.1 分层评测

- **第一层（文档转换评测）**：验证文档转换的保真度，完全基于规则，不依赖 LLM。
- **第二层（Agent QA 评测）**：在理想文档包上测试 Agent 能力，归因推理错误。

### 5.2 信息点卡

从文档中抽取 10-15 个关键事实，构建评测基准集：

```json
{
  "id": "IF_01",
  "type": "text_fact",
  "keywords": ["2024", "Q3", "利润", "1200万"],
  "page": 3
}
```

### 5.3 核心指标

| 指标 | 定义 | V1.0 目标 |
|------|------|-----------|
| 嵌入对象结构完整性 | 嵌入附件出现在正确位置的比例 | 100% |
| 嵌入对象内容保真度 | 子文档内信息点召回率 | ≥ 90% |
| 文本事实召回率 | 全文关键信息点保留率 | ≥ 95% |
| 表格单元格准确率 | 指定单元格值正确率 | ≥ 90% |

评测脚本不依赖 LLM，直接比对关键词、表格位置、文件存在性，产出冷冰冰的数字。

---

## 6. 模块与仓库结构

```
docx-embed/
├── edp/                        # 核心增强模块（import 名 edp，发布名 docx-embed）
│   ├── __init__.py
│   ├── extractor.py            # OLE 检测与提取
│   ├── main_parser.py          # MarkItDown/MinerU/Docling/Pandoc/RAGFlow 主文档解析适配
│   ├── ragflow_parser.py       # RAGFlow Markdown-export style DOCX parser
│   ├── resource_preview.py     # 内嵌资源浅 preview 与 parse_status 生成
│   ├── recursive_parser.py     # 子文件分发与递归解析
│   ├── merger.py               # 结果拼装、行内链接与位置映射
│   ├── manifest_builder.py     # manifest.json 生成
│   ├── docx_notes.py           # 页眉页脚 / 脚注 / 批注等 notes 抽取
│   └── models.py               # 数据模型与 dataclass 定义
├── evaluation/                 # 独立评测工具包
│   ├── fact_card.py            # 信息点卡数据结构
│   ├── layer_one.py            # 第一层评测脚本
│   ├── layer_two.py            # 第二层评测脚本（Agent QA）
│   ├── layer_fidelity.py       # 分维度保真评测（14 方法 × 23 维度记分板）
│   └── run_eval.py             # 评测入口
├── tests/                      # 按测试维度组织的用例
│   ├── conftest.py             # 合成 DOCX/XLSX fixture 构造
│   ├── test_T11_embed_text.py
│   ├── test_T12_embed_image.py
│   ├── test_T13_embed_package.py
│   ├── test_cli_parser.py
│   ├── test_cli_pipeline.py
│   ├── test_document_assets_markitdown.py
│   ├── test_docx_notes.py
│   ├── test_layer_fidelity.py
│   ├── test_main_parser_{markitdown,mineru,docling,ragflow}.py
│   └── test_dependency_metadata.py
├── docs/
│   ├── DESIGN.md               # 本文档（含路线图）
│   └── *-docx-parsing-research.md  # RAGFlow/Dify/WeKnora 等开源框架调研
├── examples/
│   ├── run_pipeline.py         # 增强管线入口
│   └── run_parser.py           # 纯 parser 对照入口
├── scripts/                    # 合成 fixture 生成等脚本
├── README.md
└── pyproject.toml
```

---

## 演进路线图 (Roadmap)

> **当前进度**：V0.1（资源层）、V0.2（评测体系）已完成；近期聚焦 **V0.3 图片精细化处理**，随后 **V0.4 XLSX 深度解析**。

### V0.1 — 核心资源层闭环 ✅ 已完成

**目标**：证明“发现 → 提取 → 登记 → 锚定 → 浅 preview → 拼装”这条资源层链路在 DOCX 内嵌资源场景下稳定跑通。

**交付物**：
- `extractor.py`：从 DOCX 中稳定提取附件和图片，登记 relationship、anchor、hash、size、mime 等资源元数据。
- `resource_preview.py`：为 `txt/log/csv/xlsx` 生成浅 preview，为 Word 图表 / SmartArt 生成 OOXML 文本/数据浅 preview；PDF、DOCX、PPTX、未知二进制输出 `extracted_only` 状态并保留原文件链接。
- `merger.py`：将资源 preview 或原始文件挂载到父包正确位置，生成行内 Markdown 链接、子文件列表、位置映射表和 `embedded_resources.jsonl`。
- 测试 T11、T12、T13 全部通过。
- 第一组对比数据：基座框架 vs EDP 增强后的嵌入对象召回率。

---

### V0.2 — 评测体系与对标 ✅ 已完成

**目标**：让评测自动跑起来，用数据量化基座框架的缺失与本模块的增益。

**交付物**：
- `fact_card.py`、`layer_one.py`、`layer_fidelity.py`、`run_eval.py` 完工，基于人工标注的信息点卡跑出自动化报告。
- 完成 14 方法 × 23 维度的对标记分板：覆盖 MarkItDown / MinerU / Docling / Pandoc / RAGFlow，每种框架的 pipeline 增强版 vs pure 对照版（详见 [INSIGHTS.md](../INSIGHTS.md)）。
- 两篇分析文章：《为什么框架不做嵌入对象解析》、《为什么框架不愿过度还原 Word 保真》。

> 实际范围超出原始“MinerU vs EDP”计划，扩展到 5 类解析器、14 个方法变体的对标。

---

### V0.3 — DOCX 图片理解（当前）

**目标**：把 DOCX 内图片从“提取 + 链接”升级到“可理解”——给每张图补足元数据、文字描述、分类与图注锚定，让下游 Agent 不看像素也能消费。

**交付物**：
- 图片元数据：格式、宽高、文件大小、hash、原始 relationship 与正文 anchor，写入 `embedded_resources.jsonl` / manifest。
- 图片分类：照片 / 图表 / 流程图 / 扫描件 / 公式截图，按类型决定后续处理路径。
- 描述生成：调视觉模型为每张图生成 `image_NN_desc.txt`（DESIGN 包结构里早有占位，此处落实）；无 VLM 配置时降级为只登记，保持可选。
- 可选 OCR：对扫描件 / 截图抽文字。
- Caption / 图注锚定：把“图N”样式段落绑到对应图，写入 manifest，避免图与图注错位。
- 评测：补图片描述召回、OCR 文本召回、图注锚定准确率维度。
- 配套文章：《让 Agent 看懂图：DOCX 图片的精细化处理》。

**不做什么**：
- 不做版面复原式重排（图片仍在原锚点位置，不重排版）。
- 不强制依赖 VLM；无 key 时优雅降级。

---

### V0.4 — XLSX 深度解析（下一步）

**目标**：把 XLSX 从“浅 preview（sheet 名 + 前几行）”升级到“结构化完整解析”——多 sheet、合并单元格、公式、表头识别，输出可被 Agent 直接查询的结构化数据。

**交付物**：
- `recursive_parser.py` 落地 XLSX 深解析：全 sheet 内容、合并单元格还原、公式值 / 字符串值区分、表头推断。
- 结构化 `preview.json`：sheet 列表 + 每个 sheet 的行列范围、合并区、表头、数据类型。
- 大表 / 宽表分页与截断策略，避免 preview 膨胀。
- 评测：补 XLSX 单元格准确率、sheet 召回、合并单元格还原率维度。
- 配套文章：《嵌入表格的深度解析：从 sheet 名到结构化数据》。

**不做什么**：
- 不做“排版型 Excel”稳健解析（留到 V3.0）。
- 不做 XLSX 内嵌图表数据提取（留到 V1.5）。

---

### V1.0 — DOCX 生产化收口

**目标**：把 DOCX 资源层打磨到更适合开源用户直接采用：稳定 CLI、目录级处理、清晰错误报告和可复现评测。

**交付物**：
- 正式 CLI：`docx-embed convert input.docx output/` 与 `docx-embed convert-tree docs/ output/`。
- 目录级处理：递归扫描 `.docx`，跳过 `~$*.docx`，保留相对输出路径。
- 资源安全策略：递归深度上限、循环引用检测、大小限制、路径穿越防护。
- 对嵌入 PDF / PPTX：继续保留、登记、链接与 hash 校验；不把它们扩展为独立 PDF/PPT 解析器。
- GitHub Release v1.0，正式发布。

---

### V1.5 — 复杂内容增强

**目标**：攻克复杂表格、图表、公式的非标准解析。

**交付物**：
- 表格增强模块：合并单元格拍平、嵌套表拆分。
- 图表数据提取模块：基于公式引用的图表数据源提取（含 XLSX 内嵌图表）。
- 公式解释工具（Agent 工具），用 LLM 解释公式意图。
- 评测覆盖 T03-T10（表格、公式、图表相关）。
- 第二篇文章：《从表格到图表：复杂文档的结构化增强之路》。

---

### V2.0 — 规模化与生态集成

**目标**：从单文档走向目录级处理，打通主流 AI 框架。

**交付物**：
- 批量处理能力：一次解析整个目录树。
- 与 LlamaIndex / LangChain 的集成示例。
- 评测覆盖 T14-T23（页眉页脚、脚注、批注、水印等）。
- 第三篇文章：《企业级文档 Agent 的工程化落地全链路》。

---

### V3.0 — 智能解析（远期愿景）

**方向**：
- LLM 辅助的异形表格区域识别与半自动标注工具。
- 基于人工纠正标注的微调数据集构建。
- “排版型 Excel” 的稳健解析方案。

---
