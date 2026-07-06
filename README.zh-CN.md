# Embedded Document Parser (EDP)

一个预处理增强层：在主解析器运行**之前**从 DOCX 中回收嵌入的 OLE 附件（XLSX / PDF / DOCX / PPT），再把它们以行内链接的形式回挂到一个干净、可审计的文档包里。

> English: [README.md](README.md)

---

## 为什么需要 EDP

真实世界的 DOCX 在 `word/embeddings/` 里携带嵌入附件——一个 Excel 模型、一份 PDF 摘要、一个嵌套 DOCX。我们测试的所有主流解析器（MarkItDown、MinerU、Docling、Pandoc、RAGFlow、Dify、WeKnora）都**静默丢弃**它们：附件从输出里消失，没有报错、没有警告、没有链接。

EDP 只补这一块短板。它是一个**资源层**，不是又一个解析器：提取嵌入对象、登记到 manifest（hash / mime / anchor / parse_status）、生成浅 preview、在原始锚点位置回挂。主解析器（MarkItDown / MinerU / Docling / Pandoc / RAGFlow）是**可替换组件**，只负责解析清理后的主文档正文。EDP 不重写主解析器的 markdown。

---

## 核心结论：所有 pure 框架都丢嵌入附件

在我们的测试语料（`llm-social-impact-fixture.docx`，5 个嵌入附件：txt / csv / xlsx / pdf / docx）上：

| 档位 | `embedded_object_recall` |
|---|:---:|
| 所有独立框架（`pure-*`） | **0.0000** |
| 所有 EDP 增强管线（`pipeline-*`） | **1.0000** |

嵌入附件是 EDP 补上的决定性缺口。新版评测还覆盖复选框、图表文本、SmartArt 文本、嵌套表、嵌套表内资产、公式、编号、超链接、TOC 和可移植性。完整 14 方法 × 23 维度记分板：[INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)。

---

## EDP 做了什么

- **提取**：用 `python-docx` + `olefile` / `oleobj` 从 DOCX 提取嵌入 OLE 对象，包括 `word/embeddings/` 下未被任何 OLE relationship 引用的孤儿文件。
- **登记**：每条资源写入 `embedded_resources.jsonl`，含 `sha256` / mime / anchor / `parse_status`（`shallow_preview` / `extracted_only` / `preserved_image`）。
- **浅 preview**：白名单类型生成预览——`txt` / `csv` → 文本块；`xlsx` → 按 sheet 的 markdown 表格 + 带类型的 `preview.json`；Word 图表 / SmartArt → OOXML 文本/数据浅预览。PDF / DOCX / PPTX / 未知二进制只保留并链接，不生成 preview。
- **回挂**：在 `content.md` 的原始锚点位置插入行内 markdown 链接。
- **产出**：自包含、可审计的文档包 `{doc_id}/{raw,structured,manifest.json}`。

---

## 快速开始

运行前先确认输入是新版 `.docx` 文件。本项目不直接支持旧版 `.doc`，请先另存或转换为 `.docx`。输出目录不要直接写回源文档目录，并跳过 `~$report.docx` 这类 Office 临时锁文件。

```bash
# 转换单个 DOCX 文件，使用默认管线
uv run examples/run_pipeline.py input.docx output_package

# 切换主解析器
uv run examples/run_pipeline.py input.docx output_package --main-parser pandoc

# 允许 EDP 把 Ole10Native .bin 外壳拆成真实载荷
uv run examples/run_pipeline.py input.docx output_package --unsafe-unwrap-embedded

# 纯 parser 对照：不做附件提取、不生成 manifest，仅用于对比
uv run examples/run_parser.py pandoc input.docx output_pandoc
```

递归转换某个目录下的全部 `.docx` 文件（包含嵌套目录），并在输出目录中保留相对路径：

```bash
SRC_DIR="input_docs"
OUT_DIR="output/packages"

find "$SRC_DIR" -type f -iname "*.docx" ! -name "~$*" -print0 |
while IFS= read -r -d "" docx; do
  rel="${docx#$SRC_DIR/}"
  out="$OUT_DIR/${rel%.docx}"
  mkdir -p "$(dirname "$out")"
  echo "Converting: $docx -> $out"
  uv run examples/run_pipeline.py "$docx" "$out" --main-parser markitdown --unsafe-unwrap-embedded
done
```

只有在你希望把 Ole10Native `.bin` 外壳拆成真实载荷时，才使用 `--unsafe-unwrap-embedded`。如果要公开数据集，发布前请人工检查生成的 `manifest.json`。

`--main-parser` 取值：`markitdown`（默认）、`mineru-pipeline`、`mineru-vlm-engine`、`mineru-hybrid-engine`、`docling`、`pandoc`、`ragflow`（`mineru` 是 `mineru-pipeline` 的兼容别名）。`ragflow` 使用 Markdown export 路径。

### 环境变量

| 解析器 | 必需环境变量 | 说明 |
|---|---|---|
| `markitdown` | 无 | 本地，无需 API key。稳健默认。 |
| `pandoc` | 无 | 需系统 PATH 上有 `pandoc` 可执行。 |
| `ragflow` | 无 | 本地 mammoth→markdownify export-style parser，不调 RAGFlow 服务。 |
| `mineru-pipeline` / `mineru-vlm-engine` / `mineru-hybrid-engine` | `MINERU_API_KEY`、`MINERU_BASE_URL` | 黑盒 HTTP 服务。兼容别名：`MINERU_AUTHORIZATION`、`DOCUMENT_CONVERTER_AUTHORIZATION`、`MINERU_API_ENDPOINT`、`MINERU_REQUEST_TIMEOUT`、`MINERU_BACKEND`。 |
| `docling` | `DOCLING_AUTHORIZATION`、`DOCLING_API_BASE` | 本地 Docling Serve（默认 `http://127.0.0.1:5001`）。兼容别名：`DOCLING_API_ENDPOINT`、`DOCLING_REQUEST_TIMEOUT`。 |

---

## 选型指南

没有哪个解析器在所有维度上都赢。按下游需求选。

### 按需求

| 你的需求 | 推荐 | 原因 |
|---|---|---|
| 回收嵌入附件 | EDP `pipeline-*`（任意） | 只有 EDP = `1.0`；所有 `pure-*` = `0.0` |
| 最高整体保真 | `pipeline-pandoc` | `doc_score` 0.9790；本 fixture 综合最强 |
| 长 / 宽 / 跨页表格 | `pipeline-docling` | 326 行双栏个案：`table_cell_match` 1.0000 |
| Word 图表标签 | 任意 EDP `pipeline-*` | EDP 抽取 chart OOXML 到 `structured/resources/chart_*` |
| SmartArt 文本 | 任意 EDP `pipeline-*` | EDP 抽取 diagram OOXML 到 `structured/resources/diagram_*` |
| 嵌套表 | 本 fixture 上为 `pipeline-pandoc` | 只有 Pandoc 还原了抽样嵌套表结构 |
| RAG chunking（干净语义文本） | `ragflow` / `WeKnora` / `Dify` / `mineru` | 需在你的文档上检查公式、链接和编号 |

### 按内容

| 你的 DOCX 内容 | 推荐 |
|---|---|
| 简单文本 + 普通图片 | `pipeline-markitdown` 或其他本地解析器通常足够 |
| 含嵌入 Excel / PDF / DOCX 附件 | **必须**用 EDP `pipeline-*` |
| 长 / 宽 / 跨页表格 | `pipeline-docling` |
| Word 图表 / SmartArt / 嵌套表 | 图表 + SmartArt：任意 EDP `pipeline-*`；抽样嵌套表：`pipeline-pandoc` |
| 重公式 / TOC / 编号文档 | 在你的 fixture 上跑评测；这些维度已实现，但当前公开 GT 未填充 |

**默认推荐：** `pipeline-markitdown`（本地、无需 API key、综合稳健）。当保真 / 编号 / 表格行编号重要且装了 `pandoc` 时升级到 `pipeline-pandoc`。表格密集文档用 `pipeline-docling`。**含附件的 DOCX 永远不要用 `pure-*` 交付**——附件会消失。

完整决策矩阵与各框架取舍：[INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)。

---

## 文档包输出

```
{doc_id}/
├── raw/                  # original.docx + 提取出的 embedded/ 二进制
├── structured/           # content.md、child_files.md、position_map.csv、
│                         # embedded_resources.jsonl、assets/、resources/
└── manifest.json         # 文档包清单（parse_status、content_map、…）
```

`content.md` 保持干净——只在原始锚点位置有行内 markdown 链接。完整包结构与数据结构：[docs/DESIGN.md](docs/DESIGN.md)。

---

## 当前状态

**V0.1 — 资源层已交付。** 「发现 → 提取 → 登记 → 锚定 → 浅 preview → 拼装」链路在 DOCX 上稳定跑通。`txt` / `csv` / `xlsx` 生成浅 preview；Word 图表 / SmartArt 生成 OOXML 文本/数据浅预览；PDF / DOCX / PPTX / 未知二进制保留并链接（`extracted_only`）。测试 T11 / T12 / T13 绿。

**V0.2 — 评测体系已交付。** 14 方法 × 23 维度对标记分板，覆盖 MarkItDown / MinerU / Docling / Pandoc / RAGFlow（pipeline 增强版 vs pure 对照版）。详见 [INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)。

**下一步：** 补 DOCX 图片理解能力：图片元数据、可选 OCR/VLM 描述、图片分类、图注锚定。PDF / PPTX 嵌入文件继续作为资源保留和链接，不再作为独立解析器路线图主线。

---

## 文档

- **[INSIGHTS.zh-CN.md](INSIGHTS.zh-CN.md)** — 简洁 14×23 记分板、TL DR 和 parser 选型指南。
- **[docs/DESIGN.md](docs/DESIGN.md)** — v1.0 技术设计（流水线、包结构、数据结构、评测体系、模块布局、路线图）。
- **[docs/examples/converted-markdown/](docs/examples/converted-markdown/)** — 可直接在 GitHub 阅读的精选转换后 Markdown。
- **各框架调研笔记**（中文）：[RAGFlow](docs/ragflow-docx-parsing-research.md) · [Dify](docs/dify-docx-parsing-research.md) · [WeKnora](docs/weknora-docx-parsing-research.md) · [MinerU backend](docs/mineru-backend-evaluation.md) · [评测报告](docs/docx-markdown-evaluation-report.md)

---

## 许可证

MIT — 允许商用。详见 [LICENSE](LICENSE)。

English: [README.md](README.md) | [INSIGHTS.md](INSIGHTS.md)
