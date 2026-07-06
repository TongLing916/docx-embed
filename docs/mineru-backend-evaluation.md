# MinerU Backend 评测口径

> 调研对象：`docx-embed` 当前 MinerU 接入与评测脚本
> 关注点：MinerU `pipeline`、`vlm-engine`、`hybrid-engine` 三种 backend 的独立评测
> 调整日期：2026-07-01

## 1. 背景

MinerU 的 `/file_parse` 接口通过 multipart 字段 `backend` 选择解析后端。EDP 现在把一个笼统的 `mineru` parser 拆成三个显式变体，避免评测结果里混用不同 MinerU backend 的输出。

保留兼容别名：

| EDP parser 名 | MinerU backend | 用途 |
|---|---|---|
| `mineru` | `pipeline` | 旧命令兼容，不建议用于新评测命名 |
| `mineru-pipeline` | `pipeline` | pipeline backend 的显式评测名 |
| `mineru-vlm-engine` | `vlm-engine` | VLM backend 的显式评测名 |
| `mineru-hybrid-engine` | `hybrid-engine` | hybrid backend 的显式评测名 |

`MINERU_BACKEND` 环境变量仍然拥有最高优先级：如果设置了该变量，请求体会使用它覆盖 parser 名推导出的 backend。正式评测时应避免设置 `MINERU_BACKEND`，否则 method 名和实际 backend 可能不一致。

## 2. 单文档运行

增强管线评测会先做 EDP 自研 DOCX 资源提取、资源 preview 和回挂，再把清理后的主 DOCX 交给指定主解析器：

```bash
export MINERU_API_KEY='<api key>'
export MINERU_BASE_URL='http://127.0.0.1:8000'

uv run examples/run_pipeline.py input.docx output/mineru-pipeline --main-parser mineru-pipeline
uv run examples/run_pipeline.py input.docx output/mineru-vlm --main-parser mineru-vlm-engine
uv run examples/run_pipeline.py input.docx output/mineru-hybrid --main-parser mineru-hybrid-engine
```

纯 parser baseline 不做附件提取、子文件列表、位置映射和 EDP 回挂，只评估 MinerU 对原始 DOCX 的直接解析结果：

```bash
uv run examples/run_parser.py mineru-pipeline input.docx output/pure-mineru-pipeline
uv run examples/run_parser.py mineru-vlm-engine input.docx output/pure-mineru-vlm
uv run examples/run_parser.py mineru-hybrid-engine input.docx output/pure-mineru-hybrid
```

## 3. 批量评测 method 命名

`scripts/batch_run_all.sh` 会为每个 `data/*.docx` 生成独立输出目录：

| method | 入口 | parser 参数 |
|---|---|---|
| `pipeline-mineru-pipeline` | `examples/run_pipeline.py` | `--main-parser mineru-pipeline` |
| `pipeline-mineru-vlm-engine` | `examples/run_pipeline.py` | `--main-parser mineru-vlm-engine` |
| `pipeline-mineru-hybrid-engine` | `examples/run_pipeline.py` | `--main-parser mineru-hybrid-engine` |
| `pure-mineru-pipeline` | `examples/run_parser.py` | `mineru-pipeline` |
| `pure-mineru-vlm-engine` | `examples/run_parser.py` | `mineru-vlm-engine` |
| `pure-mineru-hybrid-engine` | `examples/run_parser.py` | `mineru-hybrid-engine` |

批量产物路径形如：

```text
output_batch/<doc_id>/pipeline-mineru-pipeline/
output_batch/<doc_id>/pipeline-mineru-vlm-engine/
output_batch/<doc_id>/pipeline-mineru-hybrid-engine/
output_batch/<doc_id>/pure-mineru-pipeline/
output_batch/<doc_id>/pure-mineru-vlm-engine/
output_batch/<doc_id>/pure-mineru-hybrid-engine/
```

`evaluation/run_eval.py` 按同一组 method 固定顺序读取这些目录，并输出：

```bash
uv run python evaluation/run_eval.py
```

## 4. 结果解释

对比时应同时看两组结果：

- `pipeline-mineru-*`：衡量 MinerU backend 作为 EDP 增强管线主文档解析器时的效果，嵌入附件、图片登记和 anchor 回挂仍由 EDP 负责。
- `pure-mineru-*`：衡量 MinerU backend 直接解析原始 DOCX 的 baseline，适合判断 MinerU 自身对正文、表格、图片链接的覆盖能力。

两组结果不可互相替代。前者回答“接入 EDP 后的整体包质量”，后者回答“MinerU 单独解析原文档的质量”。
