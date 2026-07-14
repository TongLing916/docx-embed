# Docling XLSX 后端性能退化源码分析：合并单元格密集型工作簿

> 调研对象：`docling/backend/msexcel_backend.py`（docling-serve v1.25.0，容器内路径
> `/opt/app-root/lib/python3.12/site-packages/docling/backend/msexcel_backend.py`）
>
> 调研目的：定位 Docling 在处理某些 XLSX 时从秒级退化到数十分钟级的根因，评估是否
> 为框架天生缺陷，并为 `docx-embed` 的 XLSX 能力设计与后端选型提供性能避坑参考。

---

## 1. 现象与反证

同一套 Docling 服务（BII 内网 docling-serve-gpu，A10），同一类操作（xlsx → md，
`do_ocr=false`），两个真实文件的耗时天差地别：

| 文件 | 体积 | 单元格数 | 合并区域数 | 内嵌图片 | Docling 耗时 |
|------|------|----------|-----------|----------|-------------|
| `Orders.xlsx`（扁平流水表） | 641 KB | ~9.1 万 | **0** | 0 | **4.5 s** |
| `19号线电动客车检修规程 A3版.xlsx`（A3 规程文档） | 160 KB | ~2.16 万（更少） | **3727** | 29 | **~40 min** |

关键反证：`Orders.xlsx` 的单元格数是 `19号线` 的 **4 倍**，体积是其 **4 倍**，耗时却是
其 **1/500**。所以退化**与文件大小、单元格数量无关**，而是 `19号线` 这份 A3 规程文档
的结构特征（数千个合并区域 + 内嵌图片）踩中了 Docling 的某条慢路径。这不是"XLSX 天生
慢"，也不是"Docling 天生慢"，是**特定结构下的算法退化**。

> 现场来源：BII docling-serve GPU 池在生产中因 `default` 租户反复投递 `19号线` 这类
> A3 规程 xlsx（单文件 40 min）把单 worker 拖死，导致对外全量 504。

---

## 2. 源码定位：慢在 `has_content` 的合并区域线性扫描

`MsExcelDocumentBackend` 的表格识别用 **BFS flood fill** 找连通单元格块。瓶颈不在 BFS
本身，而在 BFS 每一步判断邻居是否有内容时调用的 `has_content`：

```python
# docling/backend/msexcel_backend.py，_find_table_bounds 内
def has_content(r, c):
    if r < 0 or c < 0 or r >= max_row or c >= max_col:
        return False
    # 1. 直接取值
    cell = sheet.cell(row=r + 1, column=c + 1)
    if cell.value is not None:
        return True
    # 2. 遍历【全部】合并区域做包含判断  ← 瓶颈
    for mr in sheet.merged_cells.ranges:
        if cell.coordinate in mr:
            return True
    return False
```

`has_content` 对**每个待判断的单元格**都**线性遍历整个 `sheet.merged_cells.ranges`
列表**，逐一做 `cell.coordinate in mr` 判断。而 `_find_table_bounds` 的 BFS 在每个
已入队单元格的 4 个方向上都要调 `has_content`（受 `gap_tolerance` 控制，默认 0 即探 1
步），Phase 2 提取数据时又对整个 bbox 的每个单元格再扫一遍合并区域以算 `row_span /
col_span`。

因此单次表格识别的合并区域查询总次数约为：

```
O(被遍历的单元格数 × 合并区域数)
```

- `Orders.xlsx`：合并区域 = 0 → 这一项查询次数 = 0 → 秒级完成。
- `19号线`：合并区域 = 3727，被遍历单元格数万级 → 查询次数达 **2400 万级**。

### 本地最小复现（只跑一遍 `has_content` 等价扫描）

用 openpyxl 复刻"对每个非空单元格线性扫描全部合并区域"这一步，单遍耗时：

```
Orders.xlsx:    540 非空 × 0     合并 =          0 次扫描 → 0.00 s
19号线 A3版:   6570 非空 × 3727  合并 = 24,486,390 次扫描 → 141.55 s（仅一遍）
```

Docling 实际要在 BFS 每个邻居方向 + Phase 2 整个 bbox（491×44）上各扫一遍，是上面这个
数字的**数倍** → 估算 ~2400 s，与生产观测的 ~40 min 吻合。

---

## 3. 几个澄清（排除干扰项）

- **图片不是瓶颈**：29 张内嵌图只是 `PIL.Image.open` 读一遍进 `PictureItem`，可忽略。
  `do_ocr=false` 时也不跑 OCR。
- **不走 ML 表格结构模型**：xlsx backend 直接从单元格 + 合并信息构造 `TableData`，
  `do_table_structure` 对 xlsx 不生效，没有 GPU 推理开销。
- **`gap_tolerance` 默认 0**：BFS 每方向只探 1 步，不影响"每步都线性扫合并区域"这一
  事实；调大 `gap_tolerance` 反而会增多 `has_content` 调用次数，可能更慢。
- **`load_workbook` 本身不慢**：openpyxl 加载这两个文件都在百毫秒级。
- **本质是数据结构选择问题，不是算法思路问题**：BFS flood fill 找连通区域本身合理，
  问题在于把合并区域存成线性 list 并反复全量扫描。若按行/列建 dict 或区间树做空间
  索引，`has_content` 可降到近 O(1)，整体回到 O(cells)。

---

## 4. 复杂度结论

| 维度 | Docling `msexcel_backend` 现状 |
|------|-------------------------------|
| 表格区域识别算法 | BFS flood fill（合理） |
| 合并区域查询 | 线性遍历 `merged_cells.ranges`（**退化点**） |
| 整体复杂度 | ≈ O(cells × merged_ranges) |
| 退化触发条件 | 合并区域数千个的"规程/版式类"xlsx |
| 退化量级 | 秒级 → 数十分钟级 |
| 是否天生缺陷 | 否，可修：合并区域加空间索引即可降到近 O(cells) |

**一句话**：Docling 处理 xlsx 慢，不是 xlsx 天生慢，也不是 Docling 天生慢，而是
`msexcel_backend` 把合并区域当线性 list 反复全量扫描，在"合并单元格密集型"工作簿上
从 O(cells) 退化到 O(cells × merged_ranges)。扁平流水表（无合并）依然秒级。

---

## 5. 对 docx-embed 的启示

1. **EDP 的 openpyxl 流水线天然规避了这条慢路径**：`edp/xlsx/parser.py` 直接按
   `data_only=True/False` 双视图遍历单元格（见 [xlsx-parsing-guide.md](xlsx-parsing-guide.md)），
   不依赖反复查询合并区域做连通性判断，因此在同类规程文档上保持秒级。
2. **若 EDP 未来要做"表格区域检测"，务必给合并区域建空间索引**（按行聚合的区间列表或
   区间树），不要重蹈 Docling 的线性扫描。
3. **选型参考**：对"规程/版式类、合并单元格密集"的 xlsx，应避免直接用 Docling 单跑；
   EDP pipeline 或 MinerU 是更稳的选择。Docling 适合结构相对扁平、合并区域较少的场景。
4. **运营侧启示**（供跑 Docling serve 的同学参考）：单 worker + 同步接口超时不取消
   底层 job 时，一个 40 min 的病态 xlsx 会拖死整个池子；根治需 ①orchestrator 加单任务
   超时 + 取消 ②readiness 改真跑 trivial 转换 ③多 worker 并发 ④上游对病态文件限流。
