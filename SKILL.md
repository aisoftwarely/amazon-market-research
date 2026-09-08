---
name: amazon-market-research
description: "亚马逊市场调研工作流：指导从卖家精灵(SellerSprite)下载数据→分析→用XML直改技术填写Excel模板生成9表调研报告(保留图片/图表/样式)。当用户提供产品关键词要求做亚马逊市场调研、选品分析、类目分析或填写调研模板时使用。 Amazon market research: SellerSprite data → Excel report."
---

# 亚马逊市场调研工作流（卖家精灵 → Excel 9表调研报告）

## 用途

给定一个产品关键词（如 `bride and groom socks`），完成从数据到成品报告的完整调研：

1. 指导用户从**卖家精灵（SellerSprite）**下载 4 类标准数据
2. 脚本化分析：市场容量 / 价格带 / 季节性 / 关键词分层 / 竞品对标 / 利润测算
3. 将分析结果填入标准 **9 表 Excel 调研模板**——通过 XML 直改保留模板原有样式、嵌入图片（WPS DISPIMG）、图表、超链接
4. 自动验证输出文件完整性

## 前置条件

- 用户有卖家精灵网页版账号（数据由用户手动下载，无需 API/MCP）
- Python 3.8+；`openpyxl` 必装；`pypdfium2`（或 `PyMuPDF`）仅处理 PDF 市场报告时需要
- 用户提供 9 表调研模板 .xlsx（结构见 `references/template-mapping.md`）

## 环境注意（Windows / PowerShell 5）

- PowerShell 脚本执行通常被禁用（`about_Execution_Policies` 报错）→ 一律写 `.py` 文件后 `python xxx.py` 执行，不要用 `powershell -File`
- PowerShell 5 **不支持 `&&`**，串联命令用 `;`
- 含中文引号/复杂引号的 `python -c "..."` 内联代码极易引号错乱 → 也一律写 `.py` 文件
- 路径含空格时用双引号包裹

## 工作流（6 阶段）

### 阶段 0：确认输入

向用户确认三件事：
1. **调研关键词**（英文核心词，如 `bride and groom socks`）
2. **模板文件路径**（默认取工作区内最新的调研模板 xlsx）
3. **目标站点**（默认美国站 US）

在工作目录下创建 `_analysis/` 文件夹存放所有中间产物。

### 阶段 1：模板结构侦察

```bash
python <skill>/scripts/inspect_template.py "模板.xlsx" "_analysis/template_grid.txt"
```

通读输出文件，掌握：每个工作表的行列布局、标签单元格坐标、DISPIMG 嵌图位置、drawing 锚点、图表归属。据此 + `references/template-mapping.md` 建立"单元格 → 内容"映射计划。

### 阶段 2：指导用户下载数据

按 `references/data-collection.md` 指导用户从卖家精灵下载 **4 类数据**并放入工作目录：

| 数据 | 卖家精灵功能 | 文件名特征 |
|---|---|---|
| 竞品商品导出 | 商品搜索（关键词搜Top商品导出） | `Product-US-YYYY.MM-*.xlsx` |
| 关键词挖掘 | 关键词挖掘（核心词相关词） | `KeywordMining-US-*-YYYYMM-*.xlsx` |
| 关键词历史趋势 | 关键词历史趋势（选20-50个词） | `KeywordHistory-*(N)-US-*.xlsx` |
| 市场分析报告 | 市场分析（生成PDF报告） | `市场分析报告 - *.pdf` |

用户下载完成后核对文件存在与可读性再进入下一阶段。

### 阶段 3：数据分析

```bash
# 3.1 竞品导出分析(市场统计/价格带/品牌集中度/新品友好度/去重母体)
python <skill>/scripts/analyze_products.py "Product-US-*.xlsx" --out "_analysis/products.json" --top 7

# 3.2 关键词分析(挖掘+历史趋势join/季节性/峰谷比/竞争度分层)
python <skill>/scripts/analyze_keywords.py "KeywordMining-*.xlsx" "KeywordHistory-*.xlsx" --out "_analysis/keywords.json"

# 3.3 PDF市场报告转图分段(图片型PDF, 之后用 Read 工具逐段OCR)
python <skill>/scripts/extract_pdf_report.py "市场分析报告-*.pdf" --outdir "_analysis/segs" --segments 8
```

然后**亲自阅读**脚本输出的 JSON 摘要与 PDF 分段图片，结合 `references/analysis-guide.md` 的方法论完成人工判断：旺季窗口、切入价位、差异化方向、风险清单、利润模型。脚本给数字，**结论必须由你综合得出**。

### 阶段 4：编写内容映射 `content_data.py`

将全部分析结论写入 `_analysis/content_data.py`（骨架见 `assets/content_data_template.py`，完整字段说明见文件内注释）：

```python
SHEET_TEXT = {'调研基础信息': {'C7': '旺季描述...', 'D7': '详细论证...'}, ...}
DATA_SHEET = '202603竞品数据'          # 模板中原始数据表名
DATA_SHEET_RENAME = '202512竞品数据'   # 重命名为本次调研年月
CHART_UPDATES = {...}                  # 模板图表的分类/数值缓存更新
SHEET_LINKS = {'竞品深度分析': {'C3': 'https://www.amazon.com/dp/B0XXXX'}}  # 竞品超链接
DELETE_ANCHOR_RIDS = {...}             # 需删除的模板旧图锚点(可选)
MEDIA_ADD = {...}                      # 新增竞品图(可选)
```

### 阶段 5：构建报告（XML 直改）

```bash
python <skill>/scripts/build_report.py --template "模板.xlsx" --content "_analysis/content_data.py" --raw "Product-US-*.xlsx" --out "产品名市场分析_新.xlsx"
```

引擎自动完成：单元格写入(保留样式) / 数据表整表重建(含超链接) / 工作表改名 / 图表缓存更新 / sharedStrings 安全重建 / **全量 XML 语法校验**。

### 阶段 6：验证交付

```bash
python <skill>/scripts/verify_report.py --new "产品名市场分析_新.xlsx" --content "_analysis/content_data.py" --terms "旧主题词1,旧主题词2"
```

验证项：zip 完整性 / 全部 XML 合法 / 每个写入单元格值精确匹配 / 数据表行列数 / 旧主题残留扫描 / 图片图表保留。全绿后向用户交付输出文件并总结 9 表内容。

## 关键铁律（违反会导致文件损坏或数据丢失）

1. **绝不用 openpyxl 直接保存模板**——WPS 嵌入式单元格图片（DISPIMG）会被 openpyxl 丢弃。只能用 XML 直改（build_report.py 已内置）。
2. **sharedStrings 重建时原有条目原样保留**（它们自带 `<t>` 包裹，不能重复包裹，否则全部文本变空串）。引擎已内置 `N_ORIG` 边界处理，不要绕过引擎手改。
3. **写出 zip 前必须 minidom 校验所有 XML 部件**。引擎已内置。
4. openpyxl 读取时报 `Unable to read chart rId1 ... ExternalData` 警告属模板自带的良性警告，不是构建错误。
5. 数据表重建后必须同步更新 `<dimension>` 和该表的 `.rels` 超链接。
6. 每个写数字的单元格注意类型：int/float 走 `<v>`，字符串走 sharedStrings，`None` 生成空单元格但保留样式。

## 文件结构

```
amazon-market-research/
├── SKILL.md                        # 本文件(主工作流)
├── README.md                       # GitHub 发布说明/跨agent安装
├── scripts/
│   ├── inspect_template.py         # 阶段1 模板结构侦察
│   ├── analyze_products.py          # 阶段3 竞品导出分析
│   ├── analyze_keywords.py          # 阶段3 关键词挖掘+趋势分析
│   ├── extract_pdf_report.py        # 阶段3 PDF报告转图分段
│   ├── build_report.py              # 阶段5 XML直改写入引擎
│   └── verify_report.py            # 阶段6 输出验证
├── references/
│   ├── data-collection.md           # 卖家精灵4类数据下载指南
│   ├── analysis-guide.md            # 分析方法论与指标口径
│   ├── template-mapping.md          # 9表模板单元格映射表
│   └── xml-surgery.md              # XML直改技术细节与避坑
└── assets/
    └── content_data_template.py    # content_data.py 骨架模板
```

按需深入阅读 references，不必每次全读。
