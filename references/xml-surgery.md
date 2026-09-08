# XML 直改技术与避坑指南

build_report.py 引擎已内置本文全部要点。本篇解释原理 + 引擎未覆盖场景的手工处理方案。

## 为什么必须 XML 直改

WPS 的嵌入式单元格图片用 `=_xlfn.DISPIMG("ID_xxx",1)` 公式 + `xl/cellimages.xml` 实现。**openpyxl 保存时会丢弃 cellimages 及其 rels，导致所有单元格图片丢失**。因此：模板只能以 zip 形式解包 → 修改内部 XML → 重新打包。openpyxl 仅用于**读取**（分析导出文件、验证输出）。

xlsx zip 结构：

```
[Content_Types].xml          内容类型注册(新增媒体需在此登记扩展名)
_rels/.rels                  包级关系
docProps/app.xml             工作表名列表(改名需同步)
xl/workbook.xml              sheet名 -> rId 映射
xl/_rels/workbook.xml.rels   rId -> worksheets/sheetN.xml
xl/worksheets/sheetN.xml     单元格数据
xl/worksheets/_rels/sheetN.xml.rels  表级关系(超链接/图片/图表)
xl/sharedStrings.xml         共享字符串池
xl/cellimages.xml            WPS单元格图片定义
xl/media/*                   图片文件
xl/charts/chartN.xml         图表(含数据缓存)
xl/externalLinks/*           外部引用缓存(图表链接外部表时)
xl/drawings/drawingN.xml     浮动图片锚点
```

## 单元格存储格式

```xml
<c r="C7" s="11" t="s"><v>42</v></c>   <!-- s=样式索引, t="s"字符串池引用 -->
<c r="C11" s="11"><v>861</v></c>        <!-- 数字直接存 -->
```

**样式保留规则**：改值时读取原单元格的 `s=` 属性并保留；新插入单元格用 DONOR_STYLES 指定同表样式来源；写空值也要带 `s=`（否则模板底纹/边框消失）。

## sharedStrings —— 最危险的坑

原模板每个 `<si>` 条目内容自带 `<t>` 包裹（`<si><t>文本</t></si>`）。重建字符串池时若给原条目再包一层 `<t>`，会产生非法嵌套，**所有引用原条目的单元格在 openpyxl/WPS 中显示为空串**——这是最隐蔽的破坏性 bug。

正确做法（引擎已内置 `N_ORIG` 边界）：
- 原 `N_ORIG` 条目：原样保留
- 新增条目：`esc(text)` 后包 `<t>`；首尾有空白时加 `xml:space="preserve"`

## 图表缓存更新

图表 XML 内含数据缓存：`<c:cat><c:strRef><c:f>引用</c:f><c:strCache><c:ptCount/><c:pt>...` 。

**只替换 `<c:pt>` 序列与 `<c:ptCount>`，绝不能整段重建 cat/val 区块**——手工拼接极易丢失 `<c:f>` 引用或闭合标签，导致图表 XML 语法错误、openpyxl 直接抛 `XMLSyntaxError: expected '>'`。引擎的 `_rebuild_pts` 用「先删 pt、再在 ptCount 处注入新 pt」的两步法保证结构不变。

图表若引用外部链接（`[1]共性差评统计!$B$3`），externalLink1.xml 内有镜像缓存，需用 `RAW_PATCHES` 同步替换其中的 `<v>` 文本与数值，否则 WPS 打开时可能回显旧数据。

## 图片锚点删除

模板旧主题截图是 drawing 内的锚点（`<xdr:twoCellAnchor>` 或 `<xdr:oneCellAnchor>`），每个锚点通过 `r:embed="rIdN"` 关联媒体。删除用「tempered regex」：匹配以 `r:embed="rIdN"` 为中心、不跨越其他锚点闭合标签的整段（引擎 `apply_media` 已实现）。

删除后该 rId 从 rels 中失去引用属正常（rel 条目保留不影响打开）。

新增图片走 `MEDIA_ADD`（zip 内新路径 + 本地文件），再按需用 `RAW_PATCHES` 重写对应 drawing 的 rels 指向新媒体（整文件替换 rels 内容，见下）。

## 手工 RAW_PATCHES 配方（引擎未覆盖场景）

```python
RAW_PATCHES = {
    # 整体重写 drawing rels: 保留图表关系, 图片关系指向新媒体
    'xl/drawings/_rels/drawing2.xml.rels': [
        (OLD_ENTIRE_RELS_XML, NEW_ENTIRE_RELS_XML),
    ],
    # externalLink 缓存数值: (含上下文的旧串, 新串)
    'xl/externalLinks/externalLink1.xml': [
        ('<v>旧标签</v>', '<v>新标签</v>'),
    ],
}
```

替换串必须**先在解包 XML 中精确定位再复制进配置**（可用 python 打印目标片段），避免凭记忆编写导致不命中——引擎对未命中的 patch 会打印警告，务必检查构建日志。

## 工作表改名

需同时改 `xl/workbook.xml` 与 `docProps/app.xml` 中的表名（引擎已处理）。改前确认无 definedNames/calcChain 引用旧名（`inspect_template.py` 可侦察，一般调研模板无）。

## 数据表重建要点

1. 表头行与数据行的列样式从模板原 row1/2/3 逐列捕获
2. 重建后更新 `<dimension ref="A1:XX{rows}"/>`
3. 超链接：sheetN.xml 的 `<hyperlinks>` 块与 `sheetN.xml.rels` 必须成对重建；`display` 属性截断 255 字符并转义
4. 导出中的数字型字符串（'861.0'）要转成数字类型写入，否则 Excel 存为文本左对齐

## 验证清单（verify_report.py 自动执行）

1. `zip.testzip()` 完整性
2. 全部 XML/rels minidom 语法校验（**写出前**也校验一次）
3. openpyxl 打开不抛异常
4. 每个 SHEET_TEXT 单元格值精确回读比对
5. 数据表行列数 = 导出文件
6. 旧主题词残留扫描（数据表除外；有意的对比表述需人工确认）
7. 媒体/图表文件数量与模板一致（或符合预期增减）

## 已知良性现象（不是错误，不要"修复"）

- openpyxl 读 WPS 图表时 `Unable to read chart rId1 ... ExternalData.id should be str but value is NoneType` 警告——模板自带的外部链接图表特性
- 数据表 A 列（DISPIMG 图片列）重建后为空——正常，主图链接在「商品主图」列
- 保留的旧 `<t>` 条目中偶见富文本 run（`<r>`）——原样保留即可
