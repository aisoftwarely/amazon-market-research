# amazon-market-research

**亚马逊产品市场调研 Skill** —— 给定一个产品关键词，指导 agent 完成「卖家精灵数据下载 → 数据分析 → Excel 9 表调研报告生成 → 验证交付」的完整工作流。

专为 AI agent 设计（TRAE / Claude Code / Codex / WorkBuddy / Cursor 等），核心能力：

- **数据获取指引**：卖家精灵 4 类标准导出的精确下载路径与核对清单（用户手动下载，无需 API）
- **脚本化分析**：竞品统计 / 价格带 / 品牌集中度 / 评论门槛 / 新品友好度 / 关键词分层 / 季节性峰谷比 / 竞品去重
- **XML 直改写入引擎**：在保留模板样式、WPS 嵌入式单元格图片（DISPIMG）、图表、超链接的前提下填写任意 Excel 模板——绕过 openpyxl 保存丢图的问题
- **自动验证**：全量 XML 语法校验 + 单元格级回读比对 + 旧主题残留扫描

## 快速开始（agent 视角）

```
用户: 帮我调研 "dog harness" 的市场，模板在这个文件夹

agent: 1. 读 SKILL.md，按 6 阶段工作流执行
       2. 指导用户下载 Product导出/关键词挖掘/历史趋势/市场报告PDF
       3. 运行 scripts/ 下的分析脚本
       4. 综合数据撰写 content_data.py（分析结论）
       5. build_report.py 生成报告，verify_report.py 验证
       6. 交付 9 表调研报告 xlsx
```

## 安装

### TRAE / Claude Code（原生 skill 支持）

把本目录复制到项目技能目录：

```bash
# TRAE (项目级)
<project>/.trae/skills/amazon-market-research/

# Claude Code (用户级)
~/.claude/skills/amazon-market-research/
```

### Codex / 其他 agent

在 `AGENTS.md`（或系统提示可读的入口文件）中加入：

```markdown
调研任务请先阅读 amazon-market-research/SKILL.md 并按其工作流执行。
```

或直接把 SKILL.md 全文作为上下文提供。

## 依赖

```bash
pip install openpyxl        # 必装: Excel读写/验证
pip install pypdfium2       # 仅处理PDF市场报告时需要(PyMuPDF亦可)
```

Python ≥ 3.8。Windows PowerShell 5 环境注意事项见 SKILL.md「环境注意」。

## 目录结构

```
├── SKILL.md                     # 主工作流(agent入口, 必读)
├── scripts/
│   ├── inspect_template.py      # Excel模板结构侦察
│   ├── analyze_products.py      # 竞品导出→市场统计JSON
│   ├── analyze_keywords.py      # 关键词挖掘+历史趋势→分层JSON
│   ├── extract_pdf_report.py    # PDF报告→分段图片(视觉OCR)
│   ├── build_report.py          # XML直改写入引擎(核心)
│   └── verify_report.py         # 输出完整性验证
├── references/
│   ├── data-collection.md      # 卖家精灵4类数据下载指南
│   ├── analysis-guide.md       # 指标口径与分析方法论
│   ├── template-mapping.md     # 9表调研模板单元格映射
│   └── xml-surgery.md          # XML直改原理与避坑
└── assets/
    └── content_data_template.py # 内容映射配置骨架
```

## 输入 / 输出

**输入**（用户提供）：
- 产品关键词（如 `bride and groom socks`）
- 9 表调研模板 `.xlsx`
- 卖家精灵 4 类导出（下载方法见 references/data-collection.md）

**输出**：
- `{产品名}市场分析_新.xlsx`：9 表完整调研报告（调研基础信息/产品画像/市场分析/竞品深度分析/竞品数据/关键词流量分析/供应链成本与利润测算/可行性评估/合规与风险评估）

## License

MIT
