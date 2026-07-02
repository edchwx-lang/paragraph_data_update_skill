# 段落数据自动更新 Skill

## 一、Skill 目标

本 Skill 用于将产业研究、政策报告、政府汇报材料中的旧数据自动更新到 2025 年至今或最新可获得口径。

核心能力：

1. 识别段落中 2025 年以前的数据；
2. 检索最新可替代数据；
3. 保持原文结构，替换部分加粗；
4. 找不到同口径数据时，小范围改写该句为最新能找到的数据表述；
5. 生成精简 Excel 数据替换来源表。

## 二、文件结构

```text
paragraph_data_update_skill/
├── SKILL.md
├── README.md
├── rules/
│   ├── data_identification_rules.md
│   ├── replacement_rules.md
│   ├── source_priority_rules.md
│   └── output_rules.md
├── prompts/
│   ├── default_trigger_examples.md
│   └── search_query_patterns.md
├── templates/
│   ├── 数据替换来源表模板.xlsx
│   └── data_replacement_records_example.json
├── scripts/
│   └── export_replacement_excel.py
└── examples/
    ├── sample_input.txt
    └── sample_output.md
```

## 三、推荐使用方式

用户输入：

```text
请更新以下段落中的2025年以前数据，更新后把更改部分加粗，并生成数据替换Excel。找不到同口径数据时，可以小范围改写该句为最新能找到的数据表述。

[粘贴段落]
```

模型输出：

1. 更新后段落；
2. 数据替换说明表；
3. Excel 文件。

## 四、Excel 表头

固定只保留 5 列：

| 序号 | 原文数据 | 替换数据 | 数据来源 | 来源URL |
|---|---|---|---|---|

## 五、替换边界

本 Skill 默认使用“严格替换模式”：

- 能只改数字就不改句子；
- 能只改分句就不改整句；
- 能只改一句就不改整段；
- 找不到同口径数据时，只允许对该句做小范围改写；
- 不得编造数据、来源或 URL。

## 六、脚本说明

`scripts/export_replacement_excel.py` 可将 JSON 记录导出为 `.xlsx` 文件。

示例：

```bash
python scripts/export_replacement_excel.py \
  --input templates/data_replacement_records_example.json \
  --output 数据替换来源表_示例.xlsx
```

JSON 格式：

```json
[
  {
    "序号": 1,
    "原文数据": "2023年我国某产业市场规模达到1200亿元",
    "替换数据": "2025年我国某产业市场规模达到1500亿元",
    "数据来源": "某行业协会报告",
    "来源URL": "https://example.com/report"
  }
]
```
