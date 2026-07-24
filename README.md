# Excel Logic Parser

将包含5列固定表头的Excel质控逻辑文档转换为结构化JSON，供后续AI编写Python脚本使用。

## 固定表头

| 列 | 表头 | 说明 |
|----|------|------|
| A | 字段名称 | 字段名，支持合并单元格 |
| B | 字段逻辑场景 | 中文描述，后续需转为Python逻辑 |
| C | 质检结果提示文案 | 提示词 |
| D | 质控解决方案 | 理想数据（只提示不修改 / 修改为对应数据） |
| E | 调整日志 | 删除/新增/调整逻辑标记，红色字体需特别关注 |

## 用法

```bash
python excel-logic-parser/scripts/parse_excel_logic.py <Excel文件路径> [输出JSON路径]
```

可选参数：
- `--sheet <名称>`：指定工作表（默认第一个）
- `--start-row <行号>`：数据起始行（默认第2行）
- `--end-row <行号>`：数据结束行（默认自动检测）

## 依赖

```
pip install openpyxl
```

## 目录结构

```
excel-logic-parser/
├── SKILL.md                        # 技能说明（触发词、工作流程、修改指南）
├── scripts/
│   └── parse_excel_logic.py        # 核心解析脚本
└── references/
    └── schema.md                   # JSON Schema与字段说明
```

## 修改指南

所有解析规则集中在脚本顶部的**配置区**：

| 需要修改 | 位置 | 说明 |
|---------|------|------|
| 列映射 | `COLUMN_MAP` | 调整A-E列对应关系 |
| 红色字体检测 | `RED_COLOR_PATTERNS` | 添加/修改颜色模式 |
| 调整日志解析 | `ACTION_PATTERNS` | 添加新的action类型 |
| 质控解决方案解析 | `parse_solution()` | 添加新的solution类型 |
