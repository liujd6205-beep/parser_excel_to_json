---
name: excel-logic-parser
description: "Excel质控逻辑文档解析器。将包含5列固定表头（字段名称、字段逻辑场景、质检结果提示文案、质控解决方案、调整日志）的Excel逻辑文档转换为结构化JSON，供后续AI编写Python脚本使用。触发词：解析逻辑文档、转换Excel逻辑、逻辑文档转JSON、质控逻辑解析、新生儿访视逻辑、RPA逻辑文档。当用户上传包含字段名称/字段逻辑场景/质检结果提示文案/质控解决方案/调整日志的Excel文件，或要求将逻辑文档转为JSON/脚本时触发此技能。"
agent_created: true
---

# Excel 质控逻辑文档解析器

## 概述

将包含5列固定表头的Excel质控逻辑文档解析为结构化JSON，输出给用户确认后供AI编写Python脚本。

## 固定表头说明

Excel文档包含以下5列固定表头：

| 列 | 表头 | 说明 |
|----|------|------|
| A | 字段名称 | 字段名，支持合并单元格（下方行继承上方字段名） |
| B | 字段逻辑场景 | 中文描述，后续需转为Python逻辑语言。**红色字体出现在此列** |
| C | 质检结果提示文案 | 提示词 |
| D | 质控解决方案 | 理想数据，分为"只提示不修改"和"修改为对应理想数据" |
| E | 调整日志 | 标记该逻辑是删除/新增/调整 |

## 关键规则

### 调整日志（E列）解析

- **删除逻辑** → 在该字段名称对应的函数方法上删除此逻辑
- **新增逻辑** → 增加此逻辑
- **逻辑调整** → 调整已有逻辑（如范围值变更）
- **删除访视方式为面对面** → 删除面对面访视相关逻辑

### 质控解决方案（D列）解析

- **只提示不修改** → ideal_data = null（仅提示，不修改数据）
- **清空** → ideal_data = "清空"（注意：是"清空"两个字本身，不是空字符串）
- **选xxx** → 选择固定值，ideal_data = "xxx"
- **添加xxx** → 添加内容，ideal_data = "xxx"
- **添加勾选"xxx"** → 勾选项
- **必须包含xxx** → 必须包含某内容
- **xxx随机取值** → 范围内随机取值，ideal_data = {"min": x, "max": y}
- **填xxx** → 填写固定文本
- **选择异常** → 选择"异常"选项
- 其他文本 → 直接作为填写值

### 字段逻辑场景（B列）红色字体

红色字体**只出现在B列（字段逻辑场景）**中，表示特殊条件，需特别关注。
- 部分特殊条目下方会有一行红色备注行，需提取并特别处理
- 支持 Rich Text 检测（同一单元格内混合黑白+红色文字）
- `logic_scene_is_red`：B列是否包含红色字体
- `logic_scene_red_text`：B列中红色字体的文本内容列表（用于提取红色备注）
- `action.is_red`：基于B列红色字体判断，是否需要特别关注

## 工作流程

### 第一步：解析Excel为JSON

运行脚本将Excel转换为JSON：

```bash
python {SKILL_ROOT}/scripts/parse_excel_logic.py <Excel文件路径> [输出JSON路径]
```

可选参数：
- `--sheet <名称>`：指定工作表（默认第一个）
- `--start-row <行号>`：数据起始行（默认第2行）
- `--end-row <行号>`：数据结束行（默认自动检测）

### 第二步：输出JSON给用户确认

将生成的JSON展示给用户，说明：
- 字段总数、逻辑条目总数
- 各action类型统计（delete_logic / add_logic / adjust_logic等）
- 各solution类型统计（prompt_only / modify_select / modify_random等）
- 红色字体标记条目数

等待用户确认JSON格式是否符合要求。

### 第三步：根据JSON编写Python脚本（用户确认后）

用户确认后，根据JSON中的逻辑条目编写Python脚本：
1. 按字段名分组，每个字段对应一个函数方法
2. 检查每条逻辑的action.type：
   - delete_logic → 删除该逻辑
   - add_logic → 新增该逻辑
   - adjust_logic → 调整该逻辑
3. 将logic_scene（中文描述）转为Python逻辑语言
4. 根据solution.type处理理想数据：
   - prompt_only → 只提示不修改
   - modify_select / modify_fill → 修改为ideal_data
   - modify_random → 在[min, max]范围内随机取值
   - modify_clear → 清空字段值
5. 特别关注 `logic_scene_is_red=true` 和 `logic_scene_red_text` 不为空的条目（红色字体标注的特殊条件）

## 修改指南

所有解析规则集中在脚本文件的**配置区**，修改后无需改其他代码：

| 需要修改 | 位置 | 说明 |
|---------|------|------|
| 列映射 | `COLUMN_MAP` | 调整A-E列对应的列号 |
| 数据起始行 | `DATA_START_ROW` | 跳过表头后的数据起始行 |
| 红色字体检测 | `RED_COLOR_PATTERNS` | 添加/修改需要检测的颜色模式 |
| 调整日志解析 | `ACTION_PATTERNS` | 添加新的action类型匹配规则 |
| 质控解决方案解析 | `parse_solution()` | 添加新的solution类型匹配逻辑 |

详细Schema说明见 `references/schema.md`。

## 依赖

- Python 3.8+
- openpyxl: `pip install openpyxl`
