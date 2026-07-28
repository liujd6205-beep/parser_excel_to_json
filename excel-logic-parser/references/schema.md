# JSON Schema 与字段说明

## 顶层结构

```json
{
  "summary": { ... },     // 统计信息
  "fields": [ ... ]       // 按字段名分组的逻辑列表
}
```

## summary 统计信息

| 字段 | 类型 | 说明 |
|------|------|------|
| total_fields | int | 字段总数 |
| total_logics | int | 逻辑条目总数 |
| action_stats | dict | action类型统计（key=类型, value=数量） |
| solution_stats | dict | solution类型统计（key=类型, value=数量） |
| red_scene_logics | int | 字段逻辑场景(B列)含红色字体的条目数 |
| red_scene_with_text | int | 字段逻辑场景(B列)含红色备注文本的条目数 |

## fields 数组结构

每个元素代表一个字段：

```json
{
  "field_name": "新生儿听力检查阳性随访治疗",
  "field_note": null,
  "logics": [ ... ]
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| field_name | string | 字段名称 |
| field_note | string\|null | 字段备注（A列换行后的第二行文字） |
| logics | array | 该字段下的所有逻辑条目 |

## logics 数组结构

每个元素代表一条逻辑规则：

```json
{
  "row": 2,
  "field_name": "新生儿听力检查阳性随访治疗",
  "field_note": null,
  "logic_scene": "新生儿听力检查阳性随访治疗为空，提示",
  "logic_scene_is_red": false,
  "logic_scene_red_text": null,
  "prompt_text": "【新生儿听力检查阳性随访治疗】：（仅提示）新生儿听力检查阳性随访治疗为空",
  "solution": {
    "type": "prompt_only",
    "ideal_data": null,
    "description": "仅提示不修改"
  },
  "action": {
    "type": "delete_logic",
    "detail": "删除逻辑20260607",
    "is_red": false
  }
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| row | int | Excel中的行号 |
| field_name | string | 所属字段名称 |
| field_note | string\|null | 字段备注 |
| logic_scene | string\|null | 字段逻辑场景描述（需转为Python逻辑的中文描述） |
| logic_scene_is_red | bool | 字段逻辑场景(B列)是否含红色字体（特殊关注） |
| logic_scene_red_text | array\|null | B列中红色字体的文本内容列表（红色备注行） |
| prompt_text | string\|null | 质检结果提示文案 |
| solution | object | 质控解决方案（见下） |
| action | object | 调整日志（见下） |

## action 调整日志结构

| 字段 | 类型 | 说明 |
|------|------|------|
| type | string | action类型（见下表） |
| detail | string | 调整日志原文 |
| is_red | bool | 是否需特别关注（基于B列红色字体判断，非E列） |

### action.type 枚举值

| type | 说明 | 脚本处理方式 |
|------|------|-------------|
| delete_logic | 删除逻辑 | 在该字段的函数方法上删除此逻辑 |
| add_logic | 新增逻辑 | 增加此逻辑 |
| delete_visit_face_to_face | 删除访视方式为面对面 | 删除面对面访视相关逻辑 |
| adjust_logic | 逻辑调整 | 调整已有逻辑（如范围值变更） |
| none | 无调整日志 | 保持原样 |
| unknown | 未识别的类型 | 需人工确认 |

## solution 质控解决方案结构

| 字段 | 类型 | 说明 |
|------|------|------|
| type | string | solution类型（见下表） |
| ideal_data | any | 理想数据（None表示只提示不修改） |
| description | string | 人类可读的描述 |

### solution.type 枚举值

| type | 说明 | ideal_data格式 | 脚本处理方式 |
|------|------|---------------|-------------|
| prompt_only | 只提示不修改 | null | 仅记录提示，不修改数据 |
| modify_select | 选择固定值 | string | 将字段值设为ideal_data |
| modify_add | 添加内容 | string | 向字段添加ideal_data内容 |
| modify_add_check | 添加勾选 | string | 勾选ideal_data对应的选项 |
| modify_contain | 必须包含 | string | 确保字段值包含ideal_data |
| modify_random | 范围随机取值 | {"min": float, "max": float} | 在[min, max]范围内随机取值 |
| modify_fill | 填写固定文本 | string | 将字段值设为ideal_data |
| modify_clear | 清空 | "清空" | 清空该字段值（注意：ideal_data是"清空"两个字） |

## 修改指南

### 添加新的 action 类型

编辑 `scripts/parse_excel_logic.py` 中的 `ACTION_PATTERNS` 列表：

```python
ACTION_PATTERNS = [
    (r"^删除逻辑",   "delete_logic"),
    (r"^新增逻辑",   "add_logic"),
    # 添加新规则：
    (r"^你的新模式", "new_type_name"),
]
```

注意：`is_red` 不再由 ACTION_PATTERNS 配置，而是基于 B列(字段逻辑场景)的红色字体自动检测。

### 添加新的 solution 类型

编辑 `scripts/parse_excel_logic.py` 中的 `parse_solution()` 函数，在现有规则之后添加新的匹配逻辑。

### 调整红色字体检测

红色字体检测针对 **B列（字段逻辑场景）**，支持 Rich Text（同一单元格内混合黑白+红色文字）。

编辑 `RED_COLOR_PATTERNS` 列表，添加或修改需要检测的颜色模式。
同时可通过 `get_red_text()` 函数提取红色字体的文本内容。

### 调整列映射

如果Excel的列顺序不同，修改 `COLUMN_MAP`：

```python
COLUMN_MAP = {
    "field_name":   1,  # 改为实际的列号
    "logic_scene":  2,
    "prompt_text":  3,
    "solution":     4,
    "adjust_log":   5,
}
```
