# 配置文件编辑指南

> 本文档指导 Agent 如何正确编辑 `config.yaml` 文件。**编辑前必须先阅读 `config_spec.md` 了解完整字段规范。**

---

## 一、核心原则（必须遵守）

1. **只修改已有字段的值，绝不添加新字段** — 所有合法字段已在 `config_spec.md` 中列出
2. **绝不删除任何已有字段** — 即使某个字段值看起来"多余"也不能删
3. **绝不修改 YAML 锚点语法** — `&global_format` 和 `<<: *global_format` 是继承机制，不是普通文本
4. **编辑完成后必须运行验证脚本** — `python scripts/validate_config.py --config config.yaml`

---

## 二、编辑流程

### 步骤 1：提取格式需求

在编辑配置之前，**先从用户的格式要求文档中提取所有格式参数**，整理成清单。不要边看边改，容易遗漏细节。

需要提取的信息包括：
- 各部分（摘要、标题、正文、图表、参考文献、致谢）的字体、字号、加粗、对齐方式
- 行距类型和行距值
- 段前段后间距
- 缩进要求（首行缩进、左缩进）
- 关键词数量要求
- 图注/表注位置
- 参考文献特殊要求（编号格式、缩进、结尾标点）

### 步骤 2：复制模板

```bash
cp data/config.yaml config.yaml
```

### 步骤 3：逐节点修改

按照以下顺序修改，确保不遗漏：

1. `global_format` — 修改正文的默认格式（影响全局）
2. `abstract` — 修改摘要标题、正文、关键词格式
3. `headings` — 修改各级标题格式
4. `body_text` — 通常不需要改（继承 global_format）
5. `figures` — 修改图注格式
6. `tables` — 修改表注格式
7. `references` — 修改参考文献格式
8. `acknowledgements` — 修改致谢格式
9. `style_checks_warning` — 根据需要调整警告开关

### 步骤 4：验证

```bash
python scripts/validate_config.py --config config.yaml
```

---

## 三、字段取值速查表

### 对齐方式（alignment）

| 值 | 说明 |
|----|------|
| `左对齐` | 文本靠左 |
| `居中对齐` | 文本居中（常用于标题、图注） |
| `右对齐` | 文本靠右 |
| `两端对齐` | 文本两端对齐（正文常用） |
| `分散对齐` | 文本分散对齐 |

### 行距类型（line_spacingrule）+ 行距值（line_spacing）

| line_spacingrule | line_spacing | 说明 |
|------------------|--------------|------|
| `单倍行距` | `'1倍'` | 标准单倍行距 |
| `1.5倍行距` | `'1.5倍'` | 1.5 倍行距（论文常用） |
| `2倍行距` | `'2倍'` | 双倍行距 |
| `最小值` | `'12磅'` | 最小行距（值用磅） |
| `固定值` | `'20磅'` | 固定行距（值用磅） |
| `多倍行距` | `'1.25倍'` | 自定义倍数 |

> **重要**：`line_spacingrule` 和 `line_spacing` 必须配合使用。选择"固定值"或"最小值"时，`line_spacing` 的单位是磅；选择倍数行距时，`line_spacing` 的值是倍数。

### 中文字体（chinese_font_name）

| 值 | 说明 |
|----|------|
| `宋体` | 论文正文常用 |
| `黑体` | 标题常用 |
| `楷体` | 部分学校要求 |
| `仿宋` | 公文常用 |
| `微软雅黑` | 现代风格 |
| `汉仪小标宋` | 部分学校要求（需系统已安装） |

### 英文字体（english_font_name）

| 值 | 说明 |
|----|------|
| `Times New Roman` | 论文标准英文字体 |
| `Arial` | 无衬线字体 |
| `Calibri` | 现代无衬线字体 |
| `Courier New` | 等宽字体 |
| `Helvetica` | 经典无衬线字体 |

### 字号（font_size）

| 中文字号 | 磅值（约） | 常用场景 |
|----------|-----------|----------|
| `一号` | 26pt | 封面 |
| `小一` | 24pt | 封面 |
| `二号` | 22pt | 章标题 |
| `小二` | 18pt | 章标题、摘要标题 |
| `三号` | 16pt | 节标题、参考文献标题 |
| `小三` | 15pt | 二级标题 |
| `四号` | 14pt | 英文摘要标题 |
| `小四` | 12pt | 正文（最常用） |
| `五号` | 10.5pt | 图注、表注、参考文献内容 |
| `小五` | 9pt | 脚注、页眉 |
| `六号` | 7.5pt | |
| `七号` | 5.5pt | |

> 也可以直接使用数值，如 `font_size: 12` 等同于 `font_size: '小四'`。

### 间距和缩进

| 字段 | 单位 | 常用值 |
|------|------|--------|
| `space_before` | 行/磅/厘米/毫米/英寸 | `"0行"` `"0.5行"` `"12磅"` `"0.5cm"` |
| `space_after` | 同上 | `"0行"` `"0.5行"` `"12磅"` |
| `left_indent` | 字符/磅/厘米/毫米/英寸 | `"0字符"` `"2字符"` `"20磅"` |
| `right_indent` | 同上 | `"0字符"` `"2字符"` |
| `first_line_indent` | 同上 | `"0字符"` `"2字符"` `"20磅"` |

### 关键词专用字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `keywords_bold` | bool | 关键词文字是否加粗 |
| `count_min` | int | 最少关键词数量（必须 > 0） |
| `count_max` | int | 最多关键词数量（必须 > 0） |
| `trailing_punct_forbidden` | bool | 是否禁止最后一个关键词后有标点 |

### 图表专用字段

| 字段 | 可选值 | 说明 |
|------|--------|------|
| `caption_position` | `above` / `below` | 题注在上方/下方 |
| `caption_prefix` | 任意字符串 | 题注前缀，如 `图` `表` `Figure` `Table` |

### 参考文献专用字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `section_title` | string / null | 参考文献标题文字，如 `参考文献`。`null` = 不限制 |
| `numbering_format` | string / null | 编号格式描述，如 `'[1], [2], ...'`。`null` = 不限制 |
| `entry_indent` | float / null | 条目首行缩进量，`0.0` = 顶格。`null` = 不限制 |
| `entry_ending_punct` | string / null | 条目结束标点，如 `'.'` `'。'`。`null` = 不限制 |

---

## 四、常见修改示例

### 示例 1：修改正文字体为小四宋体 + Times New Roman，两端对齐

```yaml
global_format: &global_format
  alignment: '两端对齐'
  chinese_font_name: '宋体'
  english_font_name: 'Times New Roman'
  font_size: '小四'
```

### 示例 2：修改一级标题为三号黑体居中加粗

```yaml
headings:
  level_1:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '三号'
    bold: true
    builtin_style_name: 'Heading 1'
```

### 示例 3：修改行距为固定值 20 磅

```yaml
global_format: &global_format
  line_spacingrule: "固定值"
  line_spacing: '20磅'
```

### 示例 4：修改关键词数量为 3~6 个

```yaml
abstract:
  keywords:
    chinese:
      <<: *global_format
      count_min: 3
      count_max: 6
    english:
      <<: *global_format
      count_min: 3
      count_max: 6
```

### 示例 5：修改图注为五号字居中，位于图下方

```yaml
figures:
  <<: *global_format
  caption_position: 'below'
  caption_prefix: '图'
  font_size: '五号'
  alignment: '居中对齐'
  first_line_indent: '0字符'
  builtin_style_name: '题注'
```

### 示例 6：参考文献条目悬挂缩进

```yaml
references:
  content:
    <<: *global_format
    entry_indent: 0.0
    entry_ending_punct: '.'
    numbering_format: '[1], [2], ...'
```

---

## 五、禁止事项

- ❌ **添加模板中不存在的字段**（如 `caption_numbering`、`page_margin` 等都不存在）
- ❌ **删除 `<<: *global_format` 继承引用**
- ❌ **修改 `&global_format` 锚点定义**（`global_format:` 后面的 `&global_format`）
- ❌ **修改顶层结构名称**（abstract、headings、body_text 等名称固定不可改）
- ❌ **在字符串值中使用未转义的特殊字符**（YAML 中的 `:` `{` `}` `[` `]` 等需用引号包裹）
- ❌ **修改 abstract 下的子节点名称**（chinese_title、chinese_content、english_title、english_content、keywords 固定不可改）
- ❌ **修改 headings 下的子节点名称**（level_1、level_2、level_3 固定不可改）
- ❌ **修改 references 下的子节点名称**（title、content 固定不可改）
- ❌ **修改 acknowledgements 下的子节点名称**（title、content 固定不可改）
