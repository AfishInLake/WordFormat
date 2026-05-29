# config.yaml 字段规范

> 权威参考，编辑时只能使用本文档列出的字段。完整来源：`src/wordformat/config/datamodel.py`。

## GlobalFormat 基础字段（15 个）

所有段落配置节点都基于此字段集，通过 `<<: *global_format` 继承后覆盖。

| 字段 | 类型 | 可选值 | 默认值 |
|------|------|--------|--------|
| `alignment` | string | `左对齐` `居中对齐` `右对齐` `两端对齐` `分散对齐` | `左对齐` |
| `space_before` | string | `"0行"` `"0.5行"` `"12磅"` `"0.5cm"` | `"0.5行"` |
| `space_after` | string | 同上 | `"0.5行"` |
| `line_spacingrule` | string | `单倍行距` `1.5倍行距` `2倍行距` `最小值` `固定值` `多倍行距` | `单倍行距` |
| `line_spacing` | string | `"1倍"` `"1.5倍"` `"20磅"` `"0倍"` | `"1.5倍"` |
| `left_indent` | string | `"0字符"` `"2字符"` `"20磅"` `"0.75cm"` | `"0字符"` |
| `right_indent` | string | 同上 | `"0字符"` |
| `first_line_indent` | string | 正值为首行缩进，负值为悬挂缩进 | `"2字符"` |
| `chinese_font_name` | string | `宋体` `黑体` `楷体` `仿宋` `微软雅黑` | `宋体` |
| `english_font_name` | string | `Times New Roman` `Arial` `Calibri` | `Times New Roman` |
| `font_size` | string/number | `三号` `小三` `四号` `小四` `五号` `小五` 或 `12` `14` | `小四` |
| `font_color` | string | `黑色` `红色` 或 `#FF0000` | `黑色` |
| `bold` | bool | `true` / `false` | `false` |
| `italic` | bool | `true` / `false` | `false` |
| `underline` | bool | `true` / `false` | `false` |
| `builtin_style_name` | string | `正文` `Heading 1` `Heading 2` `Heading 3` `题注` | `正文` |

`line_spacingrule` 与 `line_spacing` 必须配合：

| line_spacingrule | line_spacing | 
|------------------|--------------|
| `固定值` | `"20磅"` |
| `1.5倍行距` | `"1.5倍"` |
| `单倍行距` | `"1倍"` |
| `最小值` | `"12磅"` |

## 各配置节点

### style_checks_warning（格式警告开关）

15 个 bool 字段，名称与 GlobalFormat 相同，默认均为 `true`（`font_name` 和 `font_color` 默认 `false`）。

### global_format（全局基准）

15 个 GlobalFormat 字段。使用 YAML 锚点 `&global_format` 定义，其他节点通过 `<<: *global_format` 继承。

### abstract（摘要）

```yaml
abstract:
  chinese:
    chinese_title:       # 继承 15 字段
    chinese_content:     # 继承 15 字段
  english:
    english_title:       # 继承 15 字段
    english_content:     # 继承 15 字段
  keywords:
    chinese:             # 继承 15 字段 + label + 3 专用字段
    english:             # 同上
```

keywords 专用字段：

| 字段 | 类型 | 默认值 |
|------|------|--------|
| `label` | GlobalFormat | 关键词标签（"关键词："）的字符格式 |
| `count_min` | int | `4` |
| `count_max` | int | `4` |
| `trailing_punct_forbidden` | bool | `true` |

### headings（标题）

```yaml
headings:
  level_1:              # 继承 15 字段，builtin_style_name 必须为 Heading 1
  level_2:              # 继承 15 字段，builtin_style_name 必须为 Heading 2
  level_3:              # 继承 15 字段，builtin_style_name 必须为 Heading 3
```

### body_text（正文）

继承 15 字段，通常不覆盖。

### figures（图注）

继承 15 字段 + `caption_prefix`（默认 `图`）。

```yaml
figures:
  <<: *global_format
  caption_prefix: '图'
  font_size: '五号'
  alignment: '居中对齐'
  first_line_indent: '0字符'
  builtin_style_name: '题注'
```

### tables（表注 + 表格内容）

继承 15 字段 + `caption_prefix`（默认 `表`）+ `content` 子配置（继承 15 字段，控制单元格内文字）。

```yaml
tables:
  <<: *global_format
  caption_prefix: '表'
  builtin_style_name: '题注'
  content:
    <<: *global_format
    font_size: '五号'
    line_spacingrule: '单倍行距'
```

### references（参考文献）

```yaml
references:
  title:                # 继承 15 字段
  content:              # 继承 15 字段
```

悬挂缩进示例：`first_line_indent: '-2.2字符'` + `left_indent: '0.26字符'`。

### acknowledgements（致谢）

```yaml
acknowledgements:
  title:                # 继承 15 字段
  content:              # 继承 15 字段
```

### numbering（自动编号）

**仅在 `wordf af` 模式下生效。**

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `enabled` | bool | `false` | 总开关 |
| `captions.enabled` | bool | `false` | 题注编号开关 |
| `captions.separator` | string | `.` | 章节号-编号分隔符 |
| `captions.label_number_space` | bool | `false` | 标签与编号间加空格 |
| `level_1/2/3.enabled` | bool | `false` | 该级编号开关 |
| `level_1/2/3.template` | string | — | `%1` `%1.%2` `第%1章` |
| `level_1/2/3.suffix` | string | `space` | `tab` `space` `nothing` |
| `level_1/2/3.numbering_indent` | string | — | 可选，如 `"0.75cm"` |
| `level_1/2/3.text_indent` | string | — | 可选，如 `"2.2字符"` |
| `references.enabled` | bool | `true` | 参考文献编号 |
| `references.template` | string | `[%1]` | 编号模板 |
| `references.suffix` | string | `space` | 同上 |

```yaml
numbering:
  enabled: true
  captions:
    enabled: true
    separator: '.'
  level_1:
    enabled: true
    template: '%1'
    suffix: space
  level_2:
    enabled: true
    template: '%1.%2'
    suffix: space
```

## 字段白名单

| 配置路径 | 合法字段 |
|----------|----------|
| `style_checks_warning` | bold, italic, underline, font_size, font_name, font_color, alignment, space_before, space_after, line_spacing, line_spacingrule, left_indent, right_indent, first_line_indent, builtin_style_name |
| `global_format` | 同上 + chinese_font_name, english_font_name, font_size, font_color, bold, italic, underline |
| `abstract.{zh/en}.{title/content}` | 同 global_format |
| `abstract.keywords.{zh/en}` | 同 global_format + label(同 global_format) + count_min, count_max, trailing_punct_forbidden |
| `headings.level_1/2/3` | 同 global_format |
| `body_text` | 同 global_format |
| `figures` | 同 global_format + caption_prefix |
| `tables` | 同 global_format + caption_prefix + content(同 global_format) |
| `references.title` | 同 global_format |
| `references.content` | 同 global_format |
| `acknowledgements.title/content` | 同 global_format |
| `numbering` | enabled, captions |
| `numbering.captions` | enabled, separator, label_number_space |
| `numbering.level_1/2/3/references` | enabled, template, suffix, numbering_indent, text_indent |
