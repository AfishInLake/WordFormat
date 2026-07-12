# config.yaml 字段规范

## 基础字段

所有配置节点基于 `global_format` 定义，通过 `<<: *global_format` 继承。分为三个组：

```yaml
global_format: &global_format
  paragraph:             # 段落格式
    alignment: '两端对齐'
    space_before: "0行"
    space_after: "0行"
    line_spacingrule: "1.5倍行距"
    line_spacing: '1.5倍'
    left_indent: "0字符"
    right_indent: "0字符"
    first_line_indent: '2字符'
    builtin_style_name: '正文'
  font:                  # 字符格式
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '小四'
    font_color: '黑色'
    bold: false
    italic: false
    underline: false
```

| 组 | 字段 | 可选值 |
|----|------|--------|
| paragraph | `alignment` | 左对齐 / 居中对齐 / 右对齐 / 两端对齐 / 分散对齐 |
| paragraph | `space_before/after` | "0行" "0.5行" "12磅" "0.5cm" |
| paragraph | `line_spacingrule` | 单倍行距 / 1.5倍行距 / 2倍行距 / 最小值 / 固定值 / 多倍行距 |
| paragraph | `line_spacing` | "1.5倍" "20磅" |
| paragraph | `left/right_indent` | "0字符" "2字符" "20磅" |
| paragraph | `first_line_indent` | 正值=首行缩进，负值=悬挂缩进 |
| paragraph | `builtin_style_name` | 正文 / Heading 1/2/3 / 题注 |
| font | `chinese_font_name` | 宋体 / 黑体 / 楷体 / 仿宋 |
| font | `english_font_name` | Times New Roman / Arial |
| font | `font_size` | 三号 / 小三 / 四号 / 小四 / 五号 / 12 / 14 |
| font | `font_color` | 黑色 / 红色 / #FF0000 |
| font | `bold/italic/underline` | true / false |

> 覆盖时可直接用平铺字段（如 `alignment: '居中对齐'`），框架自动区分 paragraph/font。

## 配置路径

| 路径 | 说明 |
|------|------|
| `abstract.chinese.title` | 中文摘要标题 |
| `abstract.chinese.body` | 中文摘要正文 |
| `abstract.chinese.keywords` | 中文关键词 |
| `abstract.english.title` | 英文摘要标题 |
| `abstract.english.body` | 英文摘要正文 |
| `abstract.english.keywords` | 英文关键词 |
| `headings.level_1/2/3` | 一/二/三级标题 |
| `body.text` | 正文段落 |
| `figures.caption` | 图题注 |
| `figures.image` | 图片段落 |
| `tables.caption` | 表题注 |
| `tables.object` | 表格对象 |
| `references.title` | 参考文献标题 |
| `references.entry` | 参考文献条目 |
| `acknowledgements.title` | 致谢标题 |
| `acknowledgements.body` | 致谢正文 |

## 特殊节点字段

### 关键词 (abstract.{zh/en}.keywords)

基础字段 + `label`(字符格式) + `rules`:

```yaml
label:
  chinese_font_name: '黑体'
  font_size: '四号'
  bold: false
rules:
  keyword_count:
    enabled: true
    count_min: 3
    count_max: 5
  trailing_punctuation:    # 仅中文
    enabled: true
    forbidden_chars: "；，。、"
```

### 题注 (figures.caption / tables.caption)

基础字段 + `caption_prefix` + `rules`:

```yaml
caption_prefix: '图'       # 或 '表'
rules:
  caption_numbering:
    enabled: true
    separator: '.'
    label_number_space: false
```

### 正文 (body.text)

基础字段 + `rules`:

```yaml
rules:
  punctuation:
    enabled: true
```

### style_checks_warning

控制各字段差异是否在批注中显示:

| 字段 | 类型 | 默认 |
|------|------|------|
| `alignment` | bool | true |
| `space_before/after` | bool | true |
| `line_spacing/line_spacing_rule` | bool | true |
| `left/right/first_line_indent` | bool | true |
| `builtin_style_name` | bool | true |
| `font_size` | bool | true |
| `font_name_cn/font_name_en` | bool | false |
| `font_color` | bool | false |
| `bold` | bool | true |
| `italic/underline` | bool | false |

## numbering（自动编号）

```yaml
numbering:
  enabled: true
  level_1/2/3:
    enabled: true
    template: '%1'         # %1.%2, 第%1章 等
    suffix: space           # tab / space / nothing
    numbering_indent:      # 可选
    text_indent:           # 可选
  references:
    enabled: true
    template: '[%1]'
    suffix: space
```
