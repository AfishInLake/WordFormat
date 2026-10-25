# config.yaml 字段规范

## 基础字段

所有配置节点基于 `global_format` 定义，通过 `<<: *global_format` 继承。样式字段统一平铺，`paragraph:` / `font:` 嵌套写法已移除。

```yaml
global_format: &global_format
  # 段落格式
  alignment: '两端对齐'
  space_before: "0行"
  space_after: "0行"
  line_spacingrule: "1.5倍行距"
  line_spacing: '1.5倍'
  left_indent: "0字符"
  right_indent: "0字符"
  first_line_indent: '2字符'
  builtin_style_name: '正文'
  # 字符格式
  chinese_font_name: '宋体'
  english_font_name: 'Times New Roman'
  font_size: '小四'
  font_color: '黑色'
  bold: false
  italic: false
  underline: false
```

| 类别 | 字段 | 可选值 |
|------|------|--------|
| 段落 | `alignment` | 左对齐 / 居中对齐 / 右对齐 / 两端对齐 / 分散对齐 |
| 段落 | `space_before/after` | "0行" "0.5行" "12磅" "0.5cm" |
| 段落 | `line_spacingrule` | 单倍行距 / 1.5倍行距 / 2倍行距 / 最小值 / 固定值 / 多倍行距 |
| 段落 | `line_spacing` | "1.5倍" "20磅" |
| 段落 | `left/right_indent` | "0字符" "2字符" "20磅" |
| 段落 | `first_line_indent` | 正值=首行缩进，负值=悬挂缩进 |
| 段落 | `builtin_style_name` | 正文 / Heading 1/2/3 / 题注 |
| 字符 | `chinese_font_name` | 宋体 / 黑体 / 楷体 / 仿宋 |
| 字符 | `english_font_name` | Times New Roman / Arial |
| 字符 | `font_size` | 三号 / 小三 / 四号 / 小四 / 五号 / 12 / 14 |
| 字符 | `font_color` | 黑色 / 红色 / #FF0000 |
| 字符 | `bold/italic/underline` | true / false |

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

> `tables.object`（表格对象）是只读导出，没有格式配置；`abstract.chinese` / `abstract.english` 是"摘要和正文同段"的混合节点，同时读取该节下的 `title` 与 `body`。

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

基础字段 + `rules`（0 = 不限制该项）：

```yaml
rules:
  punctuation:
    enabled: true
  paragraph_length:
    enabled: true
    min_chars: 0        # 字数下限
    max_chars: 0        # 字数上限
    min_sentences: 0    # 句数下限
    max_sentences: 0    # 句数上限
```

### 图片段落 (figures.image)

基础字段 + `rules`：

```yaml
rules:
  figure_layout:
    enabled: true
    wrap: null          # 要求的环绕方式（null = 不查）；如 inline/四周型
    max_width: '版心'    # 允许最大宽度；'版心' = 页宽 − 左右边距，或写 '14cm' 等
```

### style_checks_warning

控制各字段差异是否在批注中显示:

| 字段 | 类型 | 默认 |
|------|------|------|
| `alignment` | bool | true |
| `space_before/after` | bool | true |
| `line_spacing/line_spacing_rule` | bool | true |
| `left/right/first_line_indent` | bool | true |
| `builtin_style_name` | bool | false |
| `font_size` | bool | true |
| `font_name_cn/font_name_en` | bool | false |
| `font_color` | bool | false |
| `bold` | bool | true |
| `italic/underline` | bool | true |

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

## header_footer（页眉页脚 / 页面设置）

`enabled` 与 `header`/`footer` 有内置缺省；`page_setup` 是 opt-in —— 只写了才校验（各校边距差异大，避免误报）。

```yaml
header_footer:
  enabled: true
  update_fields_on_open: false   # 打开文档时自动更新域（w:updateFields）
  header:
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '小五'
    alignment: '居中对齐'
    require_content: false       # 要求页眉非空
    page_number: false           # 要求存在 PAGE 页码域（apply 会插入）
  footer: { ...同上... }
  page_setup:
    top_margin: '2.54cm'
    bottom_margin: '2.54cm'
    left_margin: '3.18cm'
    right_margin: '3.18cm'
    header_distance: '1.5cm'
    footer_distance: '1.75cm'
    page_size: 'A4'              # A3 / A4 / A5 / B5 / Letter
    orientation: '纵向'          # 纵向 / 横向
    page_number_start: 1         # w:pgNumType/@w:start
    page_number_format: '小写罗马数字'
    different_first_page: false  # 首页不同（w:titlePg）
    odd_even_pages: false        # 奇偶页不同（w:evenAndOddHeaders）
```

页眉页脚只检查四项：字号、中文字体、英文字体、对齐。

## tables（表格样式档案）

镜像 OOXML tblPr/trPr/tcPr，缺省 = Table Grid。条目可用 `table_style` 指定档案名。

```yaml
tables:
  default: grid
  styles:
    grid:
      style: "Table Grid"        # Word 样式名；null = 不套样式，全由 XML 定义
      layout: autofit            # autofit|fixed
      width: {type: auto}        # auto|pct|dxa|cm
      alignment: center          # left|center|right
      indent_cm: 0
      borders:                   # 边名：top/left/bottom/right/inside_h/inside_v
        top: {style: single, size_pt: 1.5, color: "000000"}
        left: {style: none}
        inside_h: {style: single, size_pt: 0.75}
      cell_margins_cm: {top: 0, bottom: 0, left: 0.19, right: 0.19}
      col_widths: {mode: equal}      # equal|auto|custom（custom: [4, 6, 6]）
      header_row: {repeat: true, cant_split: true}
      row_bands: [{rows: [0], fill: "EEEEEE", bold: true}]   # rows: all | 整数 | 整数列表
      cell_run: {chinese_font_name: 宋体, english_font_name: Times New Roman, font_size: 小四}
```

`borders` 的边可写简写字符串（`top: double` 等价 `top: {style: double}`）；`style` 取 OOXML val（none/single/double/dashed/...）。

## inline（行内富文本）

`md` / `jd` 按 JSON 的 `inline_marks` 生成多 run 时使用：

```yaml
inline:
  code_font: {chinese_font_name: 宋体, english_font_name: Consolas, font_size: 五号}
  link: {style: Hyperlink, color: theme, underline: true}   # color: theme|<RRGGBB>
```

## structure（结构类别策略）

按类别名单驱动，不绑定业务语义：

```yaml
structure:
  void_categories: [top, heading_mulu, heading_fulu, other, document_title, footer, toc_entry]
  skip_children_categories: [heading_mulu, heading_fulu, other]
```

名单是权威的：显式写空列表（`void_categories: []`）表示不排除任何类别，不会回退到默认值。直接挂在 root 下的 `body_text` 始终被隐式跳过。
