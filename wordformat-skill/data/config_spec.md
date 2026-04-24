# WordFormat 配置文件完整格式规范

> 本文档是 `config.yaml` 配置文件的**权威参考**。所有可用字段、取值范围、继承关系均来源于源码 `src/wordformat/config/datamodel.py`。
> Agent 编辑配置时，**必须只使用本文档列出的字段**，不得添加任何未列出的字段。

---

## 一、文件整体结构

配置文件是 YAML 格式，包含以下 **9 个顶层节点**，顺序不限但缺一不可：

```yaml
style_checks_warning:   # ① 格式警告开关
global_format:          # ② 全局基础格式（锚点定义）
abstract:               # ③ 摘要（中英文标题、正文、关键词）
headings:               # ④ 各级标题（level_1/2/3）
body_text:              # ⑤ 正文段落
figures:                # ⑥ 插图及图注
tables:                 # ⑦ 表格及表注
references:             # ⑧ 参考文献（标题 + 条目）
acknowledgements:       # ⑨ 致谢（标题 + 正文）
```

---

## 二、YAML 锚点继承机制

`global_format` 使用 YAML 锚点 `&global_format` 定义，其他节点通过 `<<: *global_format` 继承其全部字段，然后只需覆盖需要不同的字段即可。

```yaml
global_format: &global_format    # 定义锚点
  alignment: '两端对齐'
  font_size: '小四'
  # ... 其他字段

headings:
  level_1:
    <<: *global_format           # 继承 global_format 的所有字段
    alignment: '居中对齐'        # 仅覆盖需要对齐方式的字段
    font_size: '小二'
```

**重要**：`<<: *global_format` 不是普通字段，是 YAML 合并键语法，**绝对不能删除或修改**。

---

## 三、各节点字段详解

### ① style_checks_warning（格式警告开关）

控制格式校验时，哪些属性不满足规范时触发警告提示。所有字段均为 `bool` 类型。

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `bold` | true | 加粗不规范时警告 |
| `italic` | true | 斜体不规范时警告 |
| `underline` | true | 下划线不规范时警告 |
| `font_size` | true | 字号不规范时警告 |
| `font_name` | false | 字体名称不规范时警告 |
| `font_color` | false | 字体颜色不规范时警告 |
| `alignment` | true | 对齐方式不规范时警告 |
| `space_before` | true | 段前间距不规范时警告 |
| `space_after` | true | 段后间距不规范时警告 |
| `line_spacing` | true | 行距数值不规范时警告 |
| `line_spacingrule` | true | 行距类型不规范时警告 |
| `left_indent` | true | 文本之前（左缩进）不规范时警告 |
| `right_indent` | true | 文本之后（右缩进）不规范时警告 |
| `first_line_indent` | true | 首行缩进不规范时警告 |
| `builtin_style_name` | true | Word 内置样式名称不规范时警告 |

---

### ② global_format（全局基础格式）

所有段落格式默认继承的基础样式。以下 15 个字段构成了"段落格式字段集"，后续多个节点都会继承这套字段。

#### 段落格式字段（8 个）

| 字段 | 类型 | 可选值 | 默认值 | 说明 |
|------|------|--------|--------|------|
| `alignment` | string | `左对齐` `居中对齐` `右对齐` `两端对齐` `分散对齐` | `左对齐` | 段落对齐方式 |
| `space_before` | string | 带单位的值，如 `"0行"` `"0.5行"` `"12磅"` `"0.5cm"` | `"0.5行"` | 段前间距 |
| `space_after` | string | 同上 | `"0.5行"` | 段后间距 |
| `line_spacingrule` | string | `单倍行距` `1.5倍行距` `2倍行距` `最小值` `固定值` `多倍行距` | `单倍行距` | 行距类型 |
| `line_spacing` | string | 倍数值，如 `"1倍"` `"1.5倍"` `"2倍"` `"0倍"` | `"1.5倍"` | 行距参数 |
| `left_indent` | string | 带单位的值，如 `"0字符"` `"2字符"` `"20磅"` | `"0字符"` | 文本之前（左缩进） |
| `right_indent` | string | 同上 | `"0字符"` | 文本之后（右缩进） |
| `first_line_indent` | string | 同上 | `"2字符"` | 段落首行缩进 |

#### 字符格式字段（7 个）

| 字段 | 类型 | 可选值 | 默认值 | 说明 |
|------|------|--------|--------|------|
| `chinese_font_name` | string | `宋体` `黑体` `楷体` `仿宋` `微软雅黑` `汉仪小标宋`，或其他自定义字体名 | `宋体` | 中文字体 |
| `english_font_name` | string | `Times New Roman` `Arial` `Calibri` `Courier New` `Helvetica`，或其他自定义字体名 | `Times New Roman` | 英文字体 |
| `font_size` | string/number | 中文字号：`一号` `小一` `二号` `小二` `三号` `小三` `四号` `小四` `五号` `小五` `六号` `七号`；或数值如 `12` `14` `10.5` | `小四` | 字号 |
| `font_color` | string | 颜色名如 `黑色` `红色`，或十六进制如 `#FF0000` | `黑色` | 字体颜色 |
| `bold` | bool | `true` / `false` | `false` | 是否加粗 |
| `italic` | bool | `true` / `false` | `false` | 是否斜体 |
| `underline` | bool | `true` / `false` | `false` | 是否有下划线 |

#### 样式映射字段（1 个）

| 字段 | 类型 | 说明 | 常用值 |
|------|------|------|--------|
| `builtin_style_name` | string | 对应 Word 内置样式名称，用于样式匹配 | `正文`、`Heading 1`、`Heading 2`、`Heading 3`、`题注` |

> **注意**：以上 15 个字段 = "GlobalFormat 字段集"，后续 abstract、headings、body_text、figures、tables、references、acknowledgements 的子节点都继承此字段集，可覆盖任意字段。

---

### ③ abstract（摘要配置）

结构如下：

```yaml
abstract:
  chinese:                          # 中文摘要
    chinese_title:                  # 中文摘要标题
      <<: *global_format            # 继承全部 15 个字段
      # 覆盖需要的字段...
    chinese_content:                # 中文摘要正文
      <<: *global_format
  english:                          # 英文摘要
    english_title:                  # 英文摘要标题
      <<: *global_format
    english_content:                # 英文摘要正文
      <<: *global_format
  keywords:                         # 关键词（字典结构）
    chinese:                        # 中文关键词
      <<: *global_format            # 继承全部 15 个字段 + 以下 4 个专用字段
      keywords_bold: true
      count_min: 4
      count_max: 4
      trailing_punct_forbidden: true
    english:                        # 英文关键词
      <<: *global_format
      keywords_bold: true
      count_min: 4
      count_max: 4
      trailing_punct_forbidden: true
```

#### 关键词专用字段（仅 keywords.chinese / keywords.en 可用）

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `keywords_bold` | bool | `true` | 关键词文字是否加粗 |
| `count_min` | int | `4` | 最少关键词数量（必须 > 0） |
| `count_max` | int | `4` | 最多关键词数量（必须 > 0） |
| `trailing_punct_forbidden` | bool | `true` | 是否禁止最后一个关键词后有标点符号 |

---

### ④ headings（各级标题配置）

```yaml
headings:
  level_1:                         # 一级标题（章标题）
    <<: *global_format              # 继承全部 15 个字段
    builtin_style_name: 'Heading 1' # 必须设置
  level_2:                         # 二级标题（节标题）
    <<: *global_format
    builtin_style_name: 'Heading 2'
  level_3:                         # 三级标题（小节标题）
    <<: *global_format
    builtin_style_name: 'Heading 3'
```

每个级别可覆盖 global_format 中的任意字段。**`builtin_style_name` 必须与级别对应**：
- `level_1` → `'Heading 1'`
- `level_2` → `'Heading 2'`
- `level_3` → `'Heading 3'`

---

### ⑤ body_text（正文段落配置）

```yaml
body_text:
  <<: *global_format                # 继承全部 15 个字段
```

通常不需要覆盖任何字段，直接继承全局格式即可。如需特殊格式可覆盖。

---

### ⑥ figures（插图及图注配置）

继承 global_format 的 15 个字段 + 以下 2 个专用字段：

| 字段 | 类型 | 默认值 | 可选值 | 说明 |
|------|------|--------|--------|------|
| `caption_position` | string | `below` | `above`（图注在图上方）`below`（图注在图下方） | 图注位置 |
| `caption_prefix` | string | `图` | 任意字符串，如 `图` `Figure` `Fig.` | 图注编号前缀 |

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

---

### ⑦ tables（表格及表注配置）

继承 global_format 的 15 个字段 + 以下 2 个专用字段：

| 字段 | 类型 | 默认值 | 可选值 | 说明 |
|------|------|--------|--------|------|
| `caption_position` | string | `above` | `above`（表注在表上方）`below`（表注在表下方） | 表注位置 |
| `caption_prefix` | string | `表` | 任意字符串，如 `表` `Table` `Tab.` | 表注编号前缀 |

```yaml
tables:
  <<: *global_format
  caption_position: 'above'
  caption_prefix: '表'
  font_size: '五号'
  alignment: '居中对齐'
  first_line_indent: '0字符'
  builtin_style_name: '题注'
```

---

### ⑧ references（参考文献配置）

结构如下：

```yaml
references:
  title:                           # 参考文献标题
    <<: *global_format              # 继承全部 15 个字段 + section_title
    section_title: '参考文献'
  content:                         # 参考文献条目
    <<: *global_format              # 继承全部 15 个字段 + 以下 3 个专用字段
    entry_indent: 0.0
    entry_ending_punct: '.'
    numbering_format: '[1], [2], ...'
```

#### references.title 专用字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `section_title` | string / null | `参考文献` | 参考文献章节的标题文字，设为 `null` 可不限制 |

#### references.content 专用字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `numbering_format` | string / null | `null` | 参考文献编号格式描述，如 `'[1], [2], ...'` 或 `'1.'`。设为 `null` 不限制 |
| `entry_indent` | float / null | `0.0` | 条目首行缩进量。`0.0` = 顶格，`null` 不限制 |
| `entry_ending_punct` | string / null | `null` | 条目结束标点，如 `'.'` 或 `'。'`。设为 `null` 表示不限制 |

---

### ⑨ acknowledgements（致谢配置）

结构如下：

```yaml
acknowledgements:
  title:                           # 致谢标题
    <<: *global_format              # 继承全部 15 个字段，无额外字段
  content:                         # 致谢正文
    <<: *global_format              # 继承全部 15 个字段，无额外字段
```

title 和 content 均只继承 global_format 的 15 个字段，无专用字段。

---

## 四、字段继承关系总览

```
GlobalFormat（15 个字段）
├── style_checks_warning     → 独立 15 个 bool 字段（名称相同但类型不同）
├── global_format            → 基准定义（锚点 &global_format）
├── abstract
│   ├── chinese.chinese_title      → 继承 15 字段
│   ├── chinese.chinese_content    → 继承 15 字段
│   ├── english.english_title      → 继承 15 字段
│   ├── english.english_content    → 继承 15 字段
│   └── keywords.chinese/english   → 继承 15 字段 + 4 个关键词专用字段
├── headings
│   ├── level_1              → 继承 15 字段
│   ├── level_2              → 继承 15 字段
│   └── level_3              → 继承 15 字段
├── body_text                → 继承 15 字段
├── figures                  → 继承 15 字段 + 2 个图表专用字段
├── tables                   → 继承 15 字段 + 2 个图表专用字段
├── references
│   ├── title                → 继承 15 字段 + 1 个专用字段（section_title）
│   └── content              → 继承 15 字段 + 3 个专用字段
└── acknowledgements
    ├── title                → 继承 15 字段
    └── content              → 继承 15 字段
```

---

## 五、完整字段白名单（供验证脚本使用）

以下是每个配置路径下**合法字段**的完整列表，任何不在此列表中的字段都是非法的：

| 配置路径 | 合法字段 |
|----------|----------|
| `style_checks_warning` | bold, italic, underline, font_size, font_name, font_color, alignment, space_before, space_after, line_spacing, line_spacingrule, left_indent, right_indent, first_line_indent, builtin_style_name |
| `global_format` | alignment, space_before, space_after, line_spacingrule, line_spacing, left_indent, right_indent, first_line_indent, builtin_style_name, chinese_font_name, english_font_name, font_size, font_color, bold, italic, underline |
| `abstract.chinese.chinese_title` | 同 global_format |
| `abstract.chinese.chinese_content` | 同 global_format |
| `abstract.english.english_title` | 同 global_format |
| `abstract.english.english_content` | 同 global_format |
| `abstract.keywords.chinese` | 同 global_format + keywords_bold, count_min, count_max, trailing_punct_forbidden |
| `abstract.keywords.english` | 同 global_format + keywords_bold, count_min, count_max, trailing_punct_forbidden |
| `headings.level_1/2/3` | 同 global_format |
| `body_text` | 同 global_format |
| `figures` | 同 global_format + caption_position, caption_prefix |
| `tables` | 同 global_format + caption_position, caption_prefix |
| `references.title` | 同 global_format + section_title |
| `references.content` | 同 global_format + entry_indent, entry_ending_punct, numbering_format |
| `acknowledgements.title` | 同 global_format |
| `acknowledgements.content` | 同 global_format |
