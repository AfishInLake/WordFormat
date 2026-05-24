# WordFormat 配置文件完整格式规范

> 本文档是 `config.yaml` 配置文件的**权威参考**。所有可用字段、取值范围、继承关系均来源于源码 `src/wordformat/config/datamodel.py`。
> Agent 编辑配置时，**必须只使用本文档列出的字段**，不得添加任何未列出的字段。

---

## 一、文件整体结构

配置文件是 YAML 格式，包含以下 **10 个顶层节点**，顺序不限但缺一不可：

```yaml
style_checks_warning:   # ① 格式警告开关
global_format:          # ② 全局基础格式（锚点定义）
abstract:               # ③ 摘要（中英文标题、正文、关键词）
headings:               # ④ 各级标题（level_1/2/3）
body_text:              # ⑤ 正文段落
figures:                # ⑥ 插图及图注
tables:                 # ⑦ 表格题注及表格内容格式
references:             # ⑧ 参考文献（标题 + 条目）
acknowledgements:       # ⑨ 致谢（标题 + 正文）
numbering:              # ⑩ 标题自动编号 + 参考文献编号
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
| `left_indent` | string | 带单位的值，如 `"0字符"` `"2字符"` `"0.26字符"` `"20磅"` | `"0字符"` | 文本之前（左缩进） |
| `right_indent` | string | 同上 | `"0字符"` | 文本之后（右缩进） |
| `first_line_indent` | string | 同上。**正值**=首行缩进，**负值**=悬挂缩进（如 `"-2.2字符"`） | `"2字符"` | 段落首行缩进/悬挂缩进 |

> **悬挂缩进说明**：`first_line_indent` 设为负值（如 `"-2.2字符"`）时，第一行保持不变、后续行向右缩进指定位数。使用时通常配合 `left_indent` 设置整体左移量。常见于参考文献条目格式。

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
      <<: *global_format            # 继承全部 15 个字段（用于关键词内容部分）
      label:                        # 关键词标签（"关键词："）的字符格式
        <<: *global_format          # 继承全部 15 个字段
        chinese_font_name: '黑体'
        font_size: '四号'
        bold: true
      count_min: 3
      count_max: 5
      trailing_punct_forbidden: true
    english:                        # 英文关键词
      <<: *global_format
      label:
        <<: *global_format
        font_size: '四号'
        bold: true
      count_min: 3
      count_max: 5
      trailing_punct_forbidden: true
```

#### 关键词标签子配置（label）

`keywords.chinese.label` 和 `keywords.english.label` 是 `KeywordLabelConfig` 类型，继承 `GlobalFormatConfig` 的 15 个字段。用于控制"关键词："或"Keywords:"标签部分（不包括后面内容）的字体、字号、加粗等字符格式。段落级字段（如 `alignment`、`line_spacingrule`）会被忽略。

| label 中常用字段 | 类型 | 说明 |
|------|------|------|
| `chinese_font_name` | string | 标签中文字体 |
| `english_font_name` | string | 标签英文字体 |
| `font_size` | string/number | 标签字号 |
| `bold` | bool | 标签是否加粗 |
| `italic` | bool | 标签是否斜体 |
| `font_color` | string | 标签字体颜色 |

#### 关键词专用字段（keywords.chinese / keywords.english 层级，不包括 label）

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `label` | KeywordLabelConfig | 见上 | 关键词标签字符格式 |
| `count_min` | int | `4` | 最少关键词数量（必须 > 0） |
| `count_max` | int | `4` | 最多关键词数量（必须 > 0） |
| `trailing_punct_forbidden` | bool | `true` | 是否禁止最后一个关键词后有标点符号（仅中文） |

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

> **段前段后间距**：设置 `space_before: "0行"` 和 `space_after: "0行"` 时，工具会同时写入 `w:beforeLines="0"` 和 `w:before="0"`，确保覆盖样式级的 pt 间距。避免 Heading 2/3 内置样式自带的 pt 间距回退。

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

### ⑦ tables（表格题注及表格内容配置）

继承 global_format 的 15 个字段 + 以下 2 个题注专用字段 + 1 个表格内容子配置：

| 字段 | 类型 | 默认值 | 可选值 | 说明 |
|------|------|--------|--------|------|
| `caption_position` | string | `above` | `above`（表注在表上方）`below`（表注在表下方） | 表注位置 |
| `caption_prefix` | string | `表` | 任意字符串，如 `表` `Table` `Tab.` | 表注编号前缀 |
| `content` | TableContentConfig | 见下 | 继承 15 个字段的子配置 | 表格内容格式 |

#### 表格内容子配置（content）

`tables.content` 控制表格单元格内文字的格式，继承 `GlobalFormatConfig` 的 15 个字段。在格式化模式（`wf af`）下，工具会遍历所有表格的所有单元格段落，应用此配置。

```yaml
tables:
  <<: *global_format
  caption_position: 'above'
  caption_prefix: '表'
  font_size: '五号'
  alignment: '居中对齐'
  first_line_indent: '0字符'
  builtin_style_name: '题注'
  content:                           # 表格内容格式
    <<: *global_format
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '五号'
    line_spacingrule: '单倍行距'
    alignment: '居中对齐'
    first_line_indent: '0字符'
    space_before: "0行"
    space_after: "0行"
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
    first_line_indent: '-2.2字符'   # 悬挂缩进
    left_indent: '0.26字符'         # 文本之前
    entry_indent: 0.0               # [预留]
    entry_ending_punct: '.'         # [预留]
    numbering_format: '[1], [2], ...'  # [预留]
```

#### references.title 专用字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `section_title` | string / null | `参考文献` | 参考文献章节的标题文字，设为 `null` 可不限制 |

#### references.content 专用字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `first_line_indent` | string | 继承 global | **悬挂缩进**：设负值如 `"-2.2字符"` 实现参考文献条目悬挂缩进 |
| `left_indent` | string | 继承 global | 文本之前（配合悬挂缩进使用，如 `"0.26字符"`） |
| `numbering_format` | string / null | `null` | **[预留]** 编号格式描述，当前暂未启用 |
| `entry_indent` | float / null | `0.0` | **[预留]** 条目首行缩进量，当前暂未启用 |
| `entry_ending_punct` | string / null | `null` | **[预留]** 条目结束标点，当前暂未启用 |

> **悬挂缩进示例**：`first_line_indent: '-2.2字符'` + `left_indent: '0.26字符'` — 第一行在 0.26 字符处，后续行在 2.46 字符（0.26 + 2.2）处。

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

### ⑩ numbering（标题自动编号 + 参考文献编号配置）

控制标题和参考文献的自动编号功能，包括编号格式、编号与文字的间距、缩进等。**仅在格式化模式（`wf af`）下生效**，检查模式（`wf cf`）不会修改编号。

结构如下：

```yaml
numbering:
  enabled: true                      # 总开关
  level_1:                           # 一级标题编号
    enabled: true
    template: '%1'                   # 编号模板
    suffix: 'space'                  # 编号之后的分隔符
    numbering_indent:                # [可选] 编号缩进
    text_indent:                     # [可选] 文本缩进（悬挂缩进）
  level_2:                           # 二级标题编号
    enabled: true
    template: '%1.%2'
    suffix: 'space'
  level_3:                           # 三级标题编号
    enabled: true
    template: '%1.%2.%3'
    suffix: 'space'
  references:                        # 参考文献条目编号
    enabled: true
    template: '[%1]'
    suffix: 'space'
    text_indent:                     # 悬挂缩进，如 '2.2字符'
```

#### numbering 顶层字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `enabled` | bool | `false` | 是否启用自动编号功能（总开关） |

#### NumberingLevelConfig 字段（level_1/level_2/level_3/references 共用）

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `enabled` | bool | `false` | 是否启用该级别自动编号 |
| `template` | string / null | `null` | 编号模板，`%1`=本级序号，`%2`=上级序号，`%3`=上上级序号 |
| `suffix` | string | `"space"` | 编号之后的分隔符：`tab`（制表符）、`space`（空格）、`nothing`（无） |
| `numbering_indent` | string / null | `null` | [可选] 编号缩进（编号距左边距的距离），如 `"0.75cm"` |
| `text_indent` | string / null | `null` | [可选] 文本缩进（悬挂缩进量），如 `"0.75cm"`、`"2.2字符"` |

#### 编号模板（template）说明

模板使用 `%1`、`%2`、`%3` 占位符表示各级序号：

| 模板 | 效果示例 | 适用场景 |
|------|---------|---------|
| `'%1'` | `1 绪论`、`2 系统设计` | 阿拉伯数字章标题 |
| `'第%1章'` | `第一章 绪论`、`第二章 系统设计` | 中文数字章标题 |
| `'%1.%2'` | `1.1 研究背景`、`2.3 系统架构` | 节标题 |
| `'%1.%2.%3'` | `1.1.1 研究背景`、`2.3.4 系统架构` | 小节标题 |
| `'[%1]'` | `[1] 作者. 标题...`、`[2] 作者. 标题...` | 参考文献条目 |

> **注意**：`template` 中包含 `第` 和 `章` 时，编号格式自动使用中文计数（`chineseCountingThousand`），否则使用阿拉伯数字（`decimal`）。

#### 自动清除手动编号

格式化时会根据标题级别自动识别并清除段落开头的手动编号文字，无需配置正则表达式。

支持的常见编号格式（按标题级别）：

**一级标题**：`第一章`、`第1章`、`第一节`、`（一）`、`(1)`、`一、`、`一.`、`1)`、`I.`
**二级标题**：`1.1`、`1.1 `、`一.1`
**三级标题**：`1.1.1`、`1.1.1 `

> 清除手动编号后会自动应用 `template` 指定的新编号样式。

#### 编号之后（suffix）说明

控制编号文字与正文之间的分隔符：

| 值 | 效果 | 示例 |
|----|------|------|
| `tab` | 制表符（Word 默认） | `1`⇥`绪论`（间距较大） |
| `space` | 空格 | `1 绪论`（间距紧凑） |
| `nothing` | 无分隔 | `1绪论`（无间距） |

#### 缩进设置说明

`numbering_indent` 和 `text_indent` 支持多种单位：

| 字段 | 含义 | 支持单位 | 示例 |
|------|------|---------|------|
| `numbering_indent` | 编号起始位置距左边距的距离 | 厘米/cm、毫米/mm、英寸/inch、磅/pt、字符 | `'0字符'`、`'0.75cm'`、`'420磅'` |
| `text_indent` | 文本相对于编号起始位置的悬挂缩进 | 同上 | `'0字符'`、`'0.75cm'`、`'420磅'`、`'2.2字符'` |

> **字符单位特殊说明**：字符单位使用 Word 底层的 `w:leftChars` / `w:hangingChars` 属性（1字符 = 100 单位），与物理单位（`w:left` / `w:hanging`，单位为 twips）是独立的属性。

#### 编号样式自动跟随标题

编号的字体、字号、加粗等样式会**自动跟随对应级别标题的 `headings` 配置**，无需单独设置。例如：
- `level_1` 的编号样式跟随 `headings.level_1` 的字体/字号/加粗
- `level_2` 的编号样式跟随 `headings.level_2` 的字体/字号/加粗

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
│   └── keywords.chinese/english   → 继承 15 字段 + label(15字段) + 3 个关键词专用字段
├── headings
│   ├── level_1              → 继承 15 字段
│   ├── level_2              → 继承 15 字段
│   └── level_3              → 继承 15 字段
├── body_text                → 继承 15 字段
├── figures                  → 继承 15 字段 + 2 个图表专用字段
├── tables                   → 继承 15 字段 + 2 个图表专用字段 + content(15字段)
├── references
│   ├── title                → 继承 15 字段 + 1 个专用字段（section_title）
│   └── content              → 继承 15 字段 + 3 个预留字段
├── acknowledgements
│   ├── title                → 继承 15 字段
│   └── content              → 继承 15 字段
└── numbering                 → 独立配置（不继承 GlobalFormat）
    ├── enabled              → bool 总开关
    ├── level_1              → enabled, template, suffix, numbering_indent, text_indent
    ├── level_2              → 同上
    ├── level_3              → 同上
    └── references           → 同上
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
| `abstract.keywords.chinese` | 同 global_format + label(同 global_format) + count_min, count_max, trailing_punct_forbidden |
| `abstract.keywords.english` | 同 global_format + label(同 global_format) + count_min, count_max, trailing_punct_forbidden |
| `abstract.keywords.chinese.label` | 同 global_format |
| `abstract.keywords.english.label` | 同 global_format |
| `headings.level_1/2/3` | 同 global_format |
| `body_text` | 同 global_format |
| `figures` | 同 global_format + caption_position, caption_prefix |
| `tables` | 同 global_format + caption_position, caption_prefix, content(同 global_format) |
| `tables.content` | 同 global_format |
| `references.title` | 同 global_format + section_title |
| `references.content` | 同 global_format + entry_indent, entry_ending_punct, numbering_format |
| `acknowledgements.title` | 同 global_format |
| `acknowledgements.content` | 同 global_format |
| `numbering` | enabled |
| `numbering.level_1/2/3/references` | enabled, template, suffix, numbering_indent, text_indent |
