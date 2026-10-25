# 配置文件说明

本文档详细说明 WordFormat 的配置文件格式和自定义配置方法，帮助您根据不同学校/期刊的格式要求定制专属格式规范。

## 配置文件格式

格式规范通过 YAML 文件定义，示例参考项目中：
- `example/undergrad_thesis.yaml`（本科毕业论文模板）

也可以导出一份包含全部节点与缺省值的空白模板，在此基础上修改：

```bash
wordf config -o my_config.yaml
```

样式字段统一使用平铺写法，段落与字符属性直接写在节点下：

```yaml
headings:
  level_1:
    alignment: '居中对齐'
    chinese_font_name: '黑体'
```

不再支持 `paragraph:` / `font:` 嵌套写法；配置加载时会直接报错并指出对应节点。

所有配置项均有清晰注释，便于理解和修改。

## 核心配置项

### 1. 全局格式配置（style_checks_warning）
用于控制格式校验时，哪些属性不满足规范时触发警告提示。
```yaml
style_checks_warning:
  bold: true
  italic: true
  underline: true
  font_size: true
  font_name: false
  font_color: false
  alignment: true
  space_before: true
  space_after: true
  line_spacing: true
  line_spacingrule: true
  left_indent: true
  right_indent: true
  first_line_indent: true
  builtin_style_name: true
```

### 2. 全局基础格式（global_format）
所有段落格式默认继承的基础样式，可通过锚点 `&global_format` 被其他节点复用。
```yaml
global_format: &global_format
  alignment: '两端对齐'
  space_before: "0.5行"
  space_after: "0.5行"
  line_spacingrule: "1.5倍行距"
  line_spacing: '1.5倍'
  left_indent: "0字符"
  right_indent: "0字符"
  first_line_indent: '2字符'
  builtin_style_name: '正文'
  chinese_font_name: '宋体'
  english_font_name: 'Times New Roman'
  font_size: '小四'
  font_color: '黑色'
  bold: false
  italic: false
  underline: false
```

### 3. 摘要及关键词配置（abstract）
包含中文摘要、英文摘要及对应关键词的统一配置节点。标题用 `title`、正文用 `body`（不是 `chinese_title`/`chinese_content`）；
关键词是 `chinese`/`english` 下的 `keywords` 子节（不是顶层 `abstract.keywords`），并通过 `label` 子配置独立控制标签（"关键词：" / "Keywords:"）与内容的格式。
```yaml
abstract:
  chinese:
    title:                            # 中文摘要标题
      <<: *global_format
      alignment: '居中对齐'
      first_line_indent: '0字符'
      chinese_font_name: '黑体'
      font_size: '小二'
      bold: true
    body:                             # 中文摘要正文
      <<: *global_format
      alignment: '两端对齐'
    keywords:                         # 中文关键词
      <<: *global_format
      alignment: '左对齐'
      first_line_indent: '0字符'
      font_size: '四号'
      label:                          # 标签（"关键词："）格式
        <<: *global_format
        chinese_font_name: '黑体'
        font_size: '四号'
        bold: true
      rules:                          # 业务规则（均通过 enabled 开关控制）
        keyword_count:
          enabled: true
          count_min: 3
          count_max: 5
        trailing_punctuation:
          enabled: true
          forbidden_chars: "；，。、"
  english:
    title:                            # 英文摘要标题
      <<: *global_format
      alignment: '居中对齐'
      first_line_indent: '0字符'
      font_size: '四号'
      bold: true
    body:                             # 英文摘要正文
      <<: *global_format
      alignment: '两端对齐'
    keywords:                         # 英文关键词
      <<: *global_format
      alignment: '左对齐'
      first_line_indent: '0字符'
      font_size: '四号'
      label:                          # 标签（"Keywords:"）格式
        <<: *global_format
        font_size: '四号'
        bold: true
      rules:
        keyword_count:
          enabled: true
          count_min: 3
          count_max: 5
```

> `title` 与 `body` 只在标题、正文分属不同段落时生效。若摘要标题与正文写在同一段（常见于"摘  要：xxx"混排），命中的是 `abstract.chinese` / `abstract.english` 这两个 `title_content` 节点，它会按 run 边界拆分后分别套用 `title` 与 `body` 的格式。

### 4. 各级标题格式（headings）
配置一级、二级、三级标题的样式，对应 Word 内置标题样式。
```yaml
headings:
  level_1:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小二'
    bold: false
    builtin_style_name: 'Heading 1'

  level_2:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '三号'
    bold: false
    builtin_style_name: 'Heading 2'

  level_3:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小四'
    bold: false
    builtin_style_name: 'Heading 3'
```

### 5. 正文段落格式（body）
论文正文主体内容格式，直接继承全局格式，无需重复定义。节点路径是 `body.text`。
```yaml
body:
  text:
    <<: *global_format
    rules:
      punctuation:
        enabled: true
      paragraph_length:        # 字数/句数阈值，0 = 不限制（默认全部不限制）
        enabled: true
        min_chars: 0           # 每段最少字数
        max_chars: 0           # 每段最多字数
        min_sentences: 0       # 每段最少句数
        max_sentences: 0       # 每段最多句数
```

| 配置项 | 说明 | 缺省 |
|--------|------|------|
| `paragraph_length.enabled` | 字数/句数检查开关 | `true` |
| `min_chars` / `max_chars` | 单段字数上下限；字数为去掉空白后的字符数（含标点） | `0`（不限制） |
| `min_sentences` / `max_sentences` | 单段句数上下限；按 `。！？!?` 切分，末尾无标点的残句也算一句 | `0`（不限制） |

> 阈值为 `0` 时不检查，所以默认配置不会产生批注；填了上限就按上限报「字数过多/句数过多」，填报下限则报「字数过少/句数过少」。检查只报批注，不改文字。

### 6. 插图格式（figures）
插图有两层配置：题注（`figures.caption`）与图片段落（`figures.image`）。图片段落只检查对齐、首行缩进与版式（环绕方式、是否超出版心）。
```yaml
figures:
  caption:
    <<: *global_format
    caption_prefix: '图'
    font_size: '五号'
    builtin_style_name: '题注'
    alignment: '居中对齐'
    first_line_indent: '0字符'
    rules:
      caption_numbering:
        enabled: true
        separator: '.'
        label_number_space: false
  image:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    rules:
      figure_layout:
        enabled: true
        wrap: '嵌入型'      # 可选：不写则只检查宽度
        max_width: '版心'   # 版心 / 带单位的宽度（如 15cm、150mm）
```

| 配置项 | 说明 | 缺省 |
|--------|------|------|
| `figure_layout.enabled` | 图片版式检查总开关 | `true` |
| `figure_layout.wrap` | 期望环绕方式：嵌入型、四周型、紧密型、穿越型、上下型、衬于文字下方、浮于文字上方 | 不校验 |
| `figure_layout.max_width` | 宽度上限：`版心`（段落所属节的页宽 − 左右页边距）或带单位数值 | `版心` |

> 版式检查只报批注、不改图片版式，check 与 apply 都会报。宽度取自 `wp:extent`，与段落所属节的版心比较；多节文档（如横向附录）按各自节的页边距计算。

### 7. 表格格式（tables）
表格有三层配置：题注格式（`tables.caption`）、单元格与网格样式（`tables.styles` 档案）、编号规则。
单元格内文字的格式由档案的 `cell_run` 控制，**没有 `tables.content` 这个配置项**。

```yaml
tables:
  default: 'three_line'                 # 未指定 table_style 时使用的档案名
  styles:
    three_line:
      style: 'Table Grid'               # 建表时使用的 Word 内置表样式
      layout: 'fixed'                   # fixed=固定列宽；其他值=自动布局
      width: { type: 'pct', value: 100 }        # type: auto / pct / cm / dxa
      alignment: 'center'               # left / center / right
      indent_cm: 0.2                    # 表格整体左缩进
      borders:                          # 边名：top/left/bottom/right/inside_h/inside_v
        top: { style: 'double', size_pt: 1.5, color: '000000' }
        bottom: { style: 'double', size_pt: 1.5, color: '000000' }
        inside_h: { style: 'single', size_pt: 0.75 }
        left: { style: 'none' }         # none 显式清边；简写 top: 'double' 亦可
        right: 'none'
        inside_v: 'none'
      cell_margins_cm: { top: 0.05, bottom: 0.05, left: 0.2, right: 0.2 }
      col_widths: { mode: 'custom', custom: [2.5, 5, 5] }   # mode: auto/equal/custom，单位 cm
      header_row: { repeat: true, cant_split: true }         # 跨页重复表头 / 行禁止拆页
      row_bands:                        # 行级覆盖，按列表顺序应用
        - { rows: [0], fill: 'F2F2F2', bold: true }           # rows: all / 整数 / 整数列表
        - { rows: [1, 3], fill: 'FAFAFA', borders: { top: 'single' } }
      cell_run:                         # 单元格文字格式，只应用显式写出的字段
        chinese_font_name: '宋体'
        english_font_name: 'Times New Roman'
        font_size: '五号'
        font_color: '黑色'
  caption:                              # 表格题注（与插图题注同构）
    alignment: '居中对齐'
    first_line_indent: '0字符'
    builtin_style_name: '题注'
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '五号'
    caption_prefix: '表'
    rules:
      caption_numbering:
        enabled: true
        separator: '.'
        label_number_space: false
```

与实现一致的几点约束：
- `tables.default` 与 JSON 条目的 `table_style` 都没写时，表格保持 `Table Grid` 现状不动；写了 `tables.styles` 里不存在的档案名会直接报错并列出可用名。
- 档案只在**新建表格**时生效（`md`/`jd`，或 `af` 中条目带 `table_rows`）；`af` 里 `exists: true` 的已有表格是只读的，不改列宽、底纹和边框。
- `tables.object` 下的段落/字体配置不影响单元格；单元格段落自身的对齐与行距没有配置项。

### 8. 参考文献格式（references）
包含参考文献标题（`references.title`）与条目（`references.entry`，不是 `content`）的格式。条目支持悬挂缩进（`first_line_indent` 设为负值）。

```yaml
references:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '三号'
    bold: true
  entry:
    <<: *global_format
    alignment: '两端对齐'
    first_line_indent: '-2.2字符'   # 悬挂缩进（负值）
    left_indent: '0.26字符'         # 文本之前
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '五号'
    rules:
      gbt_7714:                     # 著录项校验（只报批注，不改文字）
        enabled: true               # 总开关
        require_type_marker: true   # 必须有 [J]/[M]/[D]/[EB/OL]…
        check_journal_item: true    # [J] 查出版年 + 卷(期)/页码
        check_publisher_item: true  # [M]/[C]/[D]/[G]/[R] 及离线 [DB]/[CP]/[EB] 查出版项
        check_newspaper_item: true  # [N] 查报纸名 + 出版日期 + 版次
        check_standard_item: true   # [S] 查标准编号 + 出版项
        check_patent_item: true     # [P] 查专利号 + 公告/公开日期
        check_electronic_item: true # /OL 查引用日期 + 获取路径
        check_author_ellipsis: true # 查省略号与中文条目误用 et al
        check_terminator: false     # 末尾半角句点（各校分歧大）
        check_punctuation_width: false  # 著录符号全角（题名内全角合法）
```

> 条目的编号不在这里配置：`[1]` 这类编号由 `numbering.references` 统一控制（见 §10），框架会先剥离条目里的手写编号再套自动编号。
> `gbt_7714` 按 GB/T 7714-2015 校验著录字段，结果只写成批注（check 与 apply 都会报），不会改写条目文字。每项校验可用 `false` 单独关闭；`check_terminator` 与 `check_punctuation_width` 因各校要求不一，默认关闭。

| 配置项 | 检查内容 | 缺省 |
|--------|----------|------|
| `enabled` | 整条规则的总开关 | `true` |
| `require_type_marker` | 文献类型标识是否存在、主类型是否合法 | `true` |
| `check_journal_item` | `[J]`：出版年必查；页码必查（`/OL` 联机版可免）；卷(期)与页码至少一项 | `true` |
| `check_publisher_item` | `[M]`/`[C]`/`[D]`/`[G]`/`[R]`：出版地:出版者,出版年；联机（`/OL`）可省略出版项，离线载体 `[DB]`/`[CP]`/`[EB]` 同样要求出版项 | `true` |
| `check_newspaper_item` | `[N]`：报纸名 + 出版日期（仅 `YYYY-MM-DD`）+ `(版次)` | `true` |
| `check_standard_item` | `[S]`：标准编号（GB/T、ISO、IEEE…）+ 出版项 | `true` |
| `check_patent_item` | `[P]`：专利号（CN、ZL…）+ 公告/公开日期（仅 `YYYY-MM-DD`） | `true` |
| `check_electronic_item` | `/OL`：引用日期与获取路径（网址或 DOI），对 `[DB/OL]`/`[CP/OL]`/`[EB/OL]` 等全部联机载体生效 | `true` |
| `check_author_ellipsis` | 用省略号代替“，等”/“, et al”；中文条目误用 et al | `true` |
| `check_terminator` | 条目末尾是否有半角句点 | `false` |
| `check_punctuation_width` | 是否混用全角 `：，；（）．` | `false` |

文献类型与著录项校验的对应关系（`_TYPE_CHECKS`）：

| 类型 | 必查著录项 |
|------|-----------|
| `[J]` 期刊 | 出版年、页码；卷(期)与页码至少一项（联机 `/OL` 可免页码） |
| `[M]` 专著 / `[C]` 论文集 / `[D]` 学位论文 / `[G]` 汇编 / `[R]` 报告 | 出版地:出版者,出版年（联机 `/OL` 可省略） |
| `[N]` 报纸 | 报纸名、`YYYY-MM-DD` 出版日期、`(版次)` |
| `[S]` 标准 | 标准编号、出版项 |
| `[P]` 专利 | 专利号、`YYYY-MM-DD` 公告/公开日期 |
| `[DB]` 数据库 / `[CP]` 计算机程序 / `[EB]` 电子公告 | 联机（`/OL`）查引用日期与获取路径；离线载体查出版项 |

### 9. 致谢格式（acknowledgements）
配置致谢章节的标题（`acknowledgements.title`）与正文（`acknowledgements.body`，不是 `content`）格式。
```yaml
acknowledgements:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '小二'
    bold: true
  body:
    <<: *global_format
    alignment: '两端对齐'
    first_line_indent: '2字符'
```

### 10. 标题自动编号（numbering）
控制标题和参考文献条目的自动编号功能，包括编号格式、分隔符、缩进等。仅在格式化模式（`wordf af`）下生效。

```yaml
numbering:
  enabled: true
  level_1:
    enabled: true
    template: '%1'
    suffix: 'space'                 # tab / space / nothing
    numbering_indent:               # [可选] 编号缩进，如 '0字符'
    text_indent:                    # [可选] 文本悬挂缩进
  level_2:
    enabled: true
    template: '%1.%2'
    suffix: 'space'
  level_3:
    enabled: true
    template: '%1.%2.%3'
    suffix: 'space'
  references:
    enabled: true
    template: '[%1]'
    suffix: 'space'
    text_indent:                    # 悬挂缩进，如 '2.2字符'
```

**编号模板（template）**：使用 `%1`、`%2`、`%3` 占位符表示各级序号，常用格式如 `'%1'`（1 绪论）、`'%1.%2'`（1.1 研究背景）、`'第%1章'`（第一章 绪论）、`'[%1]'`（[1] 参考文献条目）。

**编号之后（suffix）**：控制编号与文字之间的分隔符，可选 `tab`（制表符）、`space`（空格）、`nothing`（无）。

**缩进设置**：`numbering_indent`（编号缩进）和 `text_indent`（文本悬挂缩进）为可选字段，支持厘米、毫米、英寸、磅、字符等多种单位。

> 编号的字体、字号、加粗等样式会自动跟随对应级别标题的 `headings` 配置，无需单独设置。

### 11. 页眉页脚与页面设置（header_footer）
页眉页脚存在于各节（section）的独立 part 中，不属于正文段落流，因此不走模型分类和节点树，
由 `builtin.header_footer` 机制在文档保存前逐节处理。它只**重排已有内容的格式**并写入页面设置，
不会改写正文文字；apply 模式额外能在页眉页脚里插入 PAGE 页码域。

```yaml
header_footer:
  enabled: true                   # false 时 check 与 apply 都整段跳过
  odd_even_pages: false           # 可选：奇偶页不同（文档级，写 settings.xml）
  update_fields_on_open: false    # 可选：打开文档时更新域（文档级，写 settings.xml）
  header:
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '小五'
    alignment: '居中对齐'
    require_content: false        # true 时空页眉报「内容缺失」
    page_number: false            # true 时要求存在 PAGE 页码域
  footer:
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '小五'
    alignment: '居中对齐'
    require_content: true
    page_number: true
  page_setup:                     # 可选：不写则完全不校验页面设置
    top_margin: '2.54cm'
    bottom_margin: '2.54cm'
    left_margin: '3.18cm'
    right_margin: '3.18cm'
    header_distance: '1.5cm'      # 页眉距边界
    footer_distance: '1.75cm'     # 页脚距边界
    page_size: 'A4'               # A3 / A4 / A5 / B5 / Letter
    orientation: '纵向'           # 纵向 / 横向
    page_number_start: 1          # 本节起始页码
    page_number_format: '小写罗马数字'
    different_first_page: true    # 首页不同（逐节）
```

| 配置项 | 写入位置 | 说明 |
|--------|----------|------|
| `enabled` | — | 页眉页脚与页面设置处理的总开关，缺省 `true` |
| `header` / `footer` 的四个格式项 | 部件内 run/段落的 `w:rPr`、`w:jc` | 字号、中文字体、英文字体、对齐；不检查加粗/颜色/行距，避免误伤页码域 |
| `require_content` | — | 该部位为空时是否报告缺失，缺省 `false`；只报告，不自动补文字 |
| `page_number` | `w:fldSimple w:instr=" PAGE "` | 该部位是否必须有页码域：check 报缺失，apply 在首段插入（并按本部位配置补齐对齐与字体） |
| `page_setup` 的六项边距 | `w:pgMar` | 单位支持 cm / mm / in / pt，缺省按 cm；因各校边距差异大而设为选填 |
| `page_size` / `orientation` | `w:pgSz` | 只写 `orientation` 而不写 `page_size` 时，对调现有宽高 |
| `page_number_start` / `page_number_format` | `w:pgNumType` 的 `w:start` / `w:fmt` | 格式支持中文标签（阿拉伯数字、小写罗马数字…）与 OOXML 原值（decimal、upperRoman…） |
| `different_first_page` | `w:titlePg` | 逐节开关，对应 Word「首页不同」 |
| `odd_even_pages` / `update_fields_on_open` | `settings.xml` 的 `w:evenAndOddHeaders` / `w:updateFields` | 文档级，只校验/写入一次；配置了就显式开或关，未配置则完全不碰（保留文档原状） |

- `header`/`footer` 只需写要规范的字段，其余与内置缺省（小五号宋体 + Times New Roman + 居中）深合并，用户配置优先。
- 页面设置类字段全部**选填**：没写的项不校验也不修改，避免对每个文档误报。
- 逐节处理时，第 2 节起没有独立部件的部位沿用上一节，会被跳过；首页/偶页变体只在对应开关已开启
  且该变体确实有部件时才检查——Word 不显示的部位报缺失只会误伤。检查模式不会凭空创建部件。
- 插入的页码域是域代码，页码数字要等 Word 刷新域才显示；配合 `update_fields_on_open: true` 可在打开文档时自动更新。
- 页码域缺失与内容缺失不重复报：`page_number: true` 时页码域本身就是内容。
- 检查出的差异以批注锚定在文档第一段——python-docx 限制批注只能挂在 document part，页眉页脚里的 run 无法直接挂批注。
- `md`/`jd` 新生成的文档页眉页脚为空，正文文字不会被自动写入；但 `page_number: true` 会在 apply 时创建部件并插入页码域。

## 字段详细说明

### 格式警告控制字段（style_checks_warning）
| 配置项 | 说明 | 取值 |
|--------|------|------|
| bold | 是否对加粗不规范进行警告 | true/false |
| italic | 是否对斜体不规范进行警告 | true/false |
| underline | 是否对下划线不规范进行警告 | true/false |
| font_size | 是否对字号不规范进行警告 | true/false |
| font_name | 是否对字体名称不规范进行警告 | true/false |
| font_color | 是否对字体颜色不规范进行警告 | true/false |
| alignment | 是否对齐方式不规范进行警告 | true/false |
| space_before | 是否对段前间距不规范进行警告 | true/false |
| space_after | 是否对段后间距不规范进行警告 | true/false |
| line_spacing | 是否对行距数值不规范进行警告 | true/false |
| line_spacingrule | 是否对行距类型不规范进行警告 | true/false |
| left_indent | 是否对左缩进不规范进行警告 | true/false |
| right_indent | 是否对右缩进不规范进行警告 | true/false |
| first_line_indent | 是否对首行缩进不规范进行警告 | true/false |
| builtin_style_name | 是否对内置样式不规范进行警告 | true/false |

### 段落格式字段
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| alignment | 段落对齐方式 | - | 左对齐、居中对齐、右对齐、两端对齐、分散对齐 |
| space_before | 段前间距 | 行/磅/毫米/厘米/英寸 | 0.5行、12磅、0.5cm |
| space_after | 段后间距 | 行/磅/毫米/厘米/英寸 | 0.5行、12磅、0.5cm |
| line_spacingrule | 行距规则 | - | 单倍行距、1.5倍行距、2倍行距、最小值、固定值、多倍行距 |
| line_spacing | 行距数值 | 倍 | 1倍、1.5倍、2倍 |
| left_indent | 左缩进 | 字符/磅/毫米/厘米/英寸 | 0字符、2字符、20磅 |
| right_indent | 右缩进 | 字符/磅/毫米/厘米/英寸 | 0字符、2字符 |
| first_line_indent | 首行缩进（正值）/ 悬挂缩进（负值） | 字符/磅/毫米/厘米/英寸 | 2字符、-2.2字符 |
| builtin_style_name | Word内置样式名 | - | 正文、Heading 1、Heading 2、题注 |

### 字符格式字段
| 配置项 | 说明 | 可选值 |
|-------|------|--------|
| chinese_font_name | 中文字体 | 宋体、黑体、楷体、仿宋、微软雅黑、汉仪小标宋 |
| english_font_name | 英文字体 | Times New Roman、Arial、Calibri、Courier New、Helvetica |
| font_size | 字号 | 一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、七号，或数值（如 12、14） |
| font_color | 字体颜色 | 黑色、红色、十六进制色值 |
| bold | 是否加粗 | true/false |
| italic | 是否斜体 | true/false |
| underline | 是否下划线 | true/false |

### 关键词专用字段
| 配置项 | 说明 | 取值 |
|-------|------|--------|
| label | 关键词标签（"关键词：" / "Keywords:"）的字符格式 | 子配置（继承 global_format 的 16 个字段） |
| label.bold | 标签是否加粗 | true/false |
| label.chinese_font_name | 标签中文字体 | 宋体、黑体 等 |
| label.font_size | 标签字号 | 四号、小四 等 |
| rules.keyword_count | 关键词数量校验规则 | 子配置 |
| rules.keyword_count.enabled | 是否启用数量校验 | true/false |
| rules.keyword_count.count_min | 最少关键词数量 | 正整数 |
| rules.keyword_count.count_max | 最多关键词数量 | 正整数 |
| rules.trailing_punctuation | 末尾标点校验规则（仅中文） | 子配置 |
| rules.trailing_punctuation.enabled | 是否启用末尾标点校验 | true/false |
| rules.trailing_punctuation.forbidden_chars | 禁止出现在末尾的标点 | 字符串 |

### 题注专用字段
| 配置项 | 说明 | 可选值 |
|-------|------|--------|
| caption_prefix | 题注前缀 | 图、表 |
| rules.caption_numbering | 题注编号校验/修正规则 | 子配置 |
| rules.caption_numbering.enabled | 是否启用编号校验 | true/false |
| rules.caption_numbering.separator | 章节号与编号间分隔符 | . - : — – |
| rules.caption_numbering.label_number_space | 标签与编号间是否加空格 | true/false |

### 参考文献专用字段
参考文献标题与条目只使用通用的段落/字符格式字段（见上文两张字段表），没有专有字段。
条目编号格式来自 `numbering.references`（见 §10 编号专用字段），悬挂缩进来自 `references.entry.first_line_indent` 负值。

> 以下字段代码中无读取点，属旧文档遗留，写了不生效：`section_title`、`entry_indent`、`entry_ending_punct`、`numbering_format`。

### 编号专用字段
| 配置项 | 说明 | 可选值 |
|-------|------|--------|
| enabled | 是否启用自动编号 | true/false |
| template | 编号模板 | '%1'、'%1.%2'、'%1.%2.%3'、'第%1章'、'[%1]' 等 |
| suffix | 编号后分隔符 | tab、space、nothing |
| numbering_indent | [可选] 编号缩进 | 带单位的值，如 '0字符'、'0.75cm' |
| text_indent | [可选] 文本悬挂缩进 | 带单位的值，如 '0字符'、'0.75cm'、'2.2字符' |

## 配置继承机制
使用 YAML 锚点 `&` 与引用 `<<:` 实现样式复用：
```yaml
global_format: &global_format
  alignment: '两端对齐'

headings:
  level_1:
    <<: *global_format
    alignment: '居中对齐'
```

## 配置验证
配置加载时校验 YAML 语法合法性。各字段的默认值由 FormatNode 子类的 `DEFAULTS` 属性提供，YAML 中只需写需要覆盖的字段。使用 `wordf config -o config.yaml` 可导出包含所有字段和默认值的完整配置模板。

## 自定义配置流程
1. 复制 `example/` 下现有模板
2. 修改对应格式节点参数
3. 保存为新 YAML 文件使用
4. 执行命令校验配置有效性

## 故障排查
- 配置加载失败：检查 YAML 缩进、引号、冒号是否规范
- 格式不生效：确认段落类别识别正确、配置路径正确
- 警告异常：调整 `style_checks_warning` 中对应开关
- 样式不匹配：检查内置样式名称 `builtin_style_name` 是否正确
