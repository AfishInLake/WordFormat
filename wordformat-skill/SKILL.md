---
name: wordformat
description: 论文格式自动化处理工具。在处理 Word 论文文档格式校验、格式修正、文档结构识别、Markdown 转 Word 场景时激活。
argument-hint: "[文件路径]"
---

# WordFormat

## 安装

```bash
pip install wordformat
```

验证：`wordf --help`

## 按需使用原则（重要）

每个命令只做一件事。请**只执行满足用户需求所需的最小命令组合**，不要执行多余的步骤：

- 用户说"修标签"就修标签（gj → 改 JSON → 需要时 af），不要顺带跑 cf 检查格式
- 用户说"检查格式"就只产出检查结果（gj → cf），不执行 af 改文档
- 用户说"格式化"就直接格式化（gj → 改必要的标签 → af）；cf（加批注）不是 af 的前置步骤
- `gj` 每次运行会**重新识别并覆盖** JSON——已手动改过 JSON 后不要再重跑 gj
- output/ 下已有 JSON 时直接复用，不重复生成

## 开始前：先确认格式来源（必问）

在执行 `cf` / `af` / `jd` / `md` 之前（`gj` 带 `-c` 时同理），先问用户一句：**学校（或学院、期刊）有没有下发格式规定文档**（Word 模板、格式要求正文、规范截图都算）。一次问清，不要边跑边追问；同一篇文档已经确认过就不再重复问，用户说"按默认来"就直接走默认分支。

- **有规定文档**：从中提取字体、字号、行距、段前段后、缩进、页边距、页眉页脚、题注编号、参考文献著录等要求，写成 YAML 配置，再用 `-c my_config.yaml` 执行。提取不到或有歧义的字段要明确告诉用户"这项按默认值处理"，不要臆造学校要求。
- **没有规定文档**：使用默认格式——`wordf config -o config.yaml` 导出内置缺省值，用它作为 `-c` 执行（仓库里的 `example/undergrad_thesis.yaml` 是同口径的本科模板，可直接取用）。不要凭空编一套"学校标准"写进配置。

## 决策表：根据用户意图选择命令

| 用户需求 | 必要命令 | 可选命令 | 不要做 |
|---|---|---|---|
| 修正/查看识别分类（标签） | 改 JSON 的 `category`；无 JSON 时先 `gj`；要应用到文档再 `af` | `tree --confidence` 定位可疑段 | 不跑 `cf`；改完 JSON 不重跑 `gj` |
| 替换正文文字 | 改 JSON 的 `replace`；无 JSON 时先 `gj`；再 `af` | — | 不跑 `cf`、`tree` |
| 只检查格式（加批注） | `gj`（如无 JSON）→ `cf` | — | 不跑 `af` |
| 修正格式（应用到文档） | `gj`（如无 JSON）→ `af` | `tree` 先复核分类 | `cf` 非必需，用户想看批注才插入 |
| 按 JSON 重建整篇文档（内容全由 JSON 承载） | `gj` → 编辑 JSON → `jd` | — | 不跑 `af`（与 `jd` 互斥）；重建不保留原图片二进制，原表格只以只读文本（`exists`）导出、不重建 |
| 在原 docx 上增删改段落（保留图片/表格/页眉页脚） | `gj` → 编辑 JSON → `af -d 原文档 -f 编辑后JSON --sync` | — | 图片段不可移动；原表格内容只读（不检查不改写），表格增删走 `jd` 或手动 |
| 查看结构树 | `tree -f xxx.json` | `--confidence` / `--index` / `--filter` | 需已有 JSON；不跑 `gj`（除非没有 JSON） |
| Markdown 转 Word | `md` | — | 与 docx 流程互不相关 |
| 生成配置模板 | `config` | — | — |
| 学校/学院定制规则 | 在 `presets/` 放预设 `.py`，命令加 `--preset 名` | `list-presets` 查看 | 预设是 Hook 插件，不是配置；不要塞进 YAML |

只有用户明确要求"完整流程/全套"时，才按场景顺序执行多步流水线。

## 命令速查

| 命令 | 功能 |
|------|------|
| `wordf md -d 论文.md -c config.yaml` | Markdown → 格式化 .docx |
| `wordf gj -d 论文.docx -c config.yaml` | AI 识别文档结构，输出 JSON |
| `wordf tree -f output/xxx.json` | 查看文档结构树 |
| `wordf cf -d 论文.docx -c config.yaml -f output/xxx.json` | 检查格式（加批注） |
| `wordf af -d 论文.docx -c config.yaml -f output/xxx.json` | 修正格式（直接改） |
| `wordf af -d 论文.docx -c config.yaml -f xxx.json --sync` | 按编辑后 JSON 在原 docx 上同步增/删/改/重排段落 |
| `wordf jd -f output/xxx.json -c config.yaml` | JSON 生成全新 docx（重建模式） |
| `wordf config -o config.yaml` | 输出配置模板 |
| `wordf list-presets` | 列出 `presets/` 下可用预设 |
| `wordf startapi` | 启动 Web 界面 |

`cf` / `af` / `md` / `jd` 均支持 `--preset <预设名>`（可多次指定，后者覆盖前者）；`tree` 支持 `--filter heading_level_1,body_text` 只显示指定类别。

## 常见场景最小编排

### 1. 修复识别错误的标签

用户让"看看分类对不对""帮我把识别错的标签改一下"时：

```bash
# 第 1 步：JSON 不存在才生成（已存在则直接复用，避免覆盖修改）
wordf gj -d 论文.docx -c config.yaml

# 第 2 步（可选）：低置信度的段往往是分类错误，配合 --index 拿到序号
wordf tree -f output/xxx.json --confidence --index

# 第 3 步：编辑 output/xxx.json，修正相应段落的 category 字段

# 第 4 步：仅在用户要求把修正应用到文档时执行
wordf af -d 论文.docx -c config.yaml -f output/xxx.json
```

- 用户只想"看/改标签"：到第 3 步就结束，不执行任何后续命令。
- 不执行 `cf`：格式检查与标签修正无关。
- 修改 JSON 后不重跑 `gj`：会覆盖手工修改。

### 2. 纯文本替换

同场景 1，仅字段不同：`gj` → 编辑 `replace` 字段 → `af`。不需要 `tree` 和 `cf`。

### 3. 仅格式检查（不改原文）

```bash
wordf gj -d 论文.docx -c config.yaml   # JSON 不存在时执行
wordf cf -d 论文.docx -c config.yaml -f output/xxx.json
```

`cf` 生成带批注的检查版 docx，不修改原文。到此结束，不跑 `af`。

### 4. 格式修正（生成格式化版文档）

```bash
wordf gj -d 论文.docx -c config.yaml   # JSON 不存在时执行
# 如发现识别错误的标签，先编辑 category 再继续
wordf af -d 论文.docx -c config.yaml -f output/xxx.json
```

`af` 直接产出格式化后的 docx。仅当用户想先看错误清单时，才在中间插入 `cf`。

### 5. Markdown 转 Word

```bash
wordf md -d 论文.md -c config.yaml
# 输出 output/论文--生成版.docx
```

### 6. Agent 编辑 JSON 后应用（两种模式二选一）

编辑 JSON：改 `category`/`paragraph`/`replace`，增删条目，调整条目顺序均可。

```bash
wordf gj -d 论文.docx -c config.yaml      # 1. 识别（已有 JSON 则直接编辑）
# 2a. 重建模式：内容以 JSON 为准，生成全新 docx
wordf jd -f output/xxx.json -c config.yaml
# 2b. 原地模式：在原 docx 上同步增/删/改/重排，图片/表格/页眉页脚保真
wordf af -d 论文.docx -c config.yaml -f output/xxx.json --sync
```

模式选择与边界：
- 需保留原文档对象（图片二进制、表格内容、页眉页脚）→ `af --sync`；内容全由 JSON 承载或新建文档 → `jd`。
- `jd`：新增图片用 `figure_image` 条目（`paragraph`=图片文件路径）；新增表格用 `table_object` 条目（`table_rows`=二维字符串列表，可选 `table_style` 指定样式档案）；gj 导出的 JSON 不携带原图片二进制，原表格以只读文本条目（`exists: true`）保留，重建时不会重复建表。
- `af --sync`：原表格判定为只读（不检查、不改写、不重插）；JSON 中新增的表格条目会被忽略（需先 `jd`）；含图片段落不可移动；大段重复空文本对齐可能漂，漂了逐段人工确认。
- 删文献条目时需同步改正文条目里的引文编号；删标题/题注条目则编号按新树序自动重算（两种模式都是）。

## 预设（`presets/`）

学校/学院的定制规则不写进 YAML，而是放 `presets/` 下的 `.py` 插件，用 `--preset 名` 加载：

- 可选 `PRESET_MANIFEST`：声明预设需要的 YAML 兜底配置，加载时深合并（用户 `-c` 配置优先）。
- 必须实现 `register(hooks)`：用 `hooks.register(事件, 回调)` 挂回调；回调签名 `callback(event_data) -> dict | None`，可直接改 `event_data` 里的可变对象（`config`/`root_node`/`issues` 等），或返回字典覆盖字段（如 `{"text": ...}` 改写批注文案）。
- 可用事件：`on_config_loaded`、`on_tree_built`、`on_tree_normalize`、`on_format_begin`、`before_node_format`、`after_node_format`、`on_caption_check`、`on_heading_numbering`、`on_summary_build`、`before_document_save`、`on_comment`。
- 想替换框架内置行为，先 `hooks.unregister(事件, "builtin.*"/"thesis.*")` 再自行实现；写法见 `presets/_template.py`。
- 多个 `--preset` 叠加时后加载者优先；`wordf list-presets` 可查看已安装预设（`_template.py` 不列出）。

## JSON 字段

| 字段 | 说明 |
|------|------|
| `category` | 段落类型，改这里修正分类 |
| `paragraph` | 段落文本（`jd`/`af --sync` 模式下直接编辑此字段即可改/增/删内容） |
| `replace` | 替换段落内容；空字符串表示清空。若同时残留 `inline_marks` 且文本不一致，以 `replace` 为准 |
| `score` | AI 置信度 |
| `table_rows` | 表格内容（二维字符串列表，`table_object` 条目；gj 导出为只读） |
| `exists` | `true` 表示对应原文档真实表格：不重复建表、不参与段落对齐 |
| `table_style` | 指定表格样式档案名（覆盖 YAML `tables.default`） |
| `inline_marks` | 行内格式片段（`{text, bold, italic, code, strikethrough, link_url}`），`jd`/`md` 据此生成多 run |

## 格式检查范围

段落格式（对齐/间距/行距/缩进）、字符格式（字体/字号/颜色/粗斜体/下划线）、标题自动编号、题注编号、关键词数量、标点符号、正文段落字数/句数阈值（可配，默认不限）、参考文献著录（GB/T 7714-2015 全部 12 种文献类型：类型标识/期刊卷期页/出版项/报纸日期版次/标准编号/专利号/联机引用日期与路径，只报批注）、图片版式（环绕方式/超出版心，只报批注）、页眉页脚（字号/字体/对齐/内容缺失/页码域）、页面设置（页边距、纸张尺寸与方向、起始页码与页码格式、首页不同、奇偶页不同、打开时更新域）。

表格：gj 会把原文档表格导出为只读 `table_object` 条目（`table_rows` + `exists: true`），供 agent 阅读与统计；表格内容本身不参与格式检查与同步（只读）。`jd`/`md` 新建表格支持任意样式档案（见下）。

行内富文本：`md`/`jd` 按 `inline_marks` 生成多 run（粗体/斜体/删除线/代码/链接），后续字符格式化不会把这三项强调拉平（YAML 显式声明 `bold`/`italic`/`underline` 时以配置为准）。

页眉页脚/页边距通过 YAML 配置 `header_footer` 段控制（`af`/`cf` 时按节确定性处理，不依赖结构识别）：

```yaml
header_footer:
  enabled: true
  update_fields_on_open: false   # 打开文档时自动更新域（settings.xml w:updateFields）
  header: { font_size: '小五', alignment: '居中对齐', require_content: false, page_number: false }
  footer: { ... }
  page_setup:
    top_margin: '2.54cm'
    bottom_margin: '2.54cm'
    left_margin: '3.18cm'
    right_margin: '3.18cm'
    header_distance: '1.5cm'
    footer_distance: '1.75cm'
    page_size: 'A4'              # A3 / A4 / A5 / B5 / Letter
    orientation: '纵向'          # 纵向 / 横向
    page_number_start: 1
    page_number_format: '小写罗马数字'
    different_first_page: false  # 首页不同（w:titlePg）
    odd_even_pages: false        # 奇偶页不同
```

`header.page_number` / `footer.page_number: true` 要求存在 PAGE 页码域：check 只报缺失，apply 会插入 `w:fldSimple w:instr=" PAGE "`（域结果需 Word 刷新才显示真实页码，可配 `update_fields_on_open: true`）。

### 表格样式（`tables`，镜像 OOXML tblPr/trPr/tcPr，任意组合；缺省 = Table Grid）

```yaml
tables:
  default: grid                  # 缺省样式档案名
  styles:
    grid:
      style: "Table Grid"        # Word 表格样式名；null = 不套样式、全由 XML 定义
      layout: autofit            # autofit|fixed
      width: {type: auto}        # auto|pct|dxa|cm
      alignment: center          # left|center|right
      indent_cm: 0
      borders:                   # 6 边独立；style 取 OOXML val（none/single/double/dashed/...）
        top: {style: single, size_pt: 1.5, color: "000000"}
        left: {style: none}
        right: {style: none}
        inside_h: {style: single, size_pt: 0.75}
        inside_v: {style: none}
      cell_margins_cm: {top: 0, bottom: 0, left: 0.19, right: 0.19}
      col_widths: {mode: equal}  # equal|auto|custom（custom: [4, 6, 6] cm）
      header_row: {repeat: true, cant_split: true}
      row_bands: [{rows: [0], fill: "EEEEEE", bold: true}]
      cell_run: {chinese_font_name: 宋体, english_font_name: Times New Roman, font_size: 小四}
```

上例即三线表；任意边框组合、底纹、列宽均由此模式表达，条目内 `table_style` 可逐表指派。

### 行内富文本（`inline`）

```yaml
inline:
  code_font: {chinese_font_name: 宋体, english_font_name: Consolas, font_size: 五号}
  link: {style: Hyperlink, color: theme, underline: true}   # color: theme|<RRGGBB>
```

### 结构类别策略（`structure`，按类别名单驱动）

```yaml
structure:
  void_categories: [top, heading_mulu, heading_fulu, other, document_title, footer, toc_entry]  # 不参与格式化
  skip_children_categories: [heading_mulu, heading_fulu, other]  # 子树不遍历
```

默认值如上；想让附录参与格式化（收批注/被修正）= 从这两个列表删 `heading_fulu`，其他类别同理。
名单是权威的：显式写空列表（`void_categories: []` / `skip_children_categories: []`）
表示不排除任何类别、全部参与格式化与遍历，不会回退到默认值。
唯一例外：直接挂在 root 下的 `body_text`（无标题归属的封面/声明类内容）始终被隐式跳过。

## 不支持

目录生成、封面排版；原文档表格内容级检查与改写（只读导出）。