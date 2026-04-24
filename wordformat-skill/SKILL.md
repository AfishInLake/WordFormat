---
name: wordformat
description: 论文格式自动化处理工具。在处理 Word 论文文档格式校验、格式修正、文档结构识别场景时激活，具备使用 AI 模型智能识别文档结构并根据 YAML 配置文件自动校验或修正论文格式的专业能力。
argument-hint: "[docx文件路径]"
---

# WordFormat - 论文格式自动化处理工具

使用 AI 模型智能识别 Word 文档结构，根据 YAML 配置文件自动校验或修正论文格式。

## 安装

```bash
pip install wordformat
# 或
uv pip install wordformat
```

验证：`wf --help`

## 完整工作流程

严格按照以下步骤执行，不可跳过。

### 第一步：准备配置文件

#### 1.0 查找已有预设

预设配置保存在**项目工作目录**下的 `presets/` 目录中，命名格式为 `{学校}_{学院/专业}_{论文类型}.yaml`。

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --list-presets
```

如果用户指定了学校/专业，查找匹配的预设。**找到匹配预设时直接使用，跳到 1.3**：
```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --use "清华大学_计算机学院_本科" --output config.yaml
```

没有匹配预设时，继续 1.1。

#### 1.1 从格式文档提取要求并生成配置

**完整读取用户提供的格式规范文件，逐一提取所有格式要求。** 以下是完整的提取对照表，每个配置节点下可覆盖的字段均已列出。

> 所有子节点都继承 `global_format` 的 15 个基础字段（对齐方式、间距、行距、缩进、字体、字号、颜色、加粗/斜体/下划线、内置样式名），只需覆盖与全局不同的字段。

**① global_format（全局基础格式，影响所有段落）**

| 格式文档中常见的描述 | config.yaml 字段 | 常见值 |
|---------------------|-----------------|--------|
| 正文对齐方式 | `alignment` | `两端对齐`、`左对齐`、`居中对齐` |
| 段前间距 | `space_before` | `"0行"` `"0.5行"` `"12磅"` |
| 段后间距 | `space_after` | `"0行"` `"0.5行"` `"12磅"` |
| 行距类型 | `line_spacingrule` | `单倍行距` `1.5倍行距` `固定值` `多倍行距` |
| 行距值 | `line_spacing` | `"1.5倍"` `"20磅"` `"1.25倍"` |
| 左缩进 | `left_indent` | `"0字符"` `"2字符"` |
| 右缩进 | `right_indent` | `"0字符"` |
| 首行缩进 | `first_line_indent` | `"2字符"` `"0字符"` |
| 中文字体 | `chinese_font_name` | `宋体` `黑体` `楷体` `仿宋` `微软雅黑` `汉仪小标宋` |
| 英文字体 | `english_font_name` | `Times New Roman` `Arial` `Calibri` `Courier New` `Helvetica` |
| 字号 | `font_size` | `小四`(12pt) `三号`(16pt) `五号`(10.5pt) 等，或数值如 `12` |
| 字体颜色 | `font_color` | `黑色` `红色` 或 `#FF0000` |
| 加粗 | `bold` | `true` / `false` |
| 斜体 | `italic` | `true` / `false` |
| 下划线 | `underline` | `true` / `false` |
| Word 内置样式名 | `builtin_style_name` | `正文`（正文默认） |

**② abstract（摘要）— 继承 global_format 15 字段**

| 格式文档中常见的描述 | config.yaml 字段路径 | 说明 |
|---------------------|---------------------|------|
| 中文摘要标题字体/字号/对齐/加粗 | `abstract.chinese.chinese_title` | 覆盖继承的字段 |
| 中文摘要正文字体/字号/行距 | `abstract.chinese.chinese_content` | 覆盖继承的字段 |
| 英文摘要标题字体/字号/对齐/加粗 | `abstract.english.english_title` | 覆盖继承的字段 |
| 英文摘要正文字体/字号/行距 | `abstract.english.english_content` | 覆盖继承的字段 |
| 关键词是否加粗 | `abstract.keywords.chinese.keywords_bold` | `true` / `false` |
| 关键词最少数量 | `abstract.keywords.chinese.count_min` | 正整数，如 `3` `4` |
| 关键词最多数量 | `abstract.keywords.chinese.count_max` | 正整数，如 `5` `6` |
| 禁止关键词末尾标点 | `abstract.keywords.chinese.trailing_punct_forbidden` | `true` / `false` |
| 英文关键词（同上 4 个字段） | `abstract.keywords.english.*` | 与中文关键词字段相同 |

**③ headings（各级标题）— 继承 global_format 15 字段**

| 格式文档中常见的描述 | config.yaml 字段路径 | 说明 |
|---------------------|---------------------|------|
| 一级标题（章标题）全部格式 | `headings.level_1` | 字体、字号、对齐、加粗、间距等 |
| 二级标题（节标题）全部格式 | `headings.level_2` | 同上 |
| 三级标题（小节标题）全部格式 | `headings.level_3` | 同上 |

> **注意**：各级标题的 `builtin_style_name` 必须对应：`level_1` → `Heading 1`、`level_2` → `Heading 2`、`level_3` → `Heading 3`

**④ body_text（正文）— 继承 global_format 15 字段**

通常不需要单独设置，直接继承 `global_format` 即可。

**⑤ figures（插图及图注）— 继承 15 字段 + 2 个专用字段**

| 格式文档中常见的描述 | config.yaml 字段 | 常见值 |
|---------------------|-----------------|--------|
| 图注位置（图上方/下方） | `caption_position` | `above` / `below` |
| 图注编号前缀 | `caption_prefix` | `图` `Figure` `Fig.` |
| 图注字号/对齐/字体等 | （覆盖继承字段） | 如 `font_size: '五号'` `alignment: '居中对齐'` |

**⑥ tables（表格及表注）— 继承 15 字段 + 2 个专用字段**

| 格式文档中常见的描述 | config.yaml 字段 | 常见值 |
|---------------------|-----------------|--------|
| 表注位置（表上方/下方） | `caption_position` | `above` / `below` |
| 表注编号前缀 | `caption_prefix` | `表` `Table` `Tab.` |
| 表注字号/对齐/字体等 | （覆盖继承字段） | 如 `font_size: '五号'` `alignment: '居中对齐'` |

**⑦ references（参考文献）— 继承 15 字段 + 专用字段**

| 格式文档中常见的描述 | config.yaml 字段路径 | 常见值 |
|---------------------|---------------------|--------|
| 参考文献标题字体/字号/对齐/加粗 | `references.title` | 覆盖继承字段 |
| 参考文献标题文字 | `references.title.section_title` | `参考文献` |
| 参考文献条目字体/字号 | `references.content` | 覆盖继承字段 |
| 参考文献编号格式 | `references.content.numbering_format` | `'[1], [2], ...'` 或 `'1.'` |
| 参考文献条目缩进 | `references.content.entry_indent` | `0.0`（顶格） |
| 参考文献条目结尾标点 | `references.content.entry_ending_punct` | `'.'` `'。'` 或 `null`（不限制） |

**⑧ acknowledgements（致谢）— 继承 15 字段，无专用字段**

| 格式文档中常见的描述 | config.yaml 字段路径 | 说明 |
|---------------------|---------------------|------|
| 致谢标题字体/字号/对齐/加粗 | `acknowledgements.title` | 覆盖继承字段 |
| 致谢正文字体/字号/行距 | `acknowledgements.content` | 覆盖继承字段 |

**⑨ style_checks_warning（格式警告开关）**

控制格式校验时哪些属性不满足规范时触发警告，通常不需要修改。

参考文档：[data/config_editing_guide.md](data/config_editing_guide.md)、[data/font_size_table.md](data/font_size_table.md)、[data/config_spec.md](data/config_spec.md)

#### 1.2 创建并编辑配置文件

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --create --output config.yaml
```

将 1.1 提取的所有格式要求填写到 config.yaml。核心原则：
- **只修改已有字段的值，不要添加新字段**
- **不要删除任何已有字段**
- **不要修改 YAML 锚点语法**（`&global_format` 和 `<<: *global_format`）

#### 1.3 验证配置文件

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --validate --config config.yaml
```

验证失败则根据提示修正，直到通过。

#### 1.4 保存到预设库（首次生成时）

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --save --config config.yaml --name "XX大学_XX学院_本科"
```

命名格式：`学校_学院或专业_论文类型`，论文类型取值：本科/硕士/博士/期刊。

### 第二步：生成文档结构 JSON

```bash
wf gj -d $ARGUMENTS -c config.yaml
```

记住终端输出的 JSON 文件路径。

### 第三步：校验并检查 JSON 标签（关键步骤）

#### 3.1 运行校验脚本

```bash
python ${CLAUDE_SKILL_DIR}/scripts/validate_json.py --json output/论文_1234567890.json --stats --show-all --threshold 0.8
```

脚本会：
- 检查所有 category 是否为合法值（共 15 种）
- 标记低置信度分类（低于阈值）
- 输出分类统计
- 逐条显示分类结果供人工检查

#### 3.2 人工检查并修正

**对照文档实际内容，检查每个段落的分类是否正确。** 重点注意：
- "摘要" → 应为 `abstract_chinese_title`，不是 `heading_level_1`
- "关键词：..." → 应为 `keywords_chinese`，不是 `body_text`
- "参考文献" → 应为 `references_title`，不是 `heading_level_1`
- "致谢" → 应为 `acknowledgements_title`，不是 `heading_level_1`
- 摘要标题+正文在同一段 → 应为 `abstract_chinese_title_content`

**⚠️ 不要修改以下 `body_text`（代码会自动处理）：**
- 摘要标题后面的正文段落 → 代码自动升级为摘要正文类型
- 参考文献标题后面的每条文献 → 代码自动升级为参考文献条目类型
- 详见 [data/category_reference.md](data/category_reference.md) 的"自动类型提升机制"章节

**完整的 category 判定规则**：参见 [data/category_reference.md](data/category_reference.md)

发现错误时，直接编辑 JSON 文件修改 `category` 字段。修正后重新运行校验脚本确认。

### 第四步：执行格式检查或格式化

**检查格式（不修改原文档）：**
```bash
wf cf -d 论文.docx -c config.yaml -f output/论文_1234567890.json
```
生成 `论文--标注版.docx`（带批注）。

**自动格式化（直接修正）：**
```bash
wf af -d 论文.docx -c config.yaml -f output/论文_1234567890.json
```
生成 `论文--修改版.docx`（格式已修正）。

## Python 编程调用

**重要：在调用 wordformat 之前，必须先设置环境变量禁用多进程，否则在沙箱环境中会崩溃。**

```python
import os
# 必须在 import wordformat 之前设置
os.environ['MULTIPROCESSING_RESOURCE_TRACKER'] = '0'
os.environ['OMP_NUM_THREADS'] = '1'
os.environ['MKL_NUM_THREADS'] = '1'

from wordformat.set_tag import set_tag_main
from wordformat.set_style import auto_format_thesis_document
import json

data = set_tag_main(docx_path="论文.docx", configpath="config.yaml")
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)

# 检查并修正 JSON 中的 category 字段后：

auto_format_thesis_document(
    jsonpath="output.json", docxpath="论文.docx",
    configpath="config.yaml", savepath="output/", check=True  # True=标注, False=修正
)
```

## 注意事项

- 仅支持 `.docx` 格式，不支持 `.doc`
- Python 版本要求 >= 3.10
- **务必检查 JSON 文件中的分类结果**，AI 识别并非 100% 准确
- 输出文件命名：检查模式为 `--标注版.docx`，格式化模式为 `--修改版.docx`
