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
```

验证：`wf --help`

## 工作流程

本技能包含**两个独立任务**，可分别执行，互不干扰：

| 任务 | 说明 | 产物 |
|------|------|------|
| **任务一：准备配置文件** | 根据格式要求生成/编辑 config.yaml | `config.yaml` |
| **任务二：执行格式化** | 使用配置文件对论文进行格式检查或修正 | `--标注版.docx` 或 `--修改版.docx` |

**执行顺序**：必须先完成任务一再执行任务二。两个任务可以分两次对话完成。

---

## 任务一：准备配置文件

> **本任务的产物是 `config.yaml`，必须保存到用户工作目录。**

严格按照以下步骤执行。**每个步骤开始前，先阅读该步骤引用的参考文档，不要提前阅读后续步骤的文档。**

### 步骤 1.0 查找已有预设

预设保存在**项目工作目录**下 `presets/` 目录，命名格式 `{学校}_{学院/专业}_{论文类型}.yaml`。

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --list-presets
```

找到匹配预设时直接使用，跳到 1.3：
```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --use "清华大学_计算机学院_本科" --output config.yaml
```

没有匹配预设时，继续 1.1。

### 步骤 1.1 提取格式要求并编辑配置

> **开始本步骤前，阅读 [data/config_editing_guide.md](data/config_editing_guide.md) 了解字段取值。如需查看完整字段白名单，阅读 [data/config_spec.md](data/config_spec.md)。**

**⚠️ 关键：必须先完整读取用户的格式规范文件，提取所有格式参数，再编辑 config.yaml。不要边看边改。**

需要提取的参数：正文字体/字号/行距/段前段后/缩进、各级标题格式、摘要格式、关键词数量、图表题注、参考文献、致谢。

**如果格式文档要求标题自动编号（如"第一章"、"1.1"），还需提取：**
- 各级标题的编号格式（如 `第%1章`、`%1.%2`、`%1.%2.%3`）
- 是否需要清除手动编号（用户手动打的编号文字）
- 编号后缀（制表符 `tab`、空格 `space`、无 `nothing`）
- 编号缩进和文本缩进（如 `"0字符"`、`"0.75cm"`）

**特别注意：`line_spacingrule` 和 `line_spacing` 必须配合设置：**

| line_spacingrule | line_spacing | 说明 |
|------------------|--------------|------|
| `固定值` | `"20磅"` | 固定 20 磅行距 |
| `1.5倍行距` | `"1.5倍"` | 1.5 倍行距 |
| `单倍行距` | `"1倍"` | 单倍行距 |
| `最小值` | `"12磅"` | 最小 12 磅 |

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --create --output config.yaml
```

编辑核心原则：
- **只修改已有字段的值，不要添加新字段**
- **不要删除任何已有字段**
- **不要修改 YAML 锚点语法**（`&global_format` 和 `<<: *global_format`）

### 步骤 1.2 验证配置文件

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --validate --config config.yaml
```

验证失败则根据提示修正，直到通过。

### 步骤 1.3 保存到预设库（首次生成时）

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --save --config config.yaml --name "XX大学_XX学院_本科"
```

### 步骤 1.4 交付配置文件

**⚠️ 必须将 `config.yaml` 复制到用户工作目录（workspace），作为任务产物交付给用户。**

```bash
cp config.yaml <用户工作目录>/config.yaml
```

然后**询问用户**：配置文件已生成，请检查各项格式参数是否正确。有些格式要求可能无法通过配置文件完全体现（如特殊的页眉页脚、目录格式等），用户可以手动补充或告知需要调整的地方。

**任务一完成。** 用户确认配置无误后，可继续执行任务二。

---

## 任务二：执行格式化

> **前置条件**：需要任务一产出的 `config.yaml` 和原始 `.docx` 论文文件。**

### 步骤 2.1 生成文档结构 JSON

```bash
wf gj -d $ARGUMENTS -c config.yaml
```

记住输出的 JSON 文件路径。

### 步骤 2.2 查看文档结构树

```bash
wf tree -f <JSON文件路径>
```

输出文档的段落分类统计和树形结构图，快速检查分类是否正确。

```bash
# 仅查看标题结构
wf tree -f <JSON文件路径> --filter heading_level_1,heading_level_2

# 显示分类置信度（低置信度的段落需要重点检查）
wf tree -f <JSON文件路径> --confidence
```

### 步骤 2.3 校验并检查 JSON 标签

> **开始本步骤前，阅读 [data/category_reference.md](data/category_reference.md) 了解 category 判定规则、自动类型提升机制、以及 `"other"` 标记规则。**

#### 运行校验脚本

```bash
python ${CLAUDE_SKILL_DIR}/scripts/validate_json.py --json <JSON文件路径> --stats --show-all --threshold 0.8
```

#### 人工检查并修正

对照文档实际内容，检查每个段落的分类是否正确。常见误判：
- "摘要" → `abstract_chinese_title`，不是 `heading_level_1`
- "关键词：..." → `keywords_chinese`，不是 `body_text`
- "参考文献" → `references_title`，不是 `heading_level_1`
- 摘要标题+正文在同一段 → `abstract_chinese_title_content`

**⚠️ 目录、页眉页脚等非内容段落 → 改为 `"other"`**（代码会自动跳过）

**⚠️ 目录部分必须全部标记为 `"other"`，不要格式化、不要添加段落标签：**
- 目录标题（"目录"、"目 录"）→ `"other"`
- 所有目录条目（带省略号和页码的行）→ `"other"`
- 目录区域内的所有段落（含空行）→ `"other"`
- 目录条目常被误识别为 `heading_level_1`、`heading_level_2`、`body_text`，必须全部纠正

**⚠️ 不要修改以下 `body_text`（代码会自动处理）：**
- 摘要标题后面的正文段落 → 自动升级为摘要正文类型
- 参考文献标题后面的每条文献 → 自动升级为参考文献条目类型

发现错误时直接编辑 JSON 文件修改 `category` 字段，修正后重新运行校验脚本。

### 步骤 2.4 执行格式检查或格式化

**检查格式（不修改原文档）：**
```bash
wf cf -d 论文.docx -c config.yaml -f <JSON文件路径>
```

**自动格式化（直接修正）：**
```bash
wf af -d 论文.docx -c config.yaml -f <JSON文件路径>
```

> **注意**：如果 config.yaml 中启用了 `numbering.enabled: true`，格式化时会自动：
> 1. 用正则清除标题段落的手动编号文字
> 2. 应用 Word 自动编号（编号由 Word 渲染，不会丢失）
> 此功能仅在格式化模式（`wf af`）下生效，检查模式（`wf cf`）不会修改编号。

### 步骤 2.5 交付产物

**⚠️ 必须将输出文件复制到用户工作目录（workspace），作为任务产物交付给用户。**

输出文件说明：
- 检查模式：`--标注版.docx`（原文档副本，批注标注格式问题）
- 格式化模式：`--修改版.docx`（直接修正后的文档）

```bash
cp <输出文件> <用户工作目录>/
```

然后**询问用户**：格式化已完成，请打开文档检查格式是否正确。由于部分格式（如页眉页脚、目录、特殊排版等）无法通过工具自动处理，可能需要手动补充调整。如有问题请告知需要修改的地方。

---

## Python 编程调用

```python
import os
os.environ['MULTIPROCESSING_RESOURCE_TRACKER'] = '0'
os.environ['OMP_NUM_THREADS'] = '1'
os.environ['MKL_NUM_THREADS'] = '1'

from wordformat.set_tag import set_tag_main
from wordformat.set_style import auto_format_thesis_document
import json

data = set_tag_main(docx_path="论文.docx", configpath="config.yaml")
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=4)

auto_format_thesis_document(
    jsonpath="output.json", docxpath="论文.docx",
    configpath="config.yaml", savepath="output/", check=True
)
```

## 注意事项

- 仅支持 `.docx` 格式，不支持 `.doc`
- Python >= 3.10
- **务必检查 JSON 中的分类结果**，AI 识别并非 100% 准确
- **配置文件和输出文件都是任务产物，必须保存到用户工作目录**
- **每个任务完成后都必须询问用户确认结果**
- 部分格式（页眉页脚、目录、特殊排版等）无法自动处理，需提醒用户手动补充
