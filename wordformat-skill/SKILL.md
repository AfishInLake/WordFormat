---
name: wordformat
description: 论文格式自动化处理工具。在处理 Word 论文文档格式校验、格式修正、文档结构识别场景时激活，具备使用 AI 模型智能识别文档结构并根据 YAML 配置文件自动校验或修正论文格式的专业能力。
argument-hint: "[docx文件路径]"
---

# WordFormat - 论文格式自动化处理工具

## 安装

```bash
pip install wordformat -i https://pypi.tuna.tsinghua.edu.cn/simple  # 国内
pip install wordformat  # 海外
```

验证：`wordf --help`

## 功能速查

### CLI 命令

| 命令 | 功能 | 输入 | 输出 |
|------|------|------|------|
| `wordf config` | 查看所有可配置字段及默认值 | 无 | 终端输出 |
| `wordf config -o config.yaml` | 输出完整配置模板 | 无 | `config.yaml` |
| `wordf gj` | AI 识别文档结构 | `.docx` + `config.yaml` | `output/xxx.json` |
| `wordf tree` | 查看结构树 | `.json` | 终端树形图 |
| `wordf cf` | 检查格式（批注） | `.docx` + `config.yaml` + `.json` | `--标注版.docx` |
| `wordf af` | 修正格式（直接改） | `.docx` + `config.yaml` + `.json` | `--修改版.docx` |
| `wordf startapi` | Web 界面 | 无 | http://127.0.0.1:8000 |

### 辅助脚本

| 脚本 | 功能 |
|------|------|
| `setup_config.py` | 验证/保存 YAML 配置 |
| `validate_json.py` | 校验 JSON category，显示统计 |

### JSON 字段（`wordf gj` 输出，可手动编辑）

| 字段 | 说明 |
|------|------|
| `category` | 段落类型（16种），改这里修正分类 |
| `replace` | 可选，填入新文本直接替换段落内容 |
| `fingerprint` | 段落指纹（不要改） |
| `score` | AI 置信度 |

### 格式检查范围

段落（对齐/间距/行距/缩进）、字符（字体/字号/颜色/粗斜体/下划线）、标题自动编号、题注编号校验、关键词数量、表格内容。

### 不支持

页眉页脚、目录生成、封面排版、图片尺寸、Word 域代码。需提前告知用户手动处理。

---

## 工作流程

根据用户意图走不同路径：

- **仅文本替换**：只想修正论文中某些段落的文字内容（如错别字、措辞），不检查格式 → [纯文本替换流程](#纯文本替换流程)
- **格式校验/修正**：检查或修正论文格式 → 下面两个独立任务，先任务一再任务二

每个任务完成后**必须将产物复制到用户工作目录**。

---

## 纯文本替换流程

> 用户只想换文字（错别字、措辞修正），不需要格式批注/编号/样式修改。

1. 生成 JSON：`wordf gj -d 论文.docx -c config.yaml`
2. 编辑 JSON：在需要替换的段落对象中添加 `"replace": "新的正确文本"`
3. 执行替换：`wordf af -d 论文.docx -c config.yaml -f output/论文_xxx.json`
4. 交付：`cp output/论文--修改版.docx <用户工作目录>/`

> `wordf af` 执行替换的同时也会应用格式修正。如果 config.yaml 格式规范与论文当前格式一致，则仅文本被替换、格式不变。**用户说"只换文字"时走此流程，不要追加 `cf` 检查或格式批注操作。**

---

## 格式校验/修正流程

---

### 任务一：准备配置文件

> **🚨 用户已提供 .yaml 配置 → 直接跳到步骤 1.2 验证，禁止新建。**
> **🚨 用户说"用上次的配置"/"和上次一样" → 复用已有配置，禁止新建。**

### 步骤 1.0 查找已有配置

**优先级：用户提供的 > 工作目录已有 > 预设库 > 新建**

```bash
# 扫描工作目录下所有 .yaml 文件
ls *.yaml 2>/dev/null || echo "无 .yaml 文件"
```

找到文件后逐个验证：

```bash
for f in *.yaml; do
  echo "--- 验证: $f ---"
  python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --validate --config "$f" 2>&1 \
    && echo "✅ 合法" || echo "❌ 非配置文件"
done
```

找到合法配置 → **询问用户确认**，不要新建或覆盖。用户确认后跳到步骤 1.2。

没有 .yaml 文件时尝试预设库：

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --list-presets
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --use "清华大学_计算机学院_本科" --output config.yaml
```

预设也找不到时，继续步骤 1.1 新建。

### 步骤 1.1 新建配置

> 运行 `wordf config` 查看所有字段和默认值，`wordf config -o config.yaml` 生成模板。

**先完整读完格式要求文件，再编辑。** 编辑原则：只改值、不增减字段、不动 YAML 锚点语法。

### 步骤 1.2 验证

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --validate --config config.yaml
```

失败则根据提示修正，直到通过。

### 步骤 1.3 保存预设（首次生成时可选）

```bash
python ${CLAUDE_SKILL_DIR}/scripts/setup_config.py --save --config config.yaml --name "XX大学_XX学院_本科"
```

### 步骤 1.4 交付

```bash
cp config.yaml <用户工作目录>/
```

询问用户确认配置，提醒页眉页脚、目录等无法自动处理。**任务一完成。**

---

### 任务二：执行格式化

> 前提：有 `config.yaml` 和 `.docx` 论文。

### 步骤 2.1 生成 JSON

```bash
wordf gj -d $ARGUMENTS -c config.yaml
```

记住输出的 JSON 路径（`output/论文_xxx.json`）。

### 步骤 2.2 检查结构

```bash
wordf tree -f <JSON路径>
wordf tree -f <JSON路径> --confidence  # 看低置信度段落
```

### 步骤 2.3 校验并修正 JSON

> 先读 [data/category_reference.md](data/category_reference.md)。

```bash
python ${CLAUDE_SKILL_DIR}/scripts/validate_json.py --json <JSON路径> --stats --show-all --threshold 0.8
```

**常见修正**：

| 误识别 | 改为 |
|--------|------|
| "摘要" → `heading_level_1` | `abstract_chinese_title` |
| "关键词：..." → `body_text` | `keywords_chinese` |
| "参考文献" → `heading_level_1` | `references_title` |
| 摘要标题+正文同一段 | `abstract_chinese_title_content` |
| 目录/封面/页眉页脚 | `other`（跳过格式化） |

**⚠️ 不修改的 body_text**：摘要标题后的正文（自动升级为摘要正文）、参考文献标题后的条目（自动升级为参考文献条目）。

**用 `replace` 修正文本**：添加 `"replace": "正确文本"`，格式化时自动替换。

修正后重新验证确认无误。

### 步骤 2.4 格式化

```bash
wordf cf -d 论文.docx -c config.yaml -f <JSON路径>  # 检查（批注，不改原文）
wordf af -d 论文.docx -c config.yaml -f <JSON路径>  # 修正（直接改）
```

`numbering.enabled: true` 时，`af` 模式会自动清除手动编号并应用 Word 自动编号。

### 步骤 2.5 交付

```bash
cp output/论文--标注版.docx <用户工作目录>/   # cf 模式
cp output/论文--修改版.docx <用户工作目录>/   # af 模式
```

询问用户检查结果，提醒手动补充。**任务二完成。**

---

## Python 调用

```python
from wordformat.set_tag import set_tag_main
from wordformat.set_style import auto_format_thesis_document

data = set_tag_main(docx_path="论文.docx", configpath="config.yaml")
auto_format_thesis_document(jsonpath=data, docxpath="论文.docx",
                            configpath="config.yaml", savepath="output/", check=True)
```

## 注意事项

- 仅 `.docx`，Python ≥ 3.10
- AI 分类非 100% 准确，**务必检查 JSON**
- 产物必须复制到用户工作目录
- 每个任务完成后询问用户确认
