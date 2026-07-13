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

## 命令速查

| 命令 | 功能 |
|------|------|
| `wordf md -d 论文.md -c config.yaml` | Markdown → 格式化 .docx |
| `wordf gj -d 论文.docx -c config.yaml` | AI 识别文档结构，输出 JSON |
| `wordf tree -f output/xxx.json` | 查看文档结构树 |
| `wordf cf -d 论文.docx -c config.yaml -f output/xxx.json` | 检查格式（加批注） |
| `wordf af -d 论文.docx -c config.yaml -f output/xxx.json` | 修正格式（直接改） |
| `wordf config -o config.yaml` | 输出配置模板 |
| `wordf startapi` | 启动 Web 界面 |

## 工作流程

根据用户意图选择路径：

### 纯文本替换

只换文字不查格式。生成 JSON → 编辑 `replace` 字段 → 执行替换。

```bash
wordf gj -d 论文.docx -c config.yaml
# 编辑 output/xxx.json，添加 "replace": "正确文本"
wordf af -d 论文.docx -c config.yaml -f output/xxx.json
```

### Markdown 转 Word

```bash
wordf md -d 论文.md -c config.yaml
# 输出 output/论文--生成版.docx
```

### 格式校验/修正

```bash
# 1. 准备配置（已有则跳过）
wordf config -o config.yaml
# 编辑 config.yaml 适配学校要求

# 2. 生成结构
wordf gj -d 论文.docx -c config.yaml

# 3. 检查结构（可选）
wordf tree -f output/xxx.json --confidence

# 4. 格式化
wordf cf -d 论文.docx -c config.yaml -f output/xxx.json  # 检查
wordf af -d 论文.docx -c config.yaml -f output/xxx.json  # 修正
```

## JSON 字段

| 字段 | 说明 |
|------|------|
| `category` | 段落类型，改这里修正分类 |
| `replace` | 填入新文本替换段落内容 |
| `score` | AI 置信度 |

## 格式检查范围

段落格式（对齐/间距/行距/缩进）、字符格式（字体/字号/颜色/粗斜体/下划线）、标题自动编号、题注编号、关键词数量、标点符号。

## 不支持

页眉页脚、目录生成、封面排版。需提前告知用户手动处理。
