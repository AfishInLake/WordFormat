# 使用指南（已更新至 极简命令版）

本文档详细说明 WordFormat 的使用方法，包括命令行交互、Python 编程调用和 API 调用三种方式。

## 命令行使用（推荐 · 极简版）

WordFormat 现已支持**超短命令**，输入更快捷、不易出错。
支持 **`wordformat` / `wf`** 双命令启动。

---

## 极简命令速查（最常用）
```bash
# 生成文档结构 JSON（自动保存，无需指定json路径）
wf gj -d 论文.docx -c 配置.yaml

# 检查格式错误（添加批注，不修改原文）
wf cf -d 论文.docx -c 配置.yaml -f 生成的文件.json

# 自动格式化论文（一键修正格式）
wf af -d 论文.docx -c 配置.yaml -f 生成的文件.json
```

---

## 详细使用说明

### 命令说明
- `wf` 或 `wordformat`：工具主命令
- `gj`：generate-json → 生成文档结构 JSON
- `cf`：check-format → 检查格式
- `af`：apply-format → 自动格式化

### 通用参数
- `-d`：**必填**，Word 文档路径
- `-c`：**必填**，YAML 格式配置文件路径
- `-f`：JSON 文件路径（**仅 cf/af 需要**）
- `-o`：输出目录（可选，默认 `output/`）

---

## 1. 生成文档结构 JSON
**自动生成 JSON 文件，保存到 `-o` 目录**
**文件名 = 文档名 + 10位时间戳**，永不重复

```bash
# 最简用法
wf gj -d your_document.docx -c example/undergrad_thesis.yaml

# 自定义输出目录
wf gj -d your_document.docx -c example/undergrad_thesis.yaml -o output/
```

---

## 2. 执行格式校验
仅检查错误、添加 Word 批注，**不修改原文**

```bash
# 基础用法
wf cf -d your_document.docx -c example/undergrad_thesis.yaml -f output/论文_1744123456.json

# 自定义输出目录
wf cf -d your_document.docx -c example/undergrad_thesis.yaml -f output/论文_1744123456.json -o check_result/
```

---

## 3. 执行自动格式化
一键自动修正论文格式，生成新的规范文档

```bash
# 基础用法
wf af -d your_document.docx -c example/undergrad_thesis.yaml -f output/论文_1744123456.json

# 自定义输出目录
wf af -d your_document.docx -c example/undergrad_thesis.yaml -f output/论文_1744123456.json -o final_output/
```

---

## 完整测试示例
```bash
# 1. 生成 JSON（自动命名）
wf gj -d "tmp/毕业设计说明书.docx" -c "example/undergrad_thesis.yaml"

# 2. 格式检查
wf cf -d "tmp/毕业设计说明书.docx" -c "example/undergrad_thesis.yaml" -f "output/毕业设计说明书_1744123456.json"

# 3. 自动格式化
wf af -d "tmp/毕业设计说明书.docx" -c "example/undergrad_thesis.yaml" -f "output/毕业设计说明书_1744123456.json"
```

---

## 参数对照表（极简版）

| 命令 | 全称 | 作用 | 必填参数 |
|------|------|------|----------|
| `wf gj` | generate-json | 生成文档结构 JSON | `-d`,`-c` |
| `wf cf` | check-format | 检查格式并添加批注 | `-d`,`-c`,`-f` |
| `wf af` | apply-format | 自动格式化论文 | `-d`,`-c`,`-f` |

| 参数 | 作用 | 必填 |
|------|------|------|
| `-d` | Word 文档路径 | 是 |
| `-c` | YAML 配置路径 | 是 |
| `-f` | JSON 文件路径 | 仅 cf/af |
| `-o` | 输出目录 | 否（默认 output） |

---

## Python 编程调用（保持不变）

### 1. 生成文档结构 JSON
```python
from wordformat.set_tag import set_tag_main

set_tag_main(
    docx_path="your_document.docx",
    configpath="example/undergrad_thesis.yaml"
)
```

### 2. 执行格式检查
```python
from wordformat.set_style import auto_format_thesis_document

auto_format_thesis_document(
    jsonpath="output/论文_1744123456.json",
    docxpath="your_document.docx",
    configpath="example/undergrad_thesis.yaml",
    savepath="check_result/",
    check=True
)
```

### 3. 执行自动格式化
```python
from wordformat.set_style import auto_format_thesis_document

auto_format_thesis_document(
    jsonpath="output/论文_1744123456.json",
    docxpath="your_document.docx",
    configpath="example/undergrad_thesis.yaml",
    savepath="final_output/",
    check=False
)
```

---

## API 调用（无变化）

### 启动服务
```bash
uv venv
source venv/bin/activate
uv sync
uv run start_api.py
```

启动服务查看接口文档：http://127.0.0.1:8000/docs

---
