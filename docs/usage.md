# 使用指南

本文档详细说明 WordFormat 的使用方法，包括命令行交互、Python 编程调用和 API 调用三种方式。

## 命令行使用（推荐）

WordFormat 提供三种核心执行模式，全局参数统一管控核心文件路径，子命令参数管控配置与输出。

### 全局参数通用规则

所有模式均需先指定全局参数（`--docx`/`-d` 待处理Word文档、`--json`/`-jf` JSON完整路径）。

### 1. 生成文档结构 JSON

解析 Word 文档并生成结构化 JSON 文件（可手动调整JSON后再执行校验/格式化），**配置文件为必填项**：

```bash
# 基础用法
wordformat -d your_document.docx -jf your_document.json generate-json -c example/undergrad_thesis.yaml

# 自定义 JSON 生成路径
wordformat -d your_document.docx -jf output/your_document.json generate-json -c example/undergrad_thesis.yaml
```

### 2. 执行格式校验

使用生成/修改后的**完整JSON路径**执行格式校验，不在原文档修改，仅在违规位置添加Word批注，生成带批注的新文档，**配置文件为必填项**：

```bash
# 基础用法（使用指定的完整JSON路径，校验后文档保存到默认 output/ 目录）
wordformat -d your_document.docx -jf output/your_document_edited.json check-format -c example/undergrad_thesis.yaml

# 自定义校验后文档输出目录
wordformat -d your_document.docx -jf output/your_document.json check-format -c example/undergrad_thesis.yaml -o check_result/
```

### 3. 执行格式格式化

使用指定的**完整JSON路径**，根据配置文件**自动修正**文档格式问题，生成格式化后的新文档，**配置文件为必填项**：

```bash
# 基础用法（使用指定的完整JSON路径，格式化后文档保存到默认 output/ 目录）
wordformat -d your_document.docx -jf output/your_document.json apply-format -c example/undergrad_thesis.yaml

# 自定义格式化后文档输出目录
wordformat -d your_document.docx -jf output/your_document_edited.json apply-format -c example/grad_thesis.yaml -o final_format/
```

### 实际测试示例

```bash
# 1. 生成JSON（到output目录）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json -j .\output\ generate-json -c .\example\undergrad_thesis.yaml
wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" generate-json --config "example/undergrad_thesis.yaml"

# 2. 执行格式化（使用上一步生成的完整JSON路径）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json apply-format -c .\example\undergrad_thesis.yaml
wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" apply-format --config "example/undergrad_thesis.yaml"

# 3. 执行校验（自定义输出目录）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json check-format -c .\example\undergrad_thesis.yaml -o .\check_output\
wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" check-format --config "example/undergrad_thesis.yaml"
```

### 命令行参数详细说明

| 层级       | 参数/子命令         | 简写 | 必填 | 适用模式          | 说明                                                                 |
|------------|---------------------|------|------|-------------------|----------------------------------------------------------------------|
| **全局参数** | `--docx`            | `-d` | ✅ 是 | 所有模式          | 待处理的 Word 文档**完整路径**，例如：`tmp/毕业设计说明书.docx`       |
| **全局参数** | `--json`            | `-jf`| ✅ 是 | 所有模式          | JSON 文件**完整路径**，例如：`output/毕业设计说明书.json`             |
| **子命令**   | `generate-json`     | -    | -    | 结构解析          | 仅生成文档结构 JSON 文件，需配合 `-c` 指定配置文件                    |
| **子命令**   | `check-format`      | -    | -    | 格式校验          | 仅执行格式检查，在违规位置添加批注，不修改原文档                      |
| **子命令**   | `apply-format`      | -    | -    | 格式修正          | 按规范自动修正格式问题，生成格式化后的新文档                          |
| 子命令参数   | `--config`          | `-c` | ✅ 是 | 所有子命令        | 格式配置 YAML**完整路径**，例如：`example/undergrad_thesis.yaml`      |
| 子命令参数   | `--output`          | `-o` | ❌ 否 | check/apply       | 校验/格式化后文档**保存目录**，默认：`output/`，自动创建不存在的目录  |

## Python 编程调用

适合集成到其他项目或自定义扩展，直接调用核心函数实现解析、校验、格式化。

### 1. 生成文档结构 JSON

```python
from wordformat.set_tag import set_tag_main as set_tag_main

# 解析文档并生成 JSON 结构文件
set_tag_main(
    docx_path="your_document.docx",  # 原始Word文档路径
    json_save_path="output/your_document.json",  # JSON保存完整路径
    configpath="example/undergrad_thesis.yaml"  # 格式配置文件路径
)
```

### 2. 执行格式检查（仅添加批注）

```python
from wordformat.set_style import auto_format_thesis_document

# 执行格式校验，生成带批注的文档（check=True 仅校验，不修改）
auto_format_thesis_document(
    jsonpath="output/your_document.json",  # 结构JSON完整路径
    docxpath="your_document.docx",         # 原始Word文档路径
    configpath="example/undergrad_thesis.yaml",  # 格式配置文件
    savepath="check_result/",              # 校验后文档保存目录
    check=True                             # 仅校验模式
)
```

### 3. 执行格式自动修正

```python
from wordformat.set_style import auto_format_thesis_document

# 执行格式自动修正（check=False 格式化模式）
auto_format_thesis_document(
    jsonpath="output/your_document_edited.json",  # 手动调整后的JSON路径
    docxpath="your_document.docx",                 # 原始Word文档路径
    configpath="example/grad_thesis.yaml",         # 格式配置文件
    savepath="final_format/",                      # 格式化后文档保存目录
    check=False                                    # 格式化模式
)
```

### 4. 查看文档结构树

辅助工具，验证文档结构解析是否正确：

```bash
# 1. 修改 print_tree.py 中的 JSON_PATH 和 YAML_PATH 为实际文件路径
# 2. 执行命令查看结构化树状结果
python print_tree.py
```

## API 调用

### 启动 API 服务

```bash
# 创建venv环境
uv venv
# 激活venv环境
source venv/bin/activate
# 安装依赖
uv sync
# 启动服务
uv run start_api.py
```

### 接口文档

启动服务后，可访问以下地址查看详细的 API 文档：

```
http://{HOST}:{PORT}/docs
```

示例：[http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)

### 常用 API 端点

- `POST /generate-json` - 生成文档结构 JSON
- `POST /check-format` - 执行格式校验
- `POST /apply-format` - 执行格式格式化
- `GET /download/{filename}` - 下载格式化/校验后的Word文档

### API 请求示例

#### 生成文档结构 JSON

```bash
curl -X POST "http://127.0.0.1:8000/generate-json" \
  -F "docx_file=@your_document.docx" \
  -F "config_file=@example/undergrad_thesis.yaml"
```

#### 执行格式校验

```bash
curl -X POST "http://127.0.0.1:8000/check-format" \
  -F "docx_file=@your_document.docx" \
  -F "config_file=@example/undergrad_thesis.yaml" \
  -F "json_data={\"paragraphs\": [...], ...}"
```

#### 执行格式格式化

```bash
curl -X POST "http://127.0.0.1:8000/apply-format" \
  -F "docx_file=@your_document.docx" \
  -F "config_file=@example/undergrad_thesis.yaml" \
  -F "json_data={\"paragraphs\": [...], ...}"
```
