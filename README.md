# WordFormat 项目（目前项目还在快速迭代中）

> 论文格式自动化处理工具

![License](https://img.shields.io/github/license/AfishInLake/WordFormat?color=blue)
![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Status](https://img.shields.io/badge/status-开发中-orange)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## 项目简介

**WordFormat** 是一个基于 Python 开发的 Word 文档自动化格式检查与修正工具，专为学术论文（本科/硕士/博士毕业论文、期刊论文等）的格式合规性审查设计。该工具能够智能解析 Word 文档结构，识别标题、摘要、正文、参考文献等不同段落类型，并依据自定义的格式规范自动校验文档格式，支持在违规位置添加批注或直接修正格式问题，大幅提升论文格式审核效率。

## 功能特性

### 核心能力
- **智能文档结构解析**：基于大语言模型（LLM）自动识别文档中的标题、摘要、正文、关键词、参考文献等段落类型，生成结构化 JSON 配置文件
- **精细化格式校验**：支持段落格式（对齐、行距、缩进、段间距）和字符格式（字体、字号、颜色、加粗/斜体/下划线）的全维度检查
- **多级别标题管理**：精准识别一级/二级/三级标题，支持自定义标题格式规范校验
- **多语言适配**：区分中英文字体/格式规则，完美支持中英文混合文档的格式检查
- **灵活的交互方式**：支持「生成结构文件→手动调整→执行校验」的分步流程，兼顾自动化与灵活性

### 实用功能
- **自动批注生成**：在格式违规位置自动添加 Word 批注，标注问题类型和修正建议
- **格式一键修正**：支持根据规范自动修正部分常见格式问题（如标题字号、正文行距）
- **自定义配置**：通过 YAML 配置文件灵活定义格式规范，适配不同学校/期刊的格式要求
- **跨平台兼容**：支持 Windows/macOS/Linux 系统，基于 python-docx 实现跨平台 Word 文档处理

## 技术架构

### 核心模块
```
src/
├── agent/           # LLM 交互模块（文档结构识别）
├── config/          # 配置管理与数据模型（Pydantic 验证）
├── rules/           # 格式检查规则集
│   ├── heading.py   # 标题格式检查规则
│   ├── abstract.py  # 摘要格式检查规则
│   ├── body.py      # 正文格式检查规则
│   ├── keywords.py  # 关键词格式检查规则
│   └── references.py# 参考文献格式检查规则
├── style/           # 格式检查与修正核心
│   ├── check_format.py # 格式校验逻辑
│   └── set_style.py    # 格式修正与批注生成
├── word_structure/  # 文档结构构建与树形表示
├── base.py          # 基础文档处理类（Word 操作/LLM 封装）
├── set_tag.py       # 文档结构标记主入口（生成 JSON）
└── set_style.py     # 格式检查主入口（执行校验/修正）
```

### 格式规范配置项
| 类别       | 可配置参数                                                                 |
|------------|----------------------------------------------------------------------------|
| 段落格式   | 对齐方式、行距（单倍/1.5倍/双倍）、缩进（字符数）、段前段后间距            |
| 字符格式   | 中/英文字体、字号（初号~八号/px/pt）、颜色（标准色/RGB）、加粗/斜体/下划线 |
| 标题格式   | 各级标题的字体、字号、对齐方式、行距、缩进规则                             |
| 特殊段落   | 摘要/关键词/参考文献的专属格式规则                                         |

## 快速开始

### 环境要求
- Python 3.11+（推荐 3.11 及以上版本）
- 依赖管理工具：uv（推荐）或 pip
- 大语言模型 API 密钥（用于文档结构解析，如 OpenAI/智谱AI/百度文心等）

### 安装步骤

#### 1. 克隆项目
```bash
git clone https://github.com/AfishInLake/WordFormat.git
cd WordFormat
```

#### 2. 安装依赖
```bash
# 使用 uv（推荐）
uv sync

# 或使用 pip
pip install -e .
```

#### 3. 配置环境变量
创建 `.env` 文件，配置 LLM API 密钥等必要参数：
```env
# 示例：OpenAI API 配置
OPENAI_API_KEY=your-api-key
OPENAI_BASE_URL=https://api.openai.com/v1
# 或其他 LLM 配置（如智谱AI）
ZHIPU_API_KEY=your-zhipu-api-key
```

### 核心使用方法

WordFormat 提供**命令行交互**和**Python 编程调用**两种使用方式，推荐使用命令行方式（更便捷）。

#### 方式一：命令行使用（推荐）
工具提供三种执行模式，支持灵活控制流程：

##### 1. 生成文档结构 JSON（第一步）
解析 Word 文档并生成结构化 JSON 文件（可手动调整后再执行校验）：
```bash
# 基础用法（JSON 保存到默认 tmp/ 目录）
uv run main.py -d your_document.docx generate-json

# 自定义 JSON 保存目录
uv run main.py -d your_document.docx -j output/ generate-json
```

##### 2. 执行格式校验（第二步）
使用生成/修改后的 JSON 文件执行格式校验，生成带批注的文档：
```bash
# 基础用法（使用默认目录的 JSON 文件）
uv run main.py -d your_document.docx -j output/ check-format -c config/undergrad_thesis.yaml

# 自定义 JSON 文件路径
uv run main.py -d your_document.docx check-format -c config/undergrad_thesis.yaml -jf output/your_document_edited.json

# 自定义输出目录
uv run main.py -d your_document.docx check-format -c config/undergrad_thesis.yaml -jf output/your_document.json -o final_output/
```

##### 3. 完整流程（生成→手动编辑→校验）
一键执行全流程，中间暂停等待手动调整 JSON 文件：
```bash
# 基础用法
uv run main.py -d your_document.docx -j output/ full-pipeline -c config/undergrad_thesis.yaml

# 自定义输出目录
uv run main.py -d your_document.docx full-pipeline -c config/grad_thesis.yaml -o final_output/
```

##### 命令行参数说明
| 层级       | 参数/子命令         | 简写 | 必填 | 说明                                                                 |
|------------|---------------------|------|------|----------------------------------------------------------------------|
| 全局参数   | `--docx`            | `-d` | ✅ 是 | 待处理的 Word 文档路径                                               |
| 全局参数   | `--json-dir`        | `-j` | ❌ 否 | JSON 文件保存/读取目录（默认：tmp/）|
| 子命令     | `generate-json`     | -    | -    | 仅生成文档结构 JSON 文件                                             |
| 子命令     | `check-format`      | -    | -    | 仅执行格式校验                                                       |
| check-format | `--config`         | `-c` | ✅ 是 | 格式配置 YAML 文件路径                                               |
| check-format | `--json`           | `-jf`| ❌ 否 | 手动修改后的 JSON 文件路径（优先级高于 --json-dir）|
| check-format | `--output`         | `-o` | ❌ 否 | 校验后文档保存目录（默认：output/）|
| full-pipeline | `--config`        | `-c` | ❌ 否 | 格式配置 YAML 文件路径（默认：config/undergrad_thesis.yaml）|
| full-pipeline | `--output`        | `-o` | ❌ 否 | 校验后文档保存目录（默认：output/）|

#### 方式二：Python 编程调用
适合集成到其他项目或自定义扩展：

##### 1. 生成文档结构 JSON
```python
from src.set_tag import main as set_tag_main

# 解析文档并生成 JSON
set_tag_main(
    docx_path="your_document.docx",
    json_save_path="output/your_document.json"
)
```

##### 2. 执行格式检查与修正
```python
from src.set_style import auto_format_thesis_document

# 执行格式校验并生成带批注的文档
auto_format_thesis_document(
    jsonpath="output/your_document.json",  # 结构 JSON 文件路径
    docxpath="your_document.docx",         # 原始文档路径
    configpath="example/undergrad_thesis.yaml",  # 格式配置文件
    savepath="output/"                     # 输出目录
)
```

## 配置文件说明

### 配置文件格式
格式规范通过 YAML 文件定义，示例如下（`example/undergrad_thesis.yaml`）：

### 自定义配置
可根据不同学校/期刊的格式要求，修改以下配置项：
- 各级标题的字体、字号、对齐方式
- 正文的行距、缩进规则
- 中英文字体的分别设置
- 特殊段落（摘要/参考文献）的格式规则

## 常见问题

### Q1：命令行参数解析错误？
**原因**：子命令专属参数（如 `-c`）放在了子命令前面  
**解决**：调整参数顺序，遵循「全局参数 → 子命令 → 子命令参数」规则：
```bash
# 错误示例
uv run main.py -d doc.docx -c config.yaml check-format
# 正确示例
uv run main.py -d doc.docx check-format -c config.yaml
```

### Q2：JSON 文件生成失败？
**排查方向**：
1. 确认 LLM API 密钥配置正确且有效
2. 检查 Word 文档路径是否正确，文件是否可读取
3. 确保文档内容完整，无损坏

### Q3：格式校验无结果/批注未生成？
**排查方向**：
1. 检查 JSON 结构文件是否正确（手动调整后需确保 JSON 语法合法）
2. 确认配置文件格式正确，规则项无拼写错误
3. 检查文档段落是否被正确识别（可查看 JSON 文件中的 category 字段）

## 贡献指南

### 贡献方式
1. 提交 Issue：反馈 bug、提出新功能建议
2. 提交 Pull Request：修复 bug、新增功能、完善文档
3. 完善配置规则：补充不同类型论文的格式配置模板

### 开发规范
1. 代码遵循 PEP 8 规范
2. 新增功能需添加对应的单元测试
3. 重大变更需先提交 Issue 讨论方案
4. 提交 PR 前需确保代码通过 lint 检查

## 许可证

[MIT License](LICENSE) - 允许自由使用、修改和分发，需保留原作者声明。

## 联系方式

- 反馈渠道：GitHub Issues
- 联系邮箱：1593699665@qq.com

## 项目贡献者
