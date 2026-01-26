# WordFormat 项目

> 论文格式自动化处理工具

![License](https://img.shields.io/github/license/AfishInLake/WordFormat?color=blue)
![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Status](https://img.shields.io/badge/status-开发中-orange)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## 项目简介

**WordFormat** 是一个基于 Python 的 Word 文档自动化格式检查工具，专门用于学术论文（如本科/硕士/博士论文）的格式合规性审查。该项目能够自动分析 Word 文档的结构，识别不同类型的段落（标题、摘要、正文、参考文献等），并根据预设的格式规范进行校验，在不符合规范的位置自动添加批注。

## 功能特性

- **智能文档结构解析**：使用大语言模型（LLM）自动识别文档中的标题、摘要、正文、参考文献等不同类型的段落
- **多级标题识别**：支持一级、二级、三级标题的自动识别和格式检查
- **格式规范校验**：支持段落格式（对齐方式、行距、缩进等）和字符格式（字体、字号、颜色、加粗等）的全面检查
- **多语言支持**：支持中英文文档的格式检查，包括中英文字体的分别设置
- **自动批注功能**：在发现格式问题时自动添加 Word 批注，便于作者修改
- **灵活配置**：支持通过 YAML/JSON 配置文件自定义格式规范

## 技术架构

### 核心模块

- **[base.py](file://G:\desktop\WordFormat\src\base.py)**：基础文档处理类，封装了 Word 文档的基本操作和 LLM 交互
- **`rules/`**：各种文档元素的格式检查规则，包括：
  - [heading.py](WordFormat/src/rules/heading.py) - 标题格式检查
  - [abstract.py](WordFormat/src/rules/abstract.py) - 摘要格式检查  
  - [body.py](WordFormat/src/rules/body.py) - 正文格式检查
  - [keywords.py](WordFormat/src/rules/keywords.py) - 关键词格式检查
  - [references.py](WordFormat/src/rules/references.py) - 参考文献格式检查
- **`config/datamodel.py`**：数据模型定义，使用 Pydantic 进行配置验证
- **`style/check_format.py`**：格式检查核心工具类
- **`word_structure/`**：文档结构构建和树形表示

### 格式规范

项目支持以下格式参数的配置：

#### 段落格式
- 对齐方式：左对齐、居中对齐、右对齐、两端对齐、分散对齐
- 行距：单倍行距、1.5倍、双倍
- 缩进：无缩进、1字符、2字符、3字符
- 段前段后间距

#### 字符格式
- 中文字体：宋体、黑体、楷体、仿宋、微软雅黑等
- 英文字体：Times New Roman、Arial、Calibri 等
- 字号：初号至八号及 px/pt 数值
- 颜色：BLACK、WHITE、RED 等标准颜色或 RGB 值
- 样式：加粗、斜体、下划线

## 快速开始

### 环境要求

- Python 3.8+
- 依赖库（见 requirements.txt）
- 大语言模型 API 密钥（用于文档结构解析）

### 安装步骤

```bash
# 克隆项目
git clone <repository-url>
cd WordFormat

# 安装依赖
uv sync

# 配置 API 密钥等环境变量
.env
```


### 使用方法

#### 1. 文档结构标记

```python
from src.set_tag import main

# 解析文档结构并保存为 JSON
main(docx_path="your_document.docx", json_save_path="tmp/")
```


#### 2. 格式检查

```python
from src.set_style import set_style

# 对文档进行格式检查并添加批注
set_style(
    jsonpath="path/to/structure.json",
    docxpath="input.docx", 
    savepath="output_with_comments.docx",
    configpath="config.yaml"
)
```


## 项目结构

```
src/
├── agent/           # AI 代理相关
├── config/          # 配置文件和数据模型
├── rules/           # 各类格式检查规则
├── style/           # 样式检查和应用工具
├── word_structure/  # 文档结构构建
├── base.py          # 基础文档处理类
├── set_style.py     # 格式检查主入口
├── set_tag.py       # 文档标记主入口
└── ...
```


## 配置说明

项目支持通过配置文件自定义格式规范，主要配置项包括：

- 标题格式（各级标题）
- 摘要格式（中英文）
- 关键词格式
- 正文格式
- 参考文献格式
- 致谢格式

## 贡献指南

欢迎提交 Issue 和 Pull Request 来改进项目。对于重大变更，请先开 Issue 讨论变更方案。

## 许可证

[在此处添加许可证信息]

## 联系方式

如有问题或建议，请联系项目维护者。