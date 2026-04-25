# WordFormat

> 论文格式自动化处理工具

![License](https://img.shields.io/github/license/AfishInLake/WordFormat?color=blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![PyPI](https://img.shields.io/pypi/v/wordformat?color=blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## 项目简介

**WordFormat** 是一个基于 Python 开发的 Word 文档自动化格式检查与修正工具，专为学术论文（本科/硕士/博士毕业论文、期刊论文等）的格式合规性审查设计。该工具能够智能解析 Word 文档结构，识别标题、摘要、正文、参考文献等不同段落类型，并依据自定义的格式规范自动校验文档格式，支持在违规位置添加批注或直接修正格式问题，大幅提升论文格式审核效率。

## 功能特性

### 核心能力
- **智能文档结构解析**：基于 ONNX 模型自动识别文档中的标题、摘要、正文、关键词、参考文献等段落类型，生成结构化 JSON 配置文件
- **精细化格式校验**：支持段落格式（对齐、行距、缩进、段间距）和字符格式（字体、字号、颜色、加粗/斜体/下划线）的全维度检查
- **多级别标题管理**：精准识别一级/二级/三级标题，支持自定义标题格式规范校验
- **多语言适配**：区分中英文字体/格式规则，完美支持中英文混合文档的格式检查
- **灵活的交互方式**：支持「生成结构文件→手动调整→执行校验」的分步流程，兼顾自动化与灵活性

### 实用功能
- **自动批注生成**：在格式违规位置自动添加 Word 批注，标注问题类型和修正建议
- **格式一键修正**：支持根据规范自动修正部分常见格式问题（如标题字号、正文行距）
- **标题自动编号**：支持自动清除手动编号并应用 Word 自动编号，可自定义编号模板（如"第X章"、"1.1.1"）、编号后缀（制表符/空格/无）和缩进设置
- **自定义配置**：通过 YAML 配置文件灵活定义格式规范，适配不同学校/期刊的格式要求
- **跨平台兼容**：支持 Windows/macOS/Linux 系统，基于 python-docx 实现跨平台 Word 文档处理

## 视频教程
[点击直达B站视频](https://www.bilibili.com/video/BV1aiDjB8Edg/?spm_id_from=333.1007.top_right_bar_window_history.content.click&vd_source=490c514f59611dc0b600c1da58948e14)

## 快速开始

### 环境要求
- Python 3.10+（推荐 3.11 及以上版本）
- 依赖管理工具：uv（推荐）或 pip

### 安装步骤

**方式一：从 PyPI 安装（推荐普通用户）**

```bash
# 使用 pip
pip install wordformat

# 或使用 uv
uv add wordformat
```

安装完成后即可使用 `wf` 和 `wordformat` 命令。

**方式二：从源码安装（开发者）**

1. **克隆项目**
   ```bash
   git clone https://github.com/AfishInLake/WordFormat.git
   cd WordFormat
   ```

2. **安装依赖**
   ```bash
   make install
   # 或使用 pip
   pip install -e .
   ```

3. **下载模型**
   ```bash
   python scripts/download_model.py
   ```

## 核心使用方法

### 命令行使用（推荐）

WordFormat 提供三种核心执行模式：

```bash
# 1. 生成文档结构 JSON
wf gj -d 论文.docx -c 配置.yaml

# 2. 执行格式校验（添加批注，不修改原文）
wf cf -d 论文.docx -c 配置.yaml -f 结构文件.json

# 3. 执行自动格式化（一键修正格式）
wf af -d 论文.docx -c 配置.yaml -f 结构文件.json
```

更多详细用法请查看 [使用指南](https://github.com/AfishInLake/WordFormat/blob/master/docs/usage.md)

## AI Skill 集成

WordFormat 提供了 **SOLO Skill**，可在 SOLO 等 AI 助手平台中直接调用，实现对话式论文格式化。

### Skill 工作流程

Skill 包含两个独立任务，可分步执行：

| 任务 | 说明 | 产物 |
|------|------|------|
| **任务一：准备配置文件** | 根据格式要求生成/编辑 config.yaml | `config.yaml` |
| **任务二：执行格式化** | 使用配置文件对论文进行格式检查或修正 | `--标注版.docx` 或 `--修改版.docx` |

### Skill 目录结构

```
wordformat-skill/
├── SKILL.md                    # Skill 定义文件
├── scripts/
│   ├── setup_config.py         # 配置文件生成/验证脚本
│   ├── validate_json.py        # JSON 标签校验脚本
│   └── validate_config.py      # 配置文件验证脚本
└── data/
    ├── config.yaml             # 默认配置模板
    ├── config_spec.md          # 配置文件完整字段规范
    ├── config_editing_guide.md # 配置编辑指南
    ├── category_reference.md   # 段落分类参考
    └── font_size_table.md      # 字号对照表
```

### 预设配置库

项目内置了多所高校的论文格式预设，保存在 `presets/` 目录下，命名格式为 `{学校}_{学院/专业}_{论文类型}.yaml`，可直接使用或在此基础上修改。

## 详细文档

更多详细文档请查看 `docs/` 目录：

- [安装指南](https://github.com/AfishInLake/WordFormat/blob/master/docs/installation.md) - 环境要求和安装步骤
- [使用指南](https://github.com/AfishInLake/WordFormat/blob/master/docs/usage.md) - 命令行、Python编程和API调用的详细使用方法
- [配置文件说明](https://github.com/AfishInLake/WordFormat/blob/master/docs/configuration.md) - 格式规范配置项和自定义配置方法
- [常见问题](https://github.com/AfishInLake/WordFormat/blob/master/docs/faq.md) - 常见问题及解决方案
- [贡献指南](https://github.com/AfishInLake/WordFormat/blob/master/docs/contributing.md) - 如何为项目贡献代码和文档
- [技术架构](https://github.com/AfishInLake/WordFormat/blob/master/docs/architecture.md) - 项目的技术架构和实现原理

## 许可证

[Apache License 2.0](LICENSE) - 允许自由使用、修改和分发，需保留原作者声明。

## 联系方式

- 反馈渠道：GitHub Issues（优先推荐，方便问题跟踪和交流）
- 联系邮箱：1593699665@qq.com

## 项目贡献者

待补充，欢迎各位开发者提交PR贡献代码，一起完善这个工具！
