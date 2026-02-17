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
- **智能文档结构解析**：基于 ONNX 模型自动识别文档中的标题、摘要、正文、关键词、参考文献等段落类型，生成结构化 JSON 配置文件
- **精细化格式校验**：支持段落格式（对齐、行距、缩进、段间距）和字符格式（字体、字号、颜色、加粗/斜体/下划线）的全维度检查
- **多级别标题管理**：精准识别一级/二级/三级标题，支持自定义标题格式规范校验
- **多语言适配**：区分中英文字体/格式规则，完美支持中英文混合文档的格式检查
- **灵活的交互方式**：支持「生成结构文件→手动调整→执行校验」的分步流程，兼顾自动化与灵活性

### 实用功能
- **自动批注生成**：在格式违规位置自动添加 Word 批注，标注问题类型和修正建议
- **格式一键修正**：支持根据规范自动修正部分常见格式问题（如标题字号、正文行距）
- **自定义配置**：通过 YAML 配置文件灵活定义格式规范，适配不同学校/期刊的格式要求
- **跨平台兼容**：支持 Windows/macOS/Linux 系统，基于 python-docx 实现跨平台 Word 文档处理

## 快速开始

### 环境要求
- Python 3.11+（推荐 3.11 及以上版本）
- 依赖管理工具：uv（推荐）或 pip

### 安装步骤

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

3. **配置环境变量**
   创建 `.env` 文件，配置 HOST、PORT 等必要参数：
   ```env
   HOST="127.0.0.1"
   # 配置服务端口
   PORT="8000"
   ```

4. **启动API服务**
   ```bash
   # 在虚拟环境下运行
   make server
   ```

5. **构建.exe程序**
   ```bash
   # 在虚拟环境下运行
   make build
   ```

## 核心使用方法

### 命令行使用（推荐）

WordFormat 提供三种核心执行模式：

1. **生成文档结构 JSON**
   ```bash
   wordformat -d your_document.docx -jf your_document.json generate-json -c example/undergrad_thesis.yaml
   ```

2. **执行格式校验**
   ```bash
   wordformat -d your_document.docx -jf output/your_document.json check-format -c example/undergrad_thesis.yaml
   ```

3. **执行格式格式化**
   ```bash
   wordformat -d your_document.docx -jf output/your_document.json apply-format -c example/undergrad_thesis.yaml
   ```

## 详细文档

更多详细文档请查看 `docs/` 目录：

- [安装指南](docs/installation.md) - 环境要求和安装步骤
- [使用指南](docs/usage.md) - 命令行、Python编程和API调用的详细使用方法
- [配置文件说明](docs/configuration.md) - 格式规范配置项和自定义配置方法
- [常见问题](docs/faq.md) - 常见问题及解决方案
- [贡献指南](docs/contributing.md) - 如何为项目贡献代码和文档
- [技术架构](docs/architecture.md) - 项目的技术架构和实现原理

## 许可证

[MIT License](LICENSE) - 允许自由使用、修改和分发，需保留原作者声明。

## 联系方式

- 反馈渠道：GitHub Issues（优先推荐，方便问题跟踪和交流）
- 联系邮箱：1593699665@qq.com

## 项目贡献者

待补充，欢迎各位开发者提交PR贡献代码，一起完善这个工具！
