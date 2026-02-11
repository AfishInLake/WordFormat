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
make install
```

#### 3. 配置环境变量
创建 `.env` 文件，配置 HOST PORT等必要参数：
```env
HOST="127.0.0.1"
# 配置服务端口
PORT="8000"
```
#### 4. 启动API服务
- 或查看 `方式三：api调用` 
```bash
# 在虚拟环境下运行
make server
```

#### 5. 构建.exe程序
```bash
# 在虚拟环境下运行
make build
```

### 核心使用方法

WordFormat 提供**命令行交互**和**Python 编程调用**两种使用方式，推荐使用命令行方式（更便捷）。

#### 方式一：命令行使用（推荐）
工具提供**三种核心执行模式**，全局参数统一管控核心文件路径，子命令参数管控配置与输出，参数顺序遵循「全局参数 → 子命令 → 子命令专属参数」规则，灵活控制处理流程：

##### 全局参数通用规则
所有模式均需先指定全局参数（`--docx`/`-d` 待处理Word文档、`--json`/`-jf` JSON完整路径），`--json-dir`/`-j` 仅为 `generate-json` 模式的JSON生成目录，其他模式忽略该参数。

##### 1. 生成文档结构 JSON（第一步：解析文档结构）
解析 Word 文档并生成结构化 JSON 文件（可手动调整JSON后再执行校验/格式化），**配置文件为必填项**：
```bash
# 基础用法（JSON 生成到默认 tmp/ 目录，自动匹配Word同名）
wordformat -d your_document.docx -jf your_document.json generate-json -c example/undergrad_thesis.yaml

# 自定义 JSON 生成目录（生成到 output/ 目录，自动匹配Word同名）
wordformat -d your_document.docx -jf your_document.json -j output/ generate-json -c example/undergrad_thesis.yaml
```

##### 2. 执行格式校验（第二步：仅检查+添加批注）
使用生成/修改后的**完整JSON路径**执行格式校验，不在原文档修改，仅在违规位置添加Word批注，生成带批注的新文档，**配置文件为必填项**：
```bash
# 基础用法（使用指定的完整JSON路径，校验后文档保存到默认 output/ 目录）
wordformat -d your_document.docx -jf output/your_document_edited.json check-format -c example/undergrad_thesis.yaml

# 自定义校验后文档输出目录
wordformat -d your_document.docx -jf output/your_document.json check-format -c example/undergrad_thesis.yaml -o check_result/
```

##### 3. 执行格式格式化（第三步：自动修正格式）
使用指定的**完整JSON路径**，根据配置文件**自动修正**文档格式问题，生成格式化后的新文档，**配置文件为必填项**：
```bash
# 基础用法（使用指定的完整JSON路径，格式化后文档保存到默认 output/ 目录）
wordformat -d your_document.docx -jf output/your_document.json apply-format -c example/undergrad_thesis.yaml

# 自定义格式化后文档输出目录
wordformat -d your_document.docx -jf output/your_document_edited.json apply-format -c example/grad_thesis.yaml -o final_format/
```

##### 实际测试示例（贴合真实使用场景）
```bash
# 1. 生成JSON（到output目录）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json -j .\output\ generate-json -c .\example\undergrad_thesis.yaml
wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" generate-json --config "example/undergrad_thesis.yaml"
# 2. 执行格式化（使用上一步生成的完整JSON路径）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json apply-format -c .\example\undergrad_thesis.yaml
wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" apply-format --config "example/undergrad_thesis.yaml"

# 3. 执行校验（自定义输出目录）
wordformat -d .\tmp\毕业设计说明书.docx -jf .\output\毕业设计说明书.json check-format -c .\example\undergrad_thesis.yaml -o .\check_output\
(wordparse) PS G:\desktop\WordFormat> wordformat --docx "G:\desktop\论文语料集\1 (2).docx" --json "test02s/1.json" check-format --config "example/undergrad_thesis.yaml"
```

##### 命令行参数详细说明
| 层级       | 参数/子命令         | 简写 | 必填 | 适用模式          | 说明                                                                 |
|------------|---------------------|------|------|-------------------|----------------------------------------------------------------------|
| **全局参数** | `--docx`            | `-d` | ✅ 是 | 所有模式          | 待处理的 Word 文档**完整路径**，例如：`tmp/毕业设计说明书.docx`       |
| **全局参数** | `--json`            | `-jf`| ✅ 是 | 所有模式          | JSON 文件**完整路径**，例如：`output/毕业设计说明书.json`             |
| **全局参数** | `--json-dir`        | `-j` | ❌ 否 | 仅generate-json   | JSON 生成目录（其他模式忽略），默认：`tmp/`，自动创建不存在的目录     |
| **子命令**   | `generate-json`     | -    | -    | 结构解析          | 仅生成文档结构 JSON 文件，需配合 `-c` 指定配置文件                    |
| **子命令**   | `check-format`      | -    | -    | 格式校验          | 仅执行格式检查，在违规位置添加批注，不修改原文档                      |
| **子命令**   | `apply-format`      | -    | -    | 格式修正          | 按规范自动修正格式问题，生成格式化后的新文档                          |
| 子命令参数   | `--config`          | `-c` | ✅ 是 | 所有子命令        | 格式配置 YAML**完整路径**，例如：`example/undergrad_thesis.yaml`      |
| 子命令参数   | `--output`          | `-o` | ❌ 否 | check/apply       | 校验/格式化后文档**保存目录**，默认：`output/`，自动创建不存在的目录  |

#### 方式二：Python 编程调用
适合集成到其他项目或自定义扩展，直接调用核心函数实现解析、校验、格式化：

##### 1. 生成文档结构 JSON

```python
from wordformat.set_tag import set_tag_main as set_tag_main

# 解析文档并生成 JSON 结构文件
set_tag_main(
    docx_path="your_document.docx",  # 原始Word文档路径
    json_save_path="output/your_document.json",  # JSON保存完整路径
    configpath="example/undergrad_thesis.yaml"  # 格式配置文件路径
)
```

##### 2. 执行格式检查（仅添加批注）
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

##### 3. 执行格式自动修正
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

##### 4. 查看文档结构树
辅助工具，验证文档结构解析是否正确：
```bash
# 1. 修改 print_tree.py 中的 JSON_PATH 和 YAML_PATH 为实际文件路径
# 2. 执行命令查看结构化树状结果
python print_tree.py 
```

#### 方式三：api调用
- 在 `.env` 配置以下参数
```.env
# 配置服务地址
HOST="127.0.0.1"
# 配置服务端口
PORT="8000"
```
- 在命令窗口中执行以下命令启动服务
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
> 💡 提示：接口文档请访问`http://{HOST}:{PORT}/docs`
>   示例： [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)

## 配置文件说明

### 配置文件格式
格式规范通过 YAML 文件定义，示例参考项目中 `example/undergrad_thesis.yaml`（本科毕业论文模板）、`example/grad_thesis.yaml`（研究生毕业论文模板），所有配置项均有清晰注释。

### 自定义配置
可根据不同学校/期刊的格式要求，灵活修改以下核心配置项，适配专属格式规范：
- 各级标题（一级/二级/三级）的字体、字号、对齐方式、行距、缩进规则
- 正文的行距、首行缩进字符数、段前段后间距
- 中/英文字体分别设置（完美支持中英文混合文档）
- 特殊段落（摘要/关键词/参考文献/致谢）的专属格式规则
- 字符格式（加粗/斜体/下划线/字体颜色）的全局/局部规则

---
## 常见问题

### Q1：命令行执行报参数解析错误？
**核心原因**：参数顺序错误，未遵循「全局参数 → 子命令 → 子命令专属参数」规则  
**错误示例**：子命令参数（`-c`）放在子命令前面
```bash
wordformat -d doc.docx -c config.yaml apply-format -jf output/doc.json
```
**正确示例**：先全局参数，再子命令，最后子命令参数
```bash
wordformat -d doc.docx -jf output/doc.json apply-format -c config.yaml
```

### Q2：generate-json 模式报 JSON 文件不存在？
**原因**：`--json/-jf` 为全局必填参数，该模式下仅作**参数占位**，无需提前创建JSON文件，工具会自动生成  
**解决**：直接指定JSON保存路径即可，工具会自动创建文件及上级目录，示例：
```bash
wordformat -d doc.docx -jf output/未创建的json文件.json generate-json -c config.yaml
```

### Q3：check-format/apply-format 模式报 JSON 文件不存在？
**排查方向**：
1. 确认 `--json/-jf` 指定的是**完整JSON文件路径**，而非目录
2. 检查JSON路径拼写是否正确（区分大小写，Windows系统注意反斜杠/正斜杠）
3. 确认已执行 `generate-json` 模式生成JSON，或手动创建了合法的JSON文件

### Q4：JSON 文件生成失败？
**排查方向**：
1. 检查 `.env` 文件中 LLM 相关配置（API_KEY/MODEL/MODEL_URL）是否正确且有效
2. 确认待处理Word文档路径正确，文件未损坏、可正常打开
3. 确保Word文档内容完整，无空文档或特殊字符导致的解析失败
4. 检查配置文件（`-c`）路径正确，YAML格式合法无语法错误

### Q5：格式校验/格式化无效果、未生成批注/修正？
**排查方向**：
1. 检查JSON文件中各段落的 `category` 字段是否正确（标题/正文/摘要等是否识别准确）
2. 确认配置文件（YAML）中的规则项无拼写错误（如字体名称、对齐方式）
3. 检查Word文档的段落/字符是否有特殊格式（如手动换行/分节符）导致的解析异常
4. 验证配置文件中的规则是否与文档实际格式一致（无违规则不会生成批注/修正）

### Q6：执行后生成的Word文档无法打开？
**排查方向**：
1. 确认使用的是 Python 3.11+ 版本，避免版本兼容问题
2. 检查待处理Word文档为 `.docx` 格式，不支持 `.doc` 旧格式
3. 确保生成的JSON文件语法合法，手动修改后未出现JSON格式错误（如缺少逗号/引号）

### Q7：为什么使用 SSH 推送时 Git LFS 会卡住或失败？
**排查方向**：
1. GitHub **不支持通过 SSH 协议上传 Git LFS 文件**。
虽然普通代码可以通过 `git@github.com` 正常推送，但 LFS 在 SSH 模式下会尝试执行 `git-lfs-transfer` 命令，而 GitHub 的 SSH 服务并不识别该命令，导致连接被立即关闭（报错如 `EOF` 或 `Unable to negotiate version`）。

**正确做法**：  
- 将远程仓库地址改为 **HTTPS**：
  ```bash
  git remote set-url origin https://github.com/AfishInLake/WordFormat.git
  ```
- 使用 **Personal Access Token (PAT)** 作为密码（非 GitHub 登录密码）
- 确保网络能访问 GitHub（国内用户若遇 LFS 上传慢/失败，可考虑配合代理或使用 Gitee 托管 LFS）

> 💡 提示：只要远程地址是 `git@github.com:...`，Git LFS **一定无法上传成功**——这是 GitHub 的设计限制，非配置错误。

---
## 贡献指南

### 贡献方式
1. 提交 Issue：反馈 bug、提出新功能建议、补充格式配置模板需求
2. 提交 Pull Request：修复 bug、新增格式检查规则、完善核心功能、优化文档
3. 完善配置模板：补充不同学校/期刊/学位的论文格式配置YAML文件，提交到 `example/` 目录

### 开发规范
1. 代码遵循 PEP 8 编码规范，使用清晰的变量/函数命名
2. 新增功能/规则需添加对应的注释，关键逻辑添加使用示例
3. 重大功能变更需先提交 Issue 讨论实现方案，避免重复开发
4. 提交 PR 前需确保代码可正常运行，无语法错误和逻辑bug
5. 新增配置项需同步更新文档中的「配置文件说明」部分

## 许可证

[MIT License](LICENSE) - 允许自由使用、修改和分发，需保留原作者声明。

## 联系方式

- 反馈渠道：GitHub Issues（优先推荐，方便问题跟踪和交流）
- 联系邮箱：1593699665@qq.com

## 项目贡献者
待补充，欢迎各位开发者提交PR贡献代码，一起完善这个工具！