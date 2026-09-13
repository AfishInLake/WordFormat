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

## 按需使用原则（重要）

每个命令只做一件事。请**只执行满足用户需求所需的最小命令组合**，不要执行多余的步骤：

- 用户说"修标签"就修标签（gj → 改 JSON → 需要时 af），不要顺带跑 cf 检查格式
- 用户说"检查格式"就只产出检查结果（gj → cf），不执行 af 改文档
- 用户说"格式化"就直接格式化（gj → 改必要的标签 → af）；cf（加批注）不是 af 的前置步骤
- `gj` 每次运行会**重新识别并覆盖** JSON——已手动改过 JSON 后不要再重跑 gj
- output/ 下已有 JSON 时直接复用，不重复生成

## 决策表：根据用户意图选择命令

| 用户需求 | 必要命令 | 可选命令 | 不要做 |
|---|---|---|---|
| 修正/查看识别分类（标签） | 改 JSON 的 `category`；无 JSON 时先 `gj`；要应用到文档再 `af` | `tree --confidence` 定位可疑段 | 不跑 `cf`；改完 JSON 不重跑 `gj` |
| 替换正文文字 | 改 JSON 的 `replace`；无 JSON 时先 `gj`；再 `af` | — | 不跑 `cf`、`tree` |
| 只检查格式（加批注） | `gj`（如无 JSON）→ `cf` | — | 不跑 `af` |
| 修正格式（应用到文档） | `gj`（如无 JSON）→ `af` | `tree` 先复核分类 | `cf` 非必需，用户想看批注才插入 |
| 查看结构树 | `tree -f xxx.json` | `--confidence` / `--index` | 需已有 JSON；不跑 `gj`（除非没有 JSON） |
| Markdown 转 Word | `md` | — | 与 docx 流程互不相关 |
| 生成配置模板 | `config` | — | — |

只有用户明确要求"完整流程/全套"时，才按场景顺序执行多步流水线。

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

## 常见场景最小编排

### 1. 修复识别错误的标签

用户让"看看分类对不对""帮我把识别错的标签改一下"时：

```bash
# 第 1 步：JSON 不存在才生成（已存在则直接复用，避免覆盖修改）
wordf gj -d 论文.docx -c config.yaml

# 第 2 步（可选）：低置信度的段往往是分类错误，配合 --index 拿到序号
wordf tree -f output/xxx.json --confidence --index

# 第 3 步：编辑 output/xxx.json，修正相应段落的 category 字段

# 第 4 步：仅在用户要求把修正应用到文档时执行
wordf af -d 论文.docx -c config.yaml -f output/xxx.json
```

- 用户只想"看/改标签"：到第 3 步就结束，不执行任何后续命令。
- 不执行 `cf`：格式检查与标签修正无关。
- 修改 JSON 后不重跑 `gj`：会覆盖手工修改。

### 2. 纯文本替换

同场景 1，仅字段不同：`gj` → 编辑 `replace` 字段 → `af`。不需要 `tree` 和 `cf`。

### 3. 仅格式检查（不改原文）

```bash
wordf gj -d 论文.docx -c config.yaml   # JSON 不存在时执行
wordf cf -d 论文.docx -c config.yaml -f output/xxx.json
```

`cf` 生成带批注的检查版 docx，不修改原文。到此结束，不跑 `af`。

### 4. 格式修正（生成格式化版文档）

```bash
wordf gj -d 论文.docx -c config.yaml   # JSON 不存在时执行
# 如发现识别错误的标签，先编辑 category 再继续
wordf af -d 论文.docx -c config.yaml -f output/xxx.json
```

`af` 直接产出格式化后的 docx。仅当用户想先看错误清单时，才在中间插入 `cf`。

### 5. Markdown 转 Word

```bash
wordf md -d 论文.md -c config.yaml
# 输出 output/论文--生成版.docx
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