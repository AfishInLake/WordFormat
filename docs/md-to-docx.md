# Markdown 转 Word 功能

## 概述

`wordf md` 命令支持将 Markdown 文件直接转换为格式化的 .docx 文档。Markdown 提供文本内容和文档结构（标题层级、段落划分），所有格式（字体、字号、段落样式、编号等）由 YAML 配置文件驱动，与现有的 docx 格式化流程共享同一套配置。

## 使用场景

- **论文初稿写作**：先用 Markdown 写论文草稿（语法简单、纯文本、易于版本控制），定稿后一键导出为符合学校格式规范的 Word 文档
- **自动化报告生成**：在数据分析、实验报告等场景中，用脚本生成 Markdown 文件，再批量转换为格式统一的 Word 文档
- **混合工作流**：团队中部分成员用 Markdown 协作，最终交付 Word 格式给导师/期刊/学校系统

## 使用方法

### 命令行

```bash
# 基本用法
wordf md -d 论文.md -c 配置.yaml

# 指定输出目录
wordf md -d 论文.md -c 配置.yaml -o output/
```

### Python API

```python
from wordformat.pipeline.orchestrate import md_to_docx

output_path = md_to_docx(
    md_path="thesis.md",
    config_path="config.yaml",
    save_dir="output/",
)
print(f"生成文件: {output_path}")
```

## Markdown 语法映射

| Markdown | 文档类别 |
|----------|----------|
| `# 标题` | heading_level_1（一级标题） |
| `## 标题` | heading_level_2（二级标题） |
| `### 标题` | heading_level_3（三级标题） |
| `#### 标题` | heading_level_3（四级及以上统一映射为三级） |
| 普通段落 | body_text（正文） |
| `![](path)` | figure_image（图片占位） |

列表、引用块、代码块等均展开为 body_text。

## 工作流程

```
Markdown 文件
    │
    ▼ mistune 解析
AST 块列表（heading / paragraph / list / table / ...）
    │
    ▼ 类别映射
扁平段落列表 [{category, paragraph, ...}]
    │
    ▼ DocumentTreeBuilder
FormatNode 文档树
    │
    ▼ DocumentCreationStage
新建 .docx Document → 逐节点 add_paragraph
    │
    ▼ StyleDefinitionFixStage
修正内置样式定义（Normal, Heading 1, ...）
    │
    ▼ FormattingExecutionStage
应用 YAML 配置中的格式规则
    │
    ▼ PostProcessingStage
标题自动编号 + 引用超链接
    │
    ▼ 保存
{文件名}--生成版.docx
```

## 格式应用与批注

md→docx 的格式应用逻辑与现有 `wordf af`（apply format）流程一致：

- **段落样式**（对齐、间距、缩进）和**字符样式**（字体、字号、颜色）由 YAML 配置定义，自动应用到每个段落
- apply 模式先应用格式，再 diff 检查残留差异——只有实际无法修正的问题才生成批注，已成功修正的变更不产生噪音
- 业务规则（标点检测、图片对齐等）照常运行，发现的问题以批注形式标注在文档中
- 批注颜色由问题严重等级决定：**错误**为红色，**提醒**为蓝色

## 注意事项

- Markdown 仅提供结构和文本，所有格式完全由 YAML 配置控制——配置文件的质量决定了输出文档的格式合规度
- 引用半角标点（如英文引号 `""`）不会自动转换，标点检测规则会以批注形式提醒
- 图片语法 `![](path)` 生成的是文本占位段，不会真正插入图片文件
- 如需对生成的 docx 做进一步检查，可先 `wordf gj` 生成 JSON，再 `wordf cf` 校验

## 依赖

Markdown 解析基于 [mistune](https://github.com/lepture/mistune) 库（>=3.0.0），支持标准 Markdown 语法及表格、删除线、链接等扩展。
