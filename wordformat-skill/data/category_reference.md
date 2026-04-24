# Category 类型参考

WordFormat AI 模型识别文档结构后，会为每个段落分配一个 `category` 类型。以下是模型实际输出的所有 category 类型及其精确判定规则。

## 完整 Category 列表（共 15 个）

### 标题类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `heading_level_1` | 段落必须以"第X章"或单个阿拉伯数字（如"1""2"）开头，后接空格和标题文字；仅为名词短语 | "1 绪论"、"3 系统设计" |
| `heading_level_2` | 段落必须以"X.Y"格式开头（如"1.1""2.3"），且**仅包含一个"."**；后接标题文字，无完整句子 | "1.1 研究背景"、"2.4 实验设置" |
| `heading_level_3` | 段落必须以"X.Y.Z"格式开头（如"1.1.1""3.2.5"），且**包含恰好两个"."**；后接标题文字，无论长短，只要无句子结构即视为三级标题 | "1.1.1 研究背景"、"2.3.4 性能对比" |
| `heading_fulu` | 段落等于"附录" | "附录" |

### 摘要类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `abstract_chinese_title` | 仅当段落是"摘要"或"摘 要"（允许尾随空格或冒号） | "摘要"、"摘要：" |
| `abstract_chinese_title_content` | 当且仅当摘要标题和摘要正文合并在同一个段落中 | "摘要 本文围绕校园二手物品交易平台..." |
| `abstract_english_title` | 仅当段落是"Abstract"（大小写不敏感，允许尾随空格或冒号） | "Abstract"、"Abstract: " |
| `abstract_english_title_content` | 当且仅当英文摘要标题和摘要正文合并在同一个段落中 | "Abstract This paper focuses on the design..." |

### 关键词类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `keywords_chinese` | 包含"关键词"或"关键字"，后面跟着术语列表 | "关键词：校园交易；二手物品；Django" |
| `keywords_english` | 包含"Keywords"（大小写不敏感），后面跟着英文术语 | "Keywords: Campus trading; Second-hand..." |

### 正文类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `body_text` | 仅当段落包含句子（有谓语动词、句号、逻辑连接词）、是摘要等段落或明确论述（如"本章""本文""包括""例如""若...则..."）时才归为此类 | "随着互联网技术的快速发展..." |

### 图表类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `caption_figure` | 以"图 X.Y"或"Figure X.Y"开头的图注 | "图2.1 系统架构图" |
| `caption_table` | 以"表 X.Y"或"Table X.Y"开头的表注 | "表3.1 用户表结构" |

### 参考文献类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `references_title` | 段落等于"参考文献"或"References" | "参考文献"、"References" |

### 致谢类

| Category | 判定规则 | 典型内容示例 |
|----------|----------|-------------|
| `acknowledgements_title` | 段落和"致谢"或"Acknowledgements"等词意思相近 | "致谢"、"Acknowledgements" |

## 检查要点

在第三步检查 JSON 文件时，重点注意以下容易误判的情况：

1. **摘要标题 vs 一级标题**："摘要" 可能被误识别为 `heading_level_1`，应为 `abstract_chinese_title`
2. **关键词 vs 正文**：关键词行可能被误识别为 `body_text`，应为 `keywords_chinese`
3. **参考文献标题 vs 一级标题**："参考文献" 可能被误识别为 `heading_level_1`，应为 `references_title`
4. **图题注 vs 表题注**：注意区分 `caption_figure` 和 `caption_table`
5. **致谢标题 vs 一级标题**："致谢" 可能被误识别为 `heading_level_1`，应为 `acknowledgements_title`
6. **摘要标题+内容合并**：如果摘要标题和正文在同一段落，应为 `abstract_chinese_title_content` 或 `abstract_english_title_content`，而非单独的标题或正文

## ⚠️ 重要：自动类型提升机制

**以下情况不是分类错误，不要修改 JSON 中的 category！**

WordFormat 内部有一个自动类型提升机制：当文档树构建完成后，代码会自动将特定父节点下的 `body_text` 子节点升级为对应的专用类型：

| 父节点类型 | 其下的 body_text 自动升级为 | 说明 |
|-----------|---------------------------|------|
| `abstract_chinese_title` | `abstract_chinese_content` | 中文摘要标题后面的正文段落 |
| `abstract_english_title` | `abstract_english_content` | 英文摘要标题后面的正文段落 |
| `references_title` | `reference_entry` | 参考文献标题后面的每条文献 |

**这意味着：**

- 摘要标题后面的正文段落，模型输出为 `body_text` 是**正确的**，代码会自动将其升级为 `abstract_chinese_content` 并应用摘要正文的格式规则
- 参考文献标题后面的每条文献，模型输出为 `body_text` 也是**正确的**，代码会自动升级为 `reference_entry` 并应用参考文献内容的格式规则
- **Agent 不需要手动修改这些 `body_text` 为其他类型**，修改反而会导致格式化逻辑混乱

**只有以下情况才需要手动修正 category：**
- 摘要标题本身被错分为 `heading_level_1` 或 `body_text`
- 关键词行被错分为 `body_text`
- 图注/表注互相错分
- 致谢标题被错分
