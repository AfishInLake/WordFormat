# Category 类型参考

AI 模型为每个段落分配的 category 类型。手动编辑 JSON 时修改此字段即可纠正分类。

## 完整列表

| Category | 判定规则 |
|----------|----------|
| `heading_level_1` | "第X章"或数字开头的标题 |
| `heading_level_2` | "X.Y" 格式的二级标题 |
| `heading_level_3` | "X.Y.Z" 格式的三级标题 |
| `heading_mulu` | "目录" 或 "目 录" |
| `heading_fulu` | "附录" |
| `abstract_chinese_title` | "摘要" |
| `abstract_chinese_title_content` | 摘要标题和正文在同一段落 |
| `abstract_english_title` | "Abstract" |
| `abstract_english_title_content` | 英文摘要标题和正文在同一段落 |
| `keywords_chinese` | 含"关键词"或"关键字" |
| `keywords_english` | 含"Keywords" |
| `body_text` | 普通正文段落 |
| `caption_figure` | "图 X.Y" 开头的图注 |
| `caption_table` | "表 X.Y" 开头的表注 |
| `references_title` | "参考文献" 或 "References" |
| `acknowledgements_title` | "致谢" 或 "Acknowledgements" |

## 常见误判修正

| 误识别 | 应改为 |
|--------|--------|
| "摘要" → `heading_level_1` | `abstract_chinese_title` |
| "关键词：..." → `body_text` | `keywords_chinese` |
| "参考文献" → `heading_level_1` | `references_title` |
| 目录/封面内容 | `other`（跳过格式化） |

## 自动提升机制

以下情况**不需要手动修改**，代码会自动处理：

| 父节点 | 其下 body_text 自动升级为 |
|--------|--------------------------|
| `abstract_chinese_title` | `abstract_chinese_content` |
| `abstract_english_title` | `abstract_english_content` |
| `references_title` | `reference_entry` |

## replace 字段

在 JSON 条目中添加 `"replace": "新文本"` 可在格式化时替换段落内容。

```json
{
    "category": "abstract_chinese_title",
    "paragraph": "摘    要",
    "replace": "摘  要"
}
```
