# Category 类型参考

AI 模型为每个段落分配的 category 类型。手动编辑 JSON 时修改此字段即可纠正分类。

## 模型输出的类别（22 种）

与 `src/wordformat/data/model/id2label.json` 一一对应，可直接出现在 `gj` 的 JSON 里：

| Category | 判定规则 |
|----------|----------|
| `body_text` | 普通正文段落 |
| `heading_level_1` | "第X章"或数字开头的标题 |
| `heading_level_2` | "X.Y" 格式的二级标题 |
| `heading_level_3` | "X.Y.Z" 格式的三级标题 |
| `heading_mulu` | "目录" 或 "目 录" |
| `heading_fulu` | "附录" |
| `document_title` | 论文题名（封面/摘要页重复出现，不参与格式化） |
| `abstract_chinese_title` | "摘要" |
| `abstract_chinese_title_content` | 摘要标题和正文在同一段落 |
| `abstract_chinese_content` | 独立成段的摘要正文 |
| `abstract_english_title` | "Abstract" |
| `abstract_english_title_content` | 英文摘要标题和正文在同一段落 |
| `abstract_english_content` | 独立成段的英文摘要正文 |
| `keywords_chinese` | 含"关键词"或"关键字" |
| `keywords_english` | 含"Keywords" |
| `caption_figure` | "图 X.Y" 开头的图注 |
| `caption_table` | "表 X.Y" 开头的表注 |
| `references_title` | "参考文献" 或 "References" |
| `references_content` | 参考文献条目（"[1] 作者. …"） |
| `acknowledgements_title` | "致谢" 或 "Acknowledgements" |
| `acknowledgements_content` | 致谢正文段落 |
| `other` | 封面、声明页等：不参与格式化，仅标记 |

## 后端追加的类别（非模型输出）

由结构识别或后处理产生，不应手动写进 JSON：

| Category | 来源 |
|----------|------|
| `figure_image` | 只含图片的段落（`gj` 识别） |
| `table_object` | 原文档表格条目（只读：`table_rows` + `exists: true`） |
| `math_block` | OMML 公式段（按 XML 结构判定，模型词表无此类） |
| `footer` | 页脚 / "由 AI 生成" 声明 |
| `toc_entry` | 目录条目（规则识别，不参与格式化） |
| `top` | 文档树根节点占位 |

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
| `references_title` | `references_content`（仅匹配 "[1]"/"1." 编号的条目） |
| `acknowledgements_title` | `acknowledgements_content` |

## replace 字段

在 JSON 条目中添加 `"replace": "新文本"` 可在格式化时替换段落内容。

```json
{
    "category": "abstract_chinese_title",
    "paragraph": "摘    要",
    "replace": "摘  要"
}
```