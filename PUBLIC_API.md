# WordFormat 公共 API 手册

## 目录

- [1. 核心数据层 (`core`)](#1-核心数据层-core)
  - [1.1 TreeNode —— 树节点](#11-treenode--树节点)
  - [1.2 Tree / Stack —— 树与栈](#12-tree--stack--树与栈)
  - [1.3 build_tree —— 统一建树](#13-build_tree--统一建树)
  - [1.4 CategoryRegistry —— 可配置类别注册表](#14-categoryregistry--可配置类别注册表)
- [2. 规则层 (`rules`)](#2-规则层-rules)
  - [2.1 FormatNode —— 格式节点基类](#21-formatnode--格式节点基类)
- [3. 编排层 (`orchestration`)](#3-编排层-orchestration)
  - [3.1 bind_and_sync —— 段落绑定同步](#31-bind_and_sync--段落绑定同步)
  - [3.2 fix_all_style_definitions —— 样式定义修正](#32-fix_all_style_definitions--样式定义修正)
  - [3.3 format_table_content —— 表格格式化](#33-format_table_content--表格格式化)
- [4. 配置层 (`config`)](#4-配置层-config)
  - [4.1 NodeConfigRoot —— 配置模型](#41-nodeconfigroot--配置模型)
  - [4.2 init_config / get_config —— 配置加载](#42-init_config--get_config--配置加载)
- [5. 入口函数](#5-入口函数)
  - [5.1 auto_format_thesis_document](#51-auto_format_thesis_document)
  - [5.2 set_tag_main](#52-set_tag_main)

---

## 1. 核心数据层 (`core`)

`core` 包是 WordFormat 的基础设施层，**零外部依赖**（仅需 Python 标准库）。可独立安装使用，不强制引入 `python-docx`、`pydantic` 等重型依赖。

### 1.1 TreeNode —— 树节点

纯数据结构，表示树中的一个节点。支持层级 JSON 序列化。

```python
from wordformat import TreeNode

# 创建节点
root = TreeNode({"category": "heading_level_1", "paragraph": "第一章 绪论"})
child = TreeNode({"category": "body_text", "paragraph": "这是正文内容。"})
root.add_child_node(child)

# 导出为层级 JSON
hierarchical_json = root.to_dict()
# {
#   "value": {"category": "heading_level_1", "paragraph": "第一章 绪论"},
#   "children": [
#     {"value": {"category": "body_text", "paragraph": "这是正文内容。"}, "children": []}
#   ]
# }

# 从层级 JSON 还原
import json
restored = TreeNode.from_dict(json.loads(json.dumps(hierarchical_json)))
assert restored.children[0].value["paragraph"] == "这是正文内容。"
```

**使用场景**：
- 构建文档语义树，无需依赖 docx
- 将树结构导出为 JSON 供外部系统消费
- 在内存中操作文档结构后进行序列化/反序列化

---

### 1.2 Tree / Stack —— 树与栈

`Tree` 封装了多叉树的遍历和查找。`Stack` 是泛型 LIFO 栈。

```python
from wordformat import Tree, Stack

# Tree —— 多叉树
tree = Tree({"category": "top"})
tree.root.add_child_node(TreeNode({"category": "heading_level_1"}))
tree.root.add_child_node(TreeNode({"category": "heading_level_2"}))

# 前序遍历
for value in tree.preorder():
    print(value["category"])  # top → heading_level_1 → heading_level_2

# 按条件查找（DFS）
found = tree.find_by_condition(lambda v: v.get("category") == "heading_level_1")
print(found.value)  # {"category": "heading_level_1"}

print(tree.height())  # 树高度
print(tree.size())    # 节点总数

# Stack —— 泛型栈
stack = Stack[int]()
stack.push(1)
stack.push(2)
print(stack.pop())       # 2
print(stack.peek_safe()) # 1
```

**使用场景**：
- 遍历文档树提取特定类型节点
- 在建树/遍历算法中维护层级状态

---

### 1.3 build_tree —— 统一建树

从扁平段落列表构建层级树，支持自定义类别注册表和节点工厂。

```python
from wordformat import build_tree, TreeNode

# 扁平段落列表（模拟 AI 分类结果）
items = [
    {"category": "heading_level_1", "paragraph": "第一章"},
    {"category": "body_text", "paragraph": "正文段落1"},
    {"category": "heading_level_2", "paragraph": "1.1 小节"},
    {"category": "body_text", "paragraph": "正文段落2"},
    {"category": "heading_level_1", "paragraph": "第二章"},
    {"category": "body_text", "paragraph": "正文段落3"},
]

# 使用默认类别注册表建树
root = build_tree(items)

# 验证树结构
print(f"一级标题数: {len(root.children)}")  # 2
print(f"第一章的子节点数: {len(root.children[0].children)}")  # 2（正文 + 1.1小节）

# 导出查看完整结构
print(root.to_dict())
```

**使用场景**：
- AI 分类后的扁平 JSON 转为层级文档结构
- 自定义类别扩展后建树
- 作为 `wordf tree` 命令的后端

---

### 1.4 CategoryRegistry —— 可配置类别注册表

管理段落类别名、逻辑层级、是否标题等元信息。支持运行时扩展。

```python
from wordformat import get_registry, CategoryRegistry, build_tree

# 获取全局注册表
registry = get_registry()

# 查询内置类别
print(registry.get_level("heading_level_1"))   # 1
print(registry.is_heading("heading_level_1"))  # True
print(registry.get_level("body_text"))         # 999
print(registry.is_heading("body_text"))        # False

# === 场景 1: 注册自定义类别 ===
registry.register(
    name="custom_abstract",
    level=1,
    is_heading=True,
    override=False,  # 已存在时报错
)

# === 场景 2: 批量扩展 ===
registry.update({
    "categories": {
        "appendix_title": {"level": 1, "is_heading": True},
        "appendix_content": {"level": 999, "is_heading": False},
        "author_info": {"level": 999, "is_heading": False},
    },
    "extend_defaults": True,  # 合并到默认值
})

# 现在可以用新类别建树
items = [
    {"category": "appendix_title", "paragraph": "附录A"},
    {"category": "appendix_content", "paragraph": "附录正文..."},
]
root = build_tree(items)
print(root.to_dict())
```

**使用场景**：
- 论文模板新增段落类型（如"附录""作者简介""基金信息"）
- 不同学校/期刊的论文格式规范扩展
- 无需修改源码即可适配新的文档结构

---

## 2. 规则层 (`rules`)

### 2.1 FormatNode —— 格式节点基类

扩展 `TreeNode`，增加格式化能力：配置加载、样式校验、样式应用。

```python
from wordformat import FormatNode, NodeConfigRoot, init_config, get_config

# 1. 定义节点子类（声明 CONFIG_PATH 和 CONFIG_MODEL）
from wordformat.config.datamodel import GlobalFormatConfig

class MyBodyText(FormatNode[GlobalFormatConfig]):
    NODE_TYPE = "body_text"
    CONFIG_PATH = "body_text"               # 配置路径（getattr 链）
    CONFIG_MODEL = GlobalFormatConfig        # Pydantic 配置模型

# 2. 加载 YAML 配置
init_config("presets/undergrad.yaml")
config = get_config()

# 3. 创建节点并加载配置
node = MyBodyText(
    value={"category": "body_text", "paragraph": "正文内容"},
    level=999,
)
node.load_config(config)

# 4. 访问类型安全的配置
cfg = node.pydantic_config
print(cfg.font_size)           # "小四"
print(cfg.chinese_font_name)   # "宋体"
print(cfg.first_line_indent)   # "2字符"
```

**使用场景**：
- 创建自定义段落类型的格式规则
- 扩展新的 FormatNode 子类处理特殊格式需求
- 读取 Pydantic 类型安全的配置对象

**内置子类**（可直接使用）：

| 类名 | 对应段落类型 |
|------|-------------|
| `HeadingLevel1Node` | 一级标题 |
| `HeadingLevel2Node` | 二级标题 |
| `HeadingLevel3Node` | 三级标题 |
| `AbstractTitleCN` | 中文摘要标题 |
| `AbstractContentCN` | 中文摘要内容 |
| `AbstractTitleEN` | 英文摘要标题 |
| `AbstractContentEN` | 英文摘要内容 |
| `KeywordsCN` | 中文关键词 |
| `KeywordsEN` | 英文关键词 |
| `BodyText` | 正文 |
| `CaptionFigure` | 图题注 |
| `CaptionTable` | 表题注 |
| `References` | 参考文献标题 |
| `ReferenceEntry` | 参考文献条目 |

---

## 3. 编排层 (`orchestration`)

编排层引入 `python-docx` 依赖，负责将虚拟节点树的操作同步到真实 Word 文档。

### 3.1 bind_and_sync —— 段落绑定同步

通过文本序列对齐，将虚拟节点树的每个节点与 docx 段落一一绑定。支持插入和删除。

```python
from wordformat import bind_and_sync, build_tree, TreeNode
from docx import Document

# 1. 准备虚拟节点树
items = [
    {"category": "heading_level_1", "paragraph": "第一章"},
    {"category": "body_text", "paragraph": "正文内容"},
    {"category": "body_text", "paragraph": "新增段落"},  # ← Word 中不存在
]
root = build_tree(items)

# 2. 加载 Word 文档
doc = Document("thesis.docx")

# 3. 绑定同步（check=True 仅报告差异，不修改文档）
bind_and_sync(root, doc, check=True)
# [WARNING] 1 个 JSON 条目在文档中找不到对应段落（需插入）

# 4. 实际同步（check=False 执行插入/删除）
bind_and_sync(root, doc, check=False)
# [INFO] 已插入段落 #2: 新增段落

# 5. 保存
doc.save("synced.docx")
```

**使用场景**：
- JSON 编辑后同步到 Word 文档（插入缺失段落、删除多余段落）
- 批量文档内容校正
- 检查模式下预览差异而不实际修改

---

### 3.2 fix_all_style_definitions —— 样式定义修正

在格式化前统一修正 Word 样式定义中的字符和段落属性，确保样式定义与 YAML 配置一致。

```python
from wordformat import fix_all_style_definitions, init_config, get_config
from docx import Document

# 加载配置
init_config("presets/undergrad.yaml")
config = get_config()

# 加载文档
doc = Document("thesis.docx")

# 修正所有样式定义（清除主题色、修正字体/字号/加粗/斜体/下划线/对齐/间距等）
fix_all_style_definitions(doc, config)

# 现在文档中的 "正文"、"Heading 1" 等样式定义已与配置一致
doc.save("style_fixed.docx")
```

**内部修正内容**：
- 字体名称（中英文分别设置）
- 字号（小四 → 12pt → w:sz 24）
- 颜色（清除 Office 主题色，改为纯色值）
- 加粗 / 斜体 / 下划线
- 对齐方式
- 段前/段后间距、行距、首行缩进、左右缩进

**使用场景**：
- 批量规范化模板中的样式定义
- 清除 Word 主题色避免打印/导出时的颜色偏差
- 确保论文正文、标题、题注等样式与学校要求一致

---

### 3.3 format_table_content —— 表格格式化

遍历文档中所有表格的单元格，对单元格内段落进行格式校验或应用。

```python
from wordformat import format_table_content, init_config, get_config
from docx import Document

init_config("presets/undergrad.yaml")
config = get_config()
doc = Document("thesis.docx")

# 检查模式：仅 diff，不修改
format_table_content(doc, config, check=True)

# 应用模式：修改表格单元格格式
format_table_content(doc, config, check=False)

doc.save("table_formatted.docx")
```

**使用场景**：
- 表格内文字统一为指定字体和字号
- 表格单元格段落格式化（与正文格式要求不同时）
- 检查模式下审核表格格式合规性

---

## 4. 配置层 (`config`)

### 4.1 NodeConfigRoot —— 配置模型

Pydantic v2 模型，定义论文格式规范的完整配置结构。

```python
from wordformat import NodeConfigRoot

# 使用默认配置
config = NodeConfigRoot()
print(config.body_text.font_size)          # "小四"
print(config.body_text.chinese_font_name)  # "宋体"

# 编程方式构建配置
from wordformat.config.datamodel import (
    GlobalFormatConfig, HeadingsConfig, HeadingLevelConfig,
    AbstractConfig, AbstractChineseConfig,
)

custom = NodeConfigRoot(
    body_text=GlobalFormatConfig(
        chinese_font_name="宋体",
        english_font_name="Times New Roman",
        font_size="小四",
        first_line_indent="2字符",
        line_spacingrule="1.5倍行距",
        line_spacing="1.5倍",
    ),
    headings=HeadingsConfig(
        level_1=HeadingLevelConfig(
            chinese_font_name="黑体",
            font_size="小二",
            bold=True,
            alignment="居中对齐",
        ),
    ),
)

# 导出为 dict（用于生成 YAML 模板）
config_dict = custom.model_dump()
```

**使用场景**：
- 程序化生成配置文件模板
- 在内存中构建配置而不依赖 YAML 文件
- 配置校验和类型安全检查

---

### 4.2 init_config / get_config —— 配置加载

懒加载单例，管理 YAML 配置文件的加载和缓存。

```python
from wordformat import init_config, get_config

# 初始化配置路径（仅记录路径，不加载）
init_config("presets/undergrad.yaml")

# 首次调用时加载并缓存，后续调用直接返回缓存
config = get_config()

# 访问配置
print(config.body_text.chinese_font_name)  # "宋体"
print(config.headings.level_1.font_size)   # "小二"
print(config.abstract.chinese.chinese_title.bold)  # True
```

**使用场景**：
- 全局配置单例，避免重复加载 YAML
- 多个模块共享同一份配置
- 测试中用 `clear_config()` 重置

---

## 5. 入口函数

### 5.1 auto_format_thesis_document

完整的论文格式化流水线入口。

```python
from wordformat import auto_format_thesis_document

# 检查模式（仅添加批注，不修改内容）
result_path = auto_format_thesis_document(
    jsonpath="classified_paragraphs.json",  # AI 分类结果
    docxpath="draft.docx",                  # 原始论文
    configpath="presets/undergrad.yaml",    # 格式规范
    savepath="output/",                     # 输出目录
    check=True,                             # 仅检查
)
# → output/draft--标注版.docx

# 应用模式（实际修改格式）
result_path = auto_format_thesis_document(
    jsonpath="classified_paragraphs.json",
    docxpath="draft.docx",
    configpath="presets/undergrad.yaml",
    savepath="output/",
    check=False,                            # 实际修改
)
# → output/draft--修改版.docx
```

**内部流程**：

```
JSON → 建树 → 段落绑定 → 节点提升 → 样式修正 → 格式校验/应用 → 表格格式化 → 编号处理 → 超链接 → 保存
```

**使用场景**：
- 一键完成论文格式检查和修正
- CI/CD 流水线中自动化格式审查
- Web API 后端（`wordf startapi`）

---

### 5.2 set_tag_main

AI 分类入口：对 Word 文档段落进行 ONNX 模型推理，输出分类 JSON。

```python
from wordformat import set_tag_main

# 生成分类 JSON
result = set_tag_main("draft.docx", configpath="presets/undergrad.yaml")
# result 是 list[dict]，每项为：
# {
#     "category": "heading_level_1",
#     "score": 0.9823,
#     "paragraph": "第一章 绪论",
#     "original_text": "第一章 绪论",
#     "index": 0,
#     "fingerprint": "abc123...",
# }

# 保存为 JSON 文件
import json
with open("classified.json", "w") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
```

**使用场景**：
- CLI `wordf gj` 命令的后端
- 批量文档分类
- 生成 JSON 后手动编辑分类结果，再执行格式化

---

## 快速参考

```python
import wordformat as wf

# ── 纯数据操作（不依赖 docx） ──────────────────────
root = wf.build_tree(items)                    # 扁平 → 层级树
json_data = root.to_dict()                     # 树 → 层级 JSON
back = wf.TreeNode.from_dict(json_data)        # JSON → 树
wf.get_registry().register("my_type", level=1, is_heading=True)

# ── 配置 ──────────────────────────────────────────
wf.init_config("presets/undergrad.yaml")
cfg = wf.get_config()
print(cfg.body_text.font_size)

# ── 文档处理 ──────────────────────────────────────
wf.fix_all_style_definitions(doc, cfg)         # 修正样式定义
wf.bind_and_sync(root, doc, check=False)       # 同步段落
wf.format_table_content(doc, cfg, check=False) # 格式化表格

# ── 一键流水线 ────────────────────────────────────
wf.auto_format_thesis_document("data.json", "draft.docx", "presets/undergrad.yaml")
```
