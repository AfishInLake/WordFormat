#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : _template.py
"""Preset 预设模板 —— 复制本文件后按需实现 register() 即可。

预设是 presets/ 目录下的 .py 插件，通过 CLI 的 --preset 参数加载：

    wordf cf -d 论文.docx -c config.yaml -f out.json -o output/ --preset my_preset
    wordf list-presets           # 查看可用预设（_template.py 不会列出）

文件约定：
1. 可选 PRESET_MANIFEST：声明预设需要的 YAML 配置项（兜底配置），
   加载时自动深合并进用户配置；用户提供的 -c 配置优先级最高。
2. 必须实现 register(hooks) 函数：在其中调用 hooks.register(event, callback)。

回调签名：``callback(event_data: dict) -> dict | None``
- 可直接修改 event_data 中的可变对象（config / root_node / issues 等）；
- 返回 dict 时键值会覆盖 event_data（如返回 {"text": "..."} 改写批注文案）。

覆盖框架机制：框架内置机制命名 ``builtin.*``（见 wordformat/handlers.py），
论文领域回调命名 ``thesis.*``（见 wordformat/domains/thesis.py）。预设可在
register() 中用 hooks.unregister(事件, 名字) 禁用后自行实现：

    hooks.unregister("on_tree_normalize", "thesis.tree_normalize")
    hooks.unregister("before_document_save", "builtin.post_process")
    hooks.unregister("before_document_save", "builtin.header_footer")
    hooks.unregister("on_summary_build", "thesis.summary_build")
    hooks.unregister("before_node_format", "thesis.caption_numbering")
"""

# ---- 可选：预设兜底配置（YAML 结构，与 -c 配置同一格式）----
# 仅当用户配置中缺省时生效；多个预设叠加时后加载的覆盖先加载的。
PRESET_MANIFEST = {
    "template_name": "示例模板（preset）",
}

# ---- 可选：预设说明（list-presets 暂不展示，供人阅读）----
PRESET_DESCRIPTION = "Preset 模板：演示全部钩子点的注册方式。"


def register(hooks):
    """注册本预设需要的全部 Hook 回调（入口函数，不要改名）。"""

    # （可选）禁用某个框架/领域机制，改用预设自己的实现：
    # hooks.unregister("on_summary_build", "thesis.summary_build")
    # hooks.unregister("on_tree_normalize", "thesis.tree_normalize")
    # hooks.unregister("before_node_format", "thesis.caption_numbering")

    # 1) on_config_loaded —— 配置加载完成后
    #    event_data: config（NodeConfigRoot，可直接修改/补充）
    def on_config_loaded(data):
        config = data["config"]
        if isinstance(config, dict):
            config.setdefault("template_name", "示例模板（preset 注入）")

    # 2) on_tree_built —— 文档树构建完成后
    #    event_data: root_node（FormatNode 根节点）、config
    def on_tree_built(data):
        root = data["root_node"]
        # 示例：给根节点打个标记（自定义节点分类的入口）
        root.custom_flag = getattr(root, "custom_flag", 0) + 1

    # 3) before_node_format / 4) after_node_format —— 每个节点格式化前/后
    #    event_data: node、paragraph、config、check
    def before_node_format(data):
        node = data["node"]
        paragraph = data["paragraph"]
        # 示例：对特定节点类型预处理段落内容
        cls_name = type(node).__name__
        if cls_name == "HeadingLevel1Node" and paragraph is not None:
            pass  # 在这里改写 paragraph runs

    def after_node_format(data):
        node = data["node"]
        # 示例：记录已格式化节点（追加自定义格式规则的落点）
        pass

    # 5) on_caption_check —— 题注校验时（仅 check 模式）
    #    event_data: node、parsed（parse_caption_text 结果）、issues（list[str]）
    #    追加到 issues 即可增加自定义题注规则；返回 {"issues": [...]} 可整体覆写
    def on_caption_check(data):
        issues = data["issues"]
        parsed = data["parsed"]
        # 示例：题注文字带句号提醒
        if parsed and str(parsed.get("name", "")).endswith("。"):
            issues.append("图注-标点问题：题注末尾不应带句号，规范：去掉句号")

    # 6) on_heading_numbering —— 每个标题自动编号时
    #    event_data: heading_node、paragraph、numbering_config、level_key、num_id
    #    返回 {"skip": True} 完全接管编号；返回 {"num_id": "自定义"} 覆盖编号定义
    def on_heading_numbering(data):
        return None  # 默认走框架编号，自定义示例：
        # return {"skip": True}  # 完全由预设处理编号

    # 7) before_document_save —— 文档保存前
    #    event_data: document、config
    #    可插入声明页、封面等（插入后如需格式化需自行处理）
    def before_document_save(data):
        document = data["document"]
        # 示例：在文档末尾追加一段说明
        # document.add_paragraph("（由 preset 追加）")

    # 8) on_comment —— 添加批注时
    #    event_data: paragraph、text（原批注文案）、doc
    #    返回 {"text": "新文案"} 可改写批注文案/级别
    def on_comment(data):
        return None  # 默认原样保留
        # return {"text": data["text"] + "\n（preset 补充说明）"}

    hooks.register("on_config_loaded", on_config_loaded)
    hooks.register("on_tree_built", on_tree_built)
    hooks.register("before_node_format", before_node_format)
    hooks.register("after_node_format", after_node_format)
    hooks.register("on_caption_check", on_caption_check)
    hooks.register("on_heading_numbering", on_heading_numbering)
    hooks.register("before_document_save", before_document_save)
    hooks.register("on_comment", on_comment)
