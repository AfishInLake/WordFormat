"""节点注册装配点：框架层与业务节点的唯一交汇处。

分层约定：
- ``structure/`` 只提供注册机制（Registry）与配置解析，不 import 任何
  具体节点或领域模块；
- ``rules/`` 与 ``domains/*/rules/`` 只通过 ``@register`` 声明自己，不关心
  谁触发装配；
- 本模块负责把两者接起来：导入节点模块以触发 ``@register`` 副作用，
  并为"无 FormatNode 的后处理扩展标签"登记占位节点类。

``wordformat/__init__.py`` 在包初始化时调用 ``register_nodes()``，
因此任何 ``import wordformat.*`` 之后注册表都已就绪。
"""

from wordformat.categories import Category
from wordformat.domains.thesis.rules import (  # noqa: F401 — 论文域节点，触发 @register
    AbstractContentCN,
    AbstractContentEN,
    AbstractTitleCN,
    AbstractTitleContentCN,
    AbstractTitleContentEN,
    AbstractTitleEN,
    Acknowledgements,
    KeywordsCN,
    KeywordsEN,
    References,
)
from wordformat.rules import (  # noqa: F401 — 通用结构节点，触发 @register
    BodyText,
    CaptionFigure,
    CaptionTable,
    FigureImage,
    HeadingLevel1Node,
    HeadingLevel2Node,
    HeadingLevel3Node,
    TableObject,
)
from wordformat.structure.registry import _level_registry, _registry

_REGISTERED = False


def register_nodes() -> None:
    """登记占位类别（幂等）。

    真实节点类由上面的 import 副作用注册；这里只补齐后处理会产出、
    但没有专属 FormatNode 的类别，统一映射到 BodyText（再由
    settings.VOIDNODELIST 跳过格式化），避免 create_node 因未知类别
    返回 None 丢弃节点，导致段落与树节点错位。
    """
    global _REGISTERED
    if _REGISTERED:
        return
    _registry.setdefault(Category.OTHER, BodyText)  # 封面/声明等无需格式化
    _registry.setdefault(Category.DOCUMENT_TITLE, BodyText)  # 文档标题
    _registry.setdefault(Category.FOOTER, BodyText)  # 页脚/AI 生成声明
    _registry.setdefault(Category.HEADING_MULU, BodyText)  # 目录标题
    _registry.setdefault(Category.HEADING_FULU, BodyText)  # 附录标题
    _registry.setdefault(Category.TOC_ENTRY, BodyText)  # 目录条目
    # 无对应 FormatNode 的 terminal 类别：层级 1，子树按 skip_children 隔离
    _level_registry.setdefault(Category.HEADING_MULU, 1)
    _level_registry.setdefault(Category.HEADING_FULU, 1)
    _REGISTERED = True
