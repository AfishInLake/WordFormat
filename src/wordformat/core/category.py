#! /usr/bin/env python
"""段落类别常量与层级映射 —— 纯数据，零外部依赖。

CategoryRegistry 是可配置的类别注册表：
  - 内置默认类别
  - 支持 register() 扩展新类别
  - 支持 update() 批量覆盖默认值
  - 模块级常量（CATEGORY_NAMES 等）从默认注册表派生，保持向后兼容
"""


class CategoryRegistry:
    """段落类别注册表。

    管理三类信息：
      - level_map:  类别名 → 逻辑层级（999 = 正文叶子）
      - heading_categories: 标题类别的集合（建树时触发栈操作）
      - body_default_level: 未被显式注册的类别默认层级
    """

    BODY_DEFAULT_LEVEL = 999

    def __init__(self):
        self._level_map: dict[str, int] = {}
        self._heading_categories: set[str] = set()
        self._load_defaults()

    # ---- 默认值 ------------------------------------------------------------

    def _load_defaults(self):
        """加载内置默认类别。"""
        defaults = {
            "heading_level_1": (1, True),
            "heading_level_2": (2, True),
            "heading_level_3": (3, True),
            "heading_mulu": (1, True),
            "heading_fulu": (1, True),
            "references_title": (1, True),
            "acknowledgements_title": (1, True),
            "abstract_chinese_title": (1, True),
            "abstract_english_title": (1, True),
            "abstract_chinese_title_content": (1, True),
            "abstract_english_title_content": (1, True),
            "abstract_chinese_content": (999, False),
            "abstract_english_content": (999, False),
            "keywords_chinese": (3, False),
            "keywords_english": (3, False),
            "caption_figure": (999, False),
            "caption_table": (999, False),
            "body_text": (999, False),
            "image": (999, False),
            "table": (999, False),
            "formula": (999, False),
        }
        for name, (level, is_heading) in defaults.items():
            self._level_map[name] = level
            if is_heading:
                self._heading_categories.add(name)

    # ---- 查询 --------------------------------------------------------------

    def get_level(self, category: str) -> int:
        """返回类别的逻辑层级，未注册的类别返回 BODY_DEFAULT_LEVEL。"""
        return self._level_map.get(category, self.BODY_DEFAULT_LEVEL)

    def is_heading(self, category: str) -> bool:
        """判断类别是否为标题（建树时触发栈操作）。"""
        return category in self._heading_categories

    @property
    def level_map(self) -> dict[str, int]:
        """返回 level_map 的只读副本。"""
        return dict(self._level_map)

    @property
    def heading_categories(self) -> frozenset[str]:
        """返回 heading_categories 的只读副本。"""
        return frozenset(self._heading_categories)

    @property
    def all_categories(self) -> frozenset[str]:
        """返回所有已注册类别名。"""
        return frozenset(self._level_map.keys())

    # ---- 扩展 --------------------------------------------------------------

    def register(
        self,
        name: str,
        level: int,
        is_heading: bool = False,
        override: bool = False,
    ) -> None:
        """注册一个新类别（或更新已有类别）。

        Args:
            name: 类别名（如 "custom_abstract"）
            level: 逻辑层级（1-3 = 标题，999 = 正文）
            is_heading: 是否作为标题类别参与建树栈
            override: 为 True 时允许覆盖已有类别；默认 False 时对已存在类别报错
        """
        if name in self._level_map and not override:
            raise ValueError(f"类别 '{name}' 已存在，使用 override=True 强制覆盖")
        self._level_map[name] = level
        if is_heading:
            self._heading_categories.add(name)
        else:
            self._heading_categories.discard(name)

    def update(self, config: dict) -> None:
        """从配置 dict 批量更新类别映射。

        config 格式：
          {
            "categories": {
              "my_type": {"level": 2, "is_heading": true},
              ...
            },
            "extend_defaults": true   # true=合并, false=替换
          }

        若 extend_defaults=False，先清空内置默认值再加载。
        """
        categories = config.get("categories", {})
        if not config.get("extend_defaults", True):
            self._level_map.clear()
            self._heading_categories.clear()
        for name, spec in categories.items():
            self.register(
                name=name,
                level=spec.get("level", self.BODY_DEFAULT_LEVEL),
                is_heading=spec.get("is_heading", False),
                override=True,
            )


# ---- 模块级默认实例 ---------------------------------------------------------

_registry = CategoryRegistry()

# 向后兼容的模块级常量（从默认注册表派生）
CATEGORY_NAMES: frozenset[str] = _registry.all_categories
HEADING_CATEGORIES: frozenset[str] = _registry.heading_categories
LEVEL_MAP: dict[str, int] = _registry.level_map


def get_registry() -> CategoryRegistry:
    """返回模块级默认注册表实例。"""
    return _registry


def reset_registry() -> CategoryRegistry:
    """重置注册表为内置默认值（测试用）。"""
    global _registry, CATEGORY_NAMES, HEADING_CATEGORIES, LEVEL_MAP
    _registry = CategoryRegistry()
    CATEGORY_NAMES = _registry.all_categories
    HEADING_CATEGORIES = _registry.heading_categories
    LEVEL_MAP = _registry.level_map
    return _registry
