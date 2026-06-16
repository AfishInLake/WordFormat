#! /usr/bin/env python
# @Time    : 2026/1/10 14:07
# @Author  : afish
# @File    : node.py
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Dict, Generic, Optional, Type, TypeVar

import yaml
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger
from pydantic import ValidationError

from wordformat.config.datamodel import BaseModel, NodeConfigRoot


class TreeNode:
    """树的节点类"""

    NODE_TYPE = "node"

    def __init__(self, value: Any):
        self.value = value
        self.__config = {}
        self.children: list[TreeNode] = []
        self.__set_fingerprint()

    @property
    def config(self):
        return self.__config

    def load_config(self, full_config: dict | Any) -> None:
        """
        根据 cls.NODE_TYPE（点分路径）从 full_config 中提取子配置。

        例如：
          NODE_TYPE = 'abstract.keywords.chinese'
          则从 full_config['abstract']['keywords']['chinese'] 提取
        """
        if not isinstance(full_config, dict):
            self.__config = {}
            return

        # 解析路径：支持 'a.b.c' 或直接 'top_level_key'
        path_parts = self.NODE_TYPE.split(".")

        current = full_config
        try:
            for part in path_parts:
                if not isinstance(current, dict):
                    raise KeyError(
                        f"Expected dict at path {'.'.join(path_parts[: path_parts.index(part) + 1])}"
                    )
                current = current[part]
            # 如果最终值不是 dict，也允许（但通常应是 dict）
            self.__config = current if isinstance(current, dict) else {}
        except (KeyError, TypeError):
            # 路径不存在或结构不匹配，返回空配置
            self.__config = {}

    def __set_fingerprint(self):
        self.fingerprint = None  # 初始化为 None
        if self.value and isinstance(self.value, dict):
            if self.value.get("category") == "top":
                return
            try:
                self.fingerprint = self.value["fingerprint"]
            except KeyError as err:
                raise ValueError(f"{self.value} must have a 'fingerprint' key") from err

    def add_child(self, child_value: Any) -> "TreeNode":
        """添加一个子节点，并返回该子节点（便于链式调用）"""
        child_node = TreeNode(child_value)
        self.children.append(child_node)
        return child_node

    def add_child_node(self, child_node: "TreeNode") -> None:
        """直接添加一个 TreeNode 作为子节点"""
        self.children.append(child_node)

    def __repr__(self) -> str:
        return f"TreeNode({self.value})"


T = TypeVar("T", bound=BaseModel)


class FormatNode(TreeNode, Generic[T]):
    """所有格式检查节点的基类"""

    CONFIG_MODEL: Type[T]

    # 子类声明：规则名 → handler 方法名，框架自动按 config 启用/禁用调度
    RULES: dict[str, str] = {}

    def __init__(
        self,
        value,
        level: int | float,
        paragraph: Paragraph = None,
        expected_rule: dict[str, Any] = None,
    ):
        super().__init__(value=value)  # value 就是 paragraph
        self.level: int | float = level
        self.paragraph: Paragraph = paragraph
        self.expected_rule = expected_rule
        self._pydantic_config: Optional[T] = None  # Pydantic配置对象

    @property
    def pydantic_config(self) -> T:
        """只读属性：获取类型安全的Pydantic配置对象"""
        if self._pydantic_config is None:
            raise ValueError(f"节点 {self.NODE_TYPE} 尚未加载Pydantic配置")
        return self._pydantic_config

    @classmethod
    def load_yaml_config(cls, config_path: str | Path) -> Dict[str, Any]:
        """加载并解析YAML配置文件"""
        try:
            with open(config_path, encoding="utf-8") as f:
                raw_config = yaml.safe_load(f)
            # 使用根模型验证整个配置结构
            from wordformat.config.datamodel import NodeConfigRoot

            root_config = NodeConfigRoot(**raw_config)
            return root_config.model_dump()
        except FileNotFoundError as e:
            raise FileNotFoundError(f"配置文件 {config_path} 不存在") from e
        except ValidationError as e:
            raise ValueError(f"配置文件格式错误: {e}") from e
        except Exception as e:
            raise RuntimeError(f"加载配置失败: {e}") from e

    def load_config(self, full_config: NodeConfigRoot):
        """重写父类方法：同时加载字典配置和Pydantic配置。

        通过 CONFIG_PATH 属性声明配置路径，沿路径逐级 getattr 解析。
        子类只需设置 CONFIG_PATH = "abstract.chinese.chinese_title" 即可，
        无需在此方法中维护硬编码映射表。
        """
        # 1. 先执行父类的字典配置加载（兼容旧逻辑）
        super().load_config(full_config)

        # 2. 有自定义 load_config 的子类（BaseHeadingNode、BaseKeywordsNode）
        #    会重写此方法，不会执行到这里
        config_path = getattr(self, "CONFIG_PATH", None)
        if config_path is None:
            return

        # 3. 沿 CONFIG_PATH 逐级 getattr
        obj = full_config
        for part in config_path.split("."):
            obj = getattr(obj, part)
        self._pydantic_config = obj

    def update_paragraph(self, paragraph: Paragraph | dict):
        self.paragraph = paragraph

    def _base(self, doc, p: bool, r: bool):
        raise NotImplementedError("Subclasses should implement this!")

    def _run_rules(self, doc: Document, p: bool) -> None:
        """自动调度业务规则：遍历 RULES，读配置，若 enabled 则执行 handler。

        p=True 表示检查模式，p=False 表示应用模式。handler 签名为
        (doc, rule_cfg, p)，p 默认 False 以兼容不需要区分模式的 handler。

        仅子类声明了 RULES 且有对应配置时才执行。提供双向验证：
        - 配置有规则但无 handler → warning
        - RULES 声明了但配置无对应项 → warning
        """
        if not self.RULES:
            return

        rules_config = getattr(self._pydantic_config, "rules", None)
        if rules_config is None:
            logger.warning(f"[{self.NODE_TYPE}] RULES 已声明但配置无 rules 节点")
            return

        declared_rules = set(self.RULES.keys())
        config_rules = set(type(rules_config).model_fields.keys())

        orphan_handlers = declared_rules - config_rules
        orphan_configs = config_rules - declared_rules
        if orphan_handlers:
            logger.warning(
                f"[{self.NODE_TYPE}] RULES 声明了 {orphan_handlers} 但配置无对应项"
            )
        if orphan_configs:
            logger.warning(
                f"[{self.NODE_TYPE}] 配置有 {orphan_configs} 但无对应 handler"
            )

        for rule_name, handler_name in self.RULES.items():
            rule_cfg = getattr(rules_config, rule_name, None)
            if rule_cfg is None or not rule_cfg.enabled:
                continue
            handler = getattr(self, handler_name)
            handler(doc, rule_cfg, p)

    def check_format(self, doc: Document):
        """格式检查：先执行样式检查，再自动调度业务规则"""
        self._base(doc, p=True, r=True)
        self._run_rules(doc, p=True)

    def apply_format(self, doc: Document):
        """格式应用：先清理、应用样式，再自动调度业务规则"""
        self._clean_paragraph_edge_spaces()
        self._base(doc, p=False, r=False)
        self._run_rules(doc, p=False)

    def apply_replace(self, doc: Document = None) -> bool:
        """替换段落文本内容（由 JSON 的 replace 字段驱动）。

        仅当 self.value 为 dict 且包含非空 "replace" 字段时执行替换。
        基类默认策略：多 run 时按原字符数分配，保持 run 边界语义。
        子类可覆写此方法实现特定类型的替换逻辑（如保留关键词标签 run、引用标记 run 等）。

        Returns:
            True 表示执行了替换，False 表示无需替换
        """
        value = self.value
        if not isinstance(value, dict):
            return False
        replace_text = value.get("replace")
        if not replace_text or not isinstance(replace_text, str):
            return False
        replace_text = replace_text.strip().replace(" ", "")
        if not replace_text:
            return False
        if self.paragraph is None:
            return False

        runs = self.paragraph.runs
        if not runs:
            return False

        if len(runs) == 1:
            runs[0].text = replace_text
        else:
            pos = 0
            for i, run in enumerate(runs):
                if i == len(runs) - 1:
                    run.text = replace_text[pos:]
                else:
                    n = min(len(run.text), len(replace_text) - pos)
                    run.text = replace_text[pos : pos + n] if n > 0 else ""
                    pos += n

        logger.debug(f"已替换段落文本 → {replace_text[:50]}...")
        return True

    def _clean_paragraph_edge_spaces(self) -> None:
        """清理段落首尾 run 中的多余空格。

        AI 生成的文档常在段落开头或结尾残留空格，此方法在格式化时自动清理：
        - 第一个非空 run 的开头空格
        - 最后一个非空 run 的结尾空格
        """
        if self.paragraph is None:
            return
        runs = self.paragraph.runs
        if not runs:
            return

        # 清理第一个非空 run 的开头空格
        for run in runs:
            if run.text:
                stripped = run.text.lstrip(" \u00a0")  # 普通空格 + 不间断空格
                if stripped != run.text:
                    run.text = stripped
                break

        # 清理最后一个非空 run 的结尾空格
        for run in reversed(runs):
            if run.text:
                stripped = run.text.rstrip(" \u00a0")
                if stripped != run.text:
                    run.text = stripped
                break

    def add_comment(self, doc: Document, runs: Run | Sequence[Run], text: str):
        if text.strip():
            doc.add_comment(runs=runs, text=text, author="论文解析器", initials="afish")
