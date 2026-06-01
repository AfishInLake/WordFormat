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
from wordformat.core.tree import TreeNode as _CoreTreeNode


class TreeNode(_CoreTreeNode):
    """树的节点类 —— 继承 core.TreeNode，增加指纹和配置支持。"""

    NODE_TYPE = "node"

    def __init__(self, value: Any):
        super().__init__(value=value)
        self.__config = {}
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
        self.fingerprint = None
        self.index = None
        if self.value and isinstance(self.value, dict):
            if self.value.get("category") == "top":
                return
            meta = (
                self.value.get("meta", {})
                if isinstance(self.value.get("meta"), dict)
                else {}
            )
            self.fingerprint = meta.get("fingerprint") or self.value.get("fingerprint")
            self.index = (
                meta.get("index")
                if meta.get("index") is not None
                else self.value.get("index")
            )

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
        self._is_insertion: bool = False
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
        """统一的配置加载：字典路径（TreeNode） + Pydantic getattr/item 链。

        子类只需设置 CONFIG_PATH = "abstract.chinese.keywords" 即可，
        无需覆写此方法或维护硬编码映射表。
        """
        # 1. 字典配置（TreeNode.load_config 处理，兼容 dict 输入）
        super().load_config(full_config)

        config_path = getattr(self, "CONFIG_PATH", None)
        if config_path is None:
            return

        # 2. Pydantic 配置：NodeConfigRoot → getattr/item 链
        if isinstance(full_config, NodeConfigRoot):
            obj = full_config
            for part in config_path.split("."):
                obj = obj[part] if isinstance(obj, dict) else getattr(obj, part)
            self._pydantic_config = obj
            return

        # 3. dict 输入兼容：从字典配置重建 Pydantic 模型
        if isinstance(full_config, dict) and self.config:
            try:
                self._pydantic_config = self.CONFIG_MODEL(**self.config)
            except Exception:
                pass

    def update_paragraph(self, paragraph: Paragraph | dict):
        self.paragraph = paragraph

    def _base(self, doc, p: bool, r: bool):
        raise NotImplementedError("Subclasses should implement this!")

    def check_format(self, doc: Document):
        """虚方法：由子类实现具体的格式检查逻辑"""
        self._base(doc, p=True, r=True)

    def apply_format(self, doc: Document):
        """虚方法：由子类实现具体的格式应用逻辑"""
        self._clean_paragraph_edge_spaces()
        self._base(doc, p=False, r=False)

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

    # ── 声明式内容接口 ──────────────────────────────────────────────
    # 子类可覆写以提供类型特定的 extract / render / patch 行为。
    # 默认实现以遗留方式工作，保持向后兼容。

    def extract(self, paragraph: Paragraph) -> dict:  # noqa: B027
        """从真实 docx 段落读取内容状态，返回 dict 合并到 self.content。"""
        return {"text": paragraph.text or ""}

    def render(self, document: Document) -> Any:  # noqa: B027
        """从虚拟状态创建 OOXML 元素。默认创建一个简单 w:p。"""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        new_p = OxmlElement("w:p")
        text = self.content.get("text", self.value.get("paragraph", ""))
        if text and text.strip():
            new_r = OxmlElement("w:r")
            new_t = OxmlElement("w:t")
            new_t.text = text
            new_t.set(qn("xml:space"), "preserve")
            new_r.append(new_t)
            new_p.append(new_r)
        return new_p

    def patch(self, paragraph: Paragraph, document: Document) -> list[str]:  # noqa: B027
        """比较虚拟内容与真实段落，返回变更描述列表。"""
        return []

    def get_alignment_text(self) -> str:
        """返回用于序列对齐的内容摘要（稳定，不依赖 XML 指纹）。"""
        text = self.content.get("text", "")
        if not text and isinstance(self.value, dict):
            text = self.value.get("paragraph", "")
        return text.strip()

    def to_value_dict(self) -> dict:
        """序列化为 JSON 兼容 dict，排除二进制数据。"""
        d = (
            dict(self.value)
            if isinstance(self.value, dict)
            else {"category": self.type}
        )
        d.update(self.content)
        return d
