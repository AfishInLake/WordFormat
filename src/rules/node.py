#! /usr/bin/env python
# @Time    : 2026/1/10 14:07
# @Author  : afish
# @File    : node.py
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Dict, Generic, Optional, Type, TypeVar, get_args, get_origin

import yaml
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from pydantic import ValidationError

from src.config.datamodel import BaseModel


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
            from src.config.datamodel import NodeConfigRoot

            root_config = NodeConfigRoot(**raw_config)
            return root_config.model_dump()
        except FileNotFoundError as e:
            raise FileNotFoundError(f"配置文件 {config_path} 不存在") from e
        except ValidationError as e:
            raise ValueError(f"配置文件格式错误: {e}") from e
        except Exception as e:
            raise RuntimeError(f"加载配置失败: {e}") from e

    def load_config(self, full_config: dict | BaseModel) -> None:
        """重写父类方法：同时加载字典配置和Pydantic配置"""
        # 1. 先执行父类的字典配置加载（兼容旧逻辑）
        super().load_config(full_config)

        # 2. 基于父类加载的字典配置，转换为Pydantic模型对象
        # if not isinstance(full_config, dict) or not self.NODE_TYPE:
        #     raise ValueError(f"{self.NODE_TYPE} 缺少有效配置字典")

        if isinstance(full_config, BaseModel):
            # 获取当前 CONFIG_MODEL 的实际类型（去掉 Generic 包装）
            target_type = self.CONFIG_MODEL

            # 遍历 full_config 的所有字段
            for field_name, field_info in full_config.model_fields.items():
                field_annotation = field_info.annotation

                # 处理 Optional[T]、Union[T, None] 等情况
                origin = get_origin(field_annotation)
                args = get_args(field_annotation)

                # 提取真实的类型（去掉 Optional / Union[ ..., None]）
                if origin is not None:
                    # 如果是 Optional[T]，args 可能是 (T, type(None))
                    real_types = [t for t in args if t is not type(None)]
                    if len(real_types) == 1:
                        field_type = real_types[0]
                    else:
                        field_type = field_annotation  # 保守处理
                else:
                    field_type = field_annotation

                # 判断是否匹配：field_type 是 target_type 或其子类
                try:
                    if (
                        isinstance(field_type, type)
                        and issubclass(field_type, BaseModel)
                        and issubclass(field_type, target_type)
                    ):
                        # 找到了！获取该字段的值
                        field_value = getattr(full_config, field_name)
                        self._pydantic_config = field_value
                        return
                except TypeError:
                    # field_type 不是 class（比如是 ForwardRef），跳过
                    continue

            # 如果没找到，可以报错或 fallback
            raise ValueError(
                f"未能在 {type(full_config).__name__} 中找到类型为 {target_type.__name__} 的字段"
            )

        elif isinstance(full_config, dict):
            self._pydantic_config = self.CONFIG_MODEL(**full_config)
        else:
            raise TypeError(f"不支持的配置类型: {type(full_config)}")

    def update_paragraph(self, paragraph: Paragraph | dict):
        self.paragraph = paragraph

    def check_format(self, doc: Document):
        """虚方法：由子类实现具体的格式检查逻辑"""
        raise NotImplementedError("Subclasses should implement this!")

    def add_comment(self, doc: Document, runs: Run | Sequence[Run], text: str):
        if text.strip():
            doc.add_comment(runs=runs, text=text, author="论文解析器", initials="afish")
