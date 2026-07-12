#! /usr/bin/env python
# @Time    : 2026/1/10 14:07
# @Author  : afish
# @File    : node.py
from collections.abc import Sequence
from typing import Any

from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger

from wordformat.config.dotdict import DotDict, deep_merge


class TreeNode:
    """树的节点类"""

    NODE_TYPE = "node"

    def __init__(self, value: Any):
        self.value = value
        self._config = DotDict()
        self.children: list[TreeNode] = []
        self.fingerprint = None

    @property
    def config(self):
        return self._config

    def load_config(self, full_config: dict) -> None:
        """根据 NODE_TYPE 从 full_config 中提取子配置，与 DEFAULTS 合并。"""
        # 沿 NODE_TYPE 路径逐级查找 YAML 配置
        path_parts = self.NODE_TYPE.split(".")
        yaml_node: dict = {}
        current = full_config
        try:
            for part in path_parts:
                if not isinstance(current, dict):
                    raise KeyError
                current = current[part]
            yaml_node = current if isinstance(current, dict) else {}
        except (KeyError, TypeError):
            yaml_node = {}

        # 合并：DEFAULTS 为底，YAML 覆盖
        defaults = getattr(type(self), "DEFAULTS", {})
        merged = deep_merge(defaults, yaml_node) if defaults else yaml_node
        self._config = DotDict(merged)

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


class FormatNode(TreeNode):
    """所有格式检查节点的基类"""

    # 子类定义的默认值，load_config 时与 YAML 合并
    DEFAULTS: dict = {}

    # 全局错误统计（类变量，一次 check 周期内累加）
    _error_stats: dict[str, int] = {"total": 0, "错误": 0, "提醒": 0}

    # 中文标签，用作批注的 [位置] 部分
    NODE_LABEL: str = ""

    # 子类声明：规则名 → handler 方法名，框架自动按 config 启用/禁用调度
    RULES: dict[str, str] = {}

    # 框架自动注入的默认规则，不依赖配置，所有节点默认启用
    DEFAULT_RULES: dict[str, str] = {
        "paragraph_style": "_handle_paragraph_style",
        "character_style": "_handle_character_style",
    }

    def __init__(
        self,
        value,
        level: int | float,
        paragraph: Paragraph = None,
        expected_rule: dict[str, Any] = None,
    ):
        super().__init__(value=value)
        self.level: int | float = level
        self.paragraph: Paragraph = paragraph
        self.expected_rule = expected_rule
        self._comment_texts: list[tuple] = []

    @property
    def pydantic_config(self) -> "DotDict":
        """返回当前节点的合并配置（DotDict）。"""
        return self._config

    def update_paragraph(self, paragraph: Paragraph | dict):
        self.paragraph = paragraph

    def _base(self, doc, p: bool, r: bool):
        """子类可覆写以添加自定义逻辑（如拆分混合 run）。标准节点无需覆写。"""

    def _run_rules(self, doc: Document, p: bool) -> None:
        """自动调度所有规则：DEFAULT_RULES（无需配置）+ RULES（需 YAML 配置）。

        p=True 表示检查模式，p=False 表示应用模式。handler 签名为
        (doc, rule_cfg, p)，p 默认 False 以兼容不需要区分模式的 handler。

        DEFAULT_RULES 总是执行，RULES 仅当配置 enabled=true 时执行。提供双向验证：
        - 配置有规则但无 handler → warning
        - RULES 声明了但配置无对应项 → warning
        """
        all_rules = {**self.DEFAULT_RULES, **self.RULES}
        if not all_rules:
            return

        rules_config = getattr(self.pydantic_config, "rules", None)

        # 双向验证（仅检查自定义 RULES）
        if self.RULES:
            if rules_config is None:
                logger.debug(f"[{self.NODE_TYPE}] RULES 已声明但配置无 rules 节点")
            else:
                declared_rules = set(self.RULES.keys())
                config_rules = (
                    set(rules_config.keys())
                    if isinstance(rules_config, dict)
                    else set()
                )
                orphan_handlers = declared_rules - config_rules
                orphan_configs = config_rules - declared_rules
                if orphan_handlers:
                    logger.debug(
                        f"[{self.NODE_TYPE}] RULES 声明了 {orphan_handlers} 但配置无对应项"
                    )
                if orphan_configs:
                    logger.debug(
                        f"[{self.NODE_TYPE}] 配置有 {orphan_configs} 但无对应 handler"
                    )

        for rule_name, handler_name in all_rules.items():
            # 默认规则无配置，总是执行
            if rule_name in self.DEFAULT_RULES:
                handler = getattr(self, handler_name)
                handler(doc, None, p)
                continue

            # 自定义规则：读配置，检查 enabled
            if rules_config is None:
                continue
            rule_cfg = getattr(rules_config, rule_name, None)
            if rule_cfg is None or not rule_cfg.enabled:
                continue
            handler = getattr(self, handler_name)
            handler(doc, rule_cfg, p)

    # ------------------------------------------------------------------
    # 默认规则 handler：段落样式 + 字符样式
    # ------------------------------------------------------------------

    def _handle_paragraph_style(self, doc, rule_cfg, p: bool):
        """默认段落样式检查/应用。配置需包含 alignment 等格式字段。"""
        if self.paragraph is None or not self.paragraph.runs:
            return
        from wordformat.style.diff import ParagraphStyle

        cfg = self.pydantic_config
        if cfg.alignment is None:
            return
        ps = ParagraphStyle.from_config(cfg)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
            if issues:
                self.add_comment(
                    doc=doc,
                    runs=self.paragraph.runs,
                    text=ParagraphStyle.to_string(issues, target=self.NODE_LABEL),
                )
        else:
            ps.apply_to_paragraph(self.paragraph)
            # apply 后再次 diff，只报告修正失败的残留差异
            remaining = ps.diff_from_paragraph(self.paragraph)
            if remaining:
                self.add_comment(
                    doc=doc,
                    runs=self.paragraph.runs,
                    text=ParagraphStyle.to_string(remaining, target=self.NODE_LABEL),
                )

    def _handle_character_style(self, doc, rule_cfg, p: bool):
        """默认字符样式检查/应用。配置需包含 chinese_font_name 等格式字段。"""
        if not self.paragraph.runs:
            return
        from wordformat.style.diff import CharacterStyle

        cfg = self.pydantic_config
        if cfg.chinese_font_name is None:
            return
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue
            if p:
                diff = cstyle.diff_from_run(run)
                if diff:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(diff, target=self.NODE_LABEL),
                    )
            else:
                cstyle.apply_to_run(run)
                # apply 后再次 diff，只报告修正失败的残留差异
                remaining = cstyle.diff_from_run(run)
                if remaining:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(
                            remaining, target=self.NODE_LABEL
                        ),
                    )

    def _collect_comment(self, runs, text: str) -> None:
        """将批注文本加入缓冲区，(runs, text) 成对存储。"""
        if text.strip():
            self._comment_texts.append((runs, text.strip()))

    def _flush_comments(self, doc: Document) -> None:
        """将缓冲的批注按锚点分组写入。段落级合并为一条，run 级按 run 分开。"""
        if not self._comment_texts or not self.paragraph.runs:
            self._comment_texts.clear()
            return
        groups = self._group_comments()
        self._write_comment_groups(doc, groups)
        self._comment_texts.clear()

    def _group_comments(self) -> dict[tuple, tuple[tuple, list[str]]]:
        """按 runs 分组批注文本。段落级用 ('__para__',)，run 级用索引 i。"""
        para_runs = list(self.paragraph.runs)
        groups: dict[tuple, tuple[tuple, list[str]]] = {}

        def _key(runs):
            if isinstance(runs, Sequence) and len(runs) == len(para_runs):
                return ("__para__",)
            r = runs[0] if isinstance(runs, Sequence) else runs
            for i, pr in enumerate(para_runs):
                if r is pr or r._element is pr._element:
                    return (i,)
            return (id(runs),)

        for runs, text in self._comment_texts:
            k = _key(runs)
            if k not in groups:
                groups[k] = (runs, [])
            groups[k][1].append(text)
        return groups

    def _write_comment_groups(self, doc, groups) -> None:
        """将分组后的批注写入文档并更新统计。

        每条批注按行拆分，`位置-问题类型：` 按严重度上色，正文保持黑色。
        """
        from wordformat.style.comments import (
            SEVERITY_ORDER,
            add_styled_comment,
            get_severity,
            split_comment_line,
        )

        for runs, texts in groups.values():
            paragraphs = [split_comment_line(t) for t in texts]
            add_styled_comment(doc, runs, paragraphs)
            max_sev = "提醒"
            for t in texts:
                sev = get_severity(t)
                if SEVERITY_ORDER.get(sev, 3) < SEVERITY_ORDER.get(max_sev, 3):
                    max_sev = sev
            FormatNode._error_stats["total"] += 1
            FormatNode._error_stats[max_sev] = (
                FormatNode._error_stats.get(max_sev, 0) + 1
            )

    @classmethod
    def reset_stats(cls) -> None:
        """重置全局错误统计。"""
        cls._error_stats = {"total": 0, "错误": 0, "提醒": 0}

    def check_format(self, doc: Document):
        """格式检查：先执行样式检查，再自动调度业务规则"""
        self._base(doc, p=True, r=True)
        self._run_rules(doc, p=True)
        self._flush_comments(doc)

    def apply_format(self, doc: Document):
        """格式应用：先清理、应用样式，再写入批注记录变更。"""
        self._clean_paragraph_edge_spaces()
        self._base(doc, p=False, r=False)
        self._run_rules(doc, p=False)
        self._flush_comments(doc)

    def apply_style(self, doc: Document):
        """格式应用（无批注）：仅应用样式，不生成批注。
        用于 md→docx 等从零生成文档的场景。"""
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
        """追加批注到缓冲区，按锚点 run 分组，flush 时同组合并为一条。"""
        self._collect_comment(runs, text)
