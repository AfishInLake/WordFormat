#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from docx.document import Document
from loguru import logger

from wordformat.config.datamodel import HeadingLevelConfig
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle


class BaseHeadingNode(FormatNode[HeadingLevelConfig]):
    """标题节点基类（复用1/2/3级标题的通用逻辑）

    CONFIG_PATH 由子类定义（如 "headings.level_1"），
    load_config 由 FormatNode 基类统一处理（CONFIG_PATH getattr 链）。
    """

    LEVEL: int = 0  # 标题层级（1/2/3）
    NODE_TYPE: str = ""
    CONFIG_MODEL = HeadingLevelConfig

    def _base(self, doc: Document, p: bool, r: bool):
        """通用标题格式检查逻辑"""
        # 修复：空值校验（直接检查底层私有属性 _pydantic_config）
        if self._pydantic_config is None:
            error_msg = f"{self.LEVEL}级标题配置未加载，跳过检查"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=error_msg)
            logger.warning(error_msg)
            return [
                {"error": error_msg, "level": self.LEVEL, "node_type": self.NODE_TYPE}
            ]

        # 类型断言（读取@property的 pydantic_config，本质是 _pydantic_config）
        cfg = self.pydantic_config

        # 段落样式：由 ParagraphStyle.apply_to_paragraph 先应用 builtin_style_name，
        # 再应用显式段落格式（对齐、缩进、间距等），后者覆盖前者的对应属性。
        ps = ParagraphStyle.from_config(cfg)
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        # 字符样式检查（完整使用所有配置字段）
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 逐run检查并添加批注
        run_issues = []
        for idx, run in enumerate(self.paragraph.runs):
            if not run.text.strip():
                continue  # 跳过空白run，减少无效检查
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                run_issue = {
                    "run_index": idx,
                    "run_text": run.text,
                    "diff": diff_result,
                }
                run_issues.append(run_issue)
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        # 段落样式批注
        all_issues = []
        if paragraph_issues:
            para_issue = {
                "type": "paragraph_style",
                "diff": paragraph_issues,
                "paragraph_text": self.paragraph.text[:50] + "..."
                if len(self.paragraph.text) > 50
                else self.paragraph.text,
            }
            all_issues.append(para_issue)
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

        # 添加字符样式问题
        if run_issues:
            all_issues.extend(run_issues)

        return all_issues


# 各层级标题节点（无需重写check_format，直接复用基类逻辑）
class HeadingLevel1Node(BaseHeadingNode):
    """一级标题节点"""

    LEVEL = 1
    NODE_TYPE = "headings.level_1"
    CONFIG_PATH = "headings.level_1"


class HeadingLevel2Node(BaseHeadingNode):
    """二级标题节点"""

    LEVEL = 2
    NODE_TYPE = "headings.level_2"
    CONFIG_PATH = "headings.level_2"


class HeadingLevel3Node(BaseHeadingNode):
    """三级标题节点"""

    LEVEL = 3
    NODE_TYPE = "headings.level_3"
    CONFIG_PATH = "headings.level_3"
