#! /usr/bin/env python
# @Time    : 2026/1/9 20:08
# @Author  : afish
# @File    : abstract.py
import re

from wordformat.config.dotdict import BASE_FORMAT
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register
from wordformat.style.diff import CharacterStyle, ParagraphStyle


@register("abstract_chinese_title", level=1)
class AbstractTitleCN(FormatNode):
    """摘要标题中文节点"""

    NODE_TYPE = "abstract.chinese.title"
    NODE_LABEL = "中文摘要标题"
    DEFAULTS = {
        **BASE_FORMAT,
        "alignment": "居中对齐",
        "first_line_indent": "0字符",
        "chinese_font_name": "黑体",
        "font_size": "小二",
        "bold": True,
    }


@register("abstract_chinese_title_content", level=1)
class AbstractTitleContentCN(FormatNode):
    """摘要标题正文混合中文节点"""

    NODE_TYPE = "abstract.chinese"
    NODE_LABEL = "中文摘要"
    DEFAULTS = {
        "title": {
            **BASE_FORMAT,
            "alignment": "居中对齐",
            "first_line_indent": "0字符",
            "chinese_font_name": "黑体",
            "font_size": "小二",
            "bold": True,
        },
        "body": {**BASE_FORMAT, "alignment": "两端对齐"},
    }

    def check_title(self, run) -> bool:
        """检查标题是否包含在正文中"""
        pattern = r"摘[^a-zA-Z0-9\u4e00-\u9fff]*要"
        if re.search(pattern, run.text):
            return True
        return False

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        # 混合节点的cfg有两个值，遂需要重新组装
        # 段落样式选择content样式
        # 字体样式"摘要"选择title，"正文"选择content
        ps = ParagraphStyle.from_config(cfg.body)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)

        for run in self.paragraph.runs:  # 检查标题是否包含在正文中
            run.text = run.text.replace("\r", "").replace("\n", "")
            if self.check_title(run):
                # 对run对象设置样式
                C = CharacterStyle(
                    font_name_cn=cfg.title.chinese_font_name,
                    font_name_en=cfg.title.english_font_name,
                    font_size=cfg.title.font_size,
                    font_color=cfg.title.font_color,
                    bold=cfg.title.bold,
                    italic=cfg.title.italic,
                    underline=cfg.title.underline,
                )
                if r:
                    diff_result = C.diff_from_run(run)
                else:
                    diff_result = C.apply_to_run(run)
            else:
                # 对剩余部分的run设置样式
                C = CharacterStyle(
                    font_name_cn=cfg.body.chinese_font_name,
                    font_name_en=cfg.body.english_font_name,
                    font_size=cfg.body.font_size,
                    font_color=cfg.body.font_color,
                    bold=cfg.body.bold,
                    italic=cfg.body.italic,
                    underline=cfg.body.underline,
                )
                if r:
                    diff_result = C.diff_from_run(run)
                else:
                    diff_result = C.apply_to_run(run)
            self.add_comment(
                doc=doc,
                runs=run,
                text=CharacterStyle.to_string(diff_result, target=self.NODE_LABEL),
            )
        self.add_comment(
            doc=doc,
            runs=self.paragraph.runs,
            text=ParagraphStyle.to_string(issues, target=self.NODE_LABEL),
        )


@register("abstract_chinese_content")
class AbstractContentCN(FormatNode):
    """摘要内容中文节点"""

    NODE_TYPE = "abstract.chinese.body"
    NODE_LABEL = "中文摘要正文"
    DEFAULTS = {**BASE_FORMAT, "alignment": "两端对齐"}

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        ps = ParagraphStyle.from_config(cfg)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for _, run in enumerate(self.paragraph.runs):
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            self.add_comment(
                doc=doc,
                runs=run,
                text=CharacterStyle.to_string(diff_result, target=self.NODE_LABEL),
            )
        self.add_comment(
            doc=doc,
            runs=self.paragraph.runs,
            text=ParagraphStyle.to_string(issues, target=self.NODE_LABEL),
        )


@register("abstract_english_title", level=1)
class AbstractTitleEN(FormatNode):
    """摘要标题英文节点"""

    NODE_TYPE = "abstract.english.title"
    NODE_LABEL = "英文摘要标题"
    DEFAULTS = {
        **BASE_FORMAT,
        "alignment": "居中对齐",
        "first_line_indent": "0字符",
        "font_size": "四号",
        "bold": True,
    }


@register("abstract_english_title_content", level=1)
class AbstractTitleContentEN(FormatNode):
    """摘要标题正文混合英文节点"""

    NODE_TYPE = "abstract.english"
    NODE_LABEL = "英文摘要"
    DEFAULTS = {
        "title": {
            **BASE_FORMAT,
            "alignment": "居中对齐",
            "first_line_indent": "0字符",
            "font_size": "四号",
            "bold": True,
        },
        "body": {**BASE_FORMAT, "alignment": "两端对齐"},
    }

    def _check_title_in_full_text(self, runs) -> int:
        """拼接全部 run 文本，返回 "Abstract" 前缀在 clean 文本中的结束位置。

        若全文不以 "abstract"（大小写不敏感）开头，返回 0。
        """
        full_text = "".join(run.text for run in runs)
        full_text = full_text.replace("\r", "").replace("\n", "")
        m = re.match(r"abstract", full_text, re.IGNORECASE)
        return m.end() if m else 0

    def _base(self, doc, p: bool, r: bool):
        """
        设置 摘要 样式
        """
        # 混合节点的cfg有两个值，遂需要重新组装
        # 段落样式选择content样式
        # 字体样式"摘要"选择title，"正文"选择content
        cfg = self.pydantic_config
        ps = ParagraphStyle.from_config(cfg.body)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)

        runs = self.paragraph.runs
        # 段落级别判断：即使 "Abstract" 被拆分到多个 run 也能正确识别
        title_end = self._check_title_in_full_text(runs)

        cum = 0  # 在 clean 全文中的累积位置
        for run in runs:
            text_clean = run.text.replace("\r", "").replace("\n", "")
            run_len = len(text_clean)
            run_start = cum
            run_end = cum + run_len
            in_title = run_start < title_end

            if in_title:
                if run_start == 0:
                    # 首个 run：修正 Abstract 大小写
                    if run_len >= title_end:
                        # "abstract" 完整在此 run 内 → 直接修正
                        m = re.match(r"abstract", text_clean, re.IGNORECASE)
                        if m:
                            run.text = "Abstract" + text_clean[m.end() :]
                    else:
                        # "Abstract" 被拆分到多个 run → 第一个 run 替换为 "Abstract"
                        run.text = "Abstract"
                elif run_end <= title_end:
                    # 后续 run 完全在标题前缀内 → 清除
                    run.text = ""
                else:
                    # 后续 run 部分在标题前缀内 → 裁掉标题部分
                    run.text = text_clean[title_end - run_start :]

                c = CharacterStyle(
                    font_name_cn=cfg.title.chinese_font_name,
                    font_name_en=cfg.title.english_font_name,
                    font_size=cfg.title.font_size,
                    font_color=cfg.title.font_color,
                    bold=cfg.title.bold,
                    italic=cfg.title.italic,
                    underline=cfg.title.underline,
                )
            else:
                c = CharacterStyle(
                    font_name_cn=cfg.body.chinese_font_name,
                    font_name_en=cfg.body.english_font_name,
                    font_size=cfg.body.font_size,
                    font_color=cfg.body.font_color,
                    bold=cfg.body.bold,
                    italic=cfg.body.italic,
                    underline=cfg.body.underline,
                )
            if r:
                diff_result = c.diff_from_run(run)
            else:
                diff_result = c.apply_to_run(run)
            self.add_comment(
                doc=doc,
                runs=run,
                text=CharacterStyle.to_string(diff_result, target=self.NODE_LABEL),
            )
            cum += len(text_clean)
        self.add_comment(
            doc=doc,
            runs=self.paragraph.runs,
            text=ParagraphStyle.to_string(issues, target=self.NODE_LABEL),
        )


@register("abstract_english_content")
class AbstractContentEN(FormatNode):
    """摘要内容英文节点"""

    NODE_TYPE = "abstract.english.body"
    NODE_LABEL = "英文摘要正文"
    DEFAULTS = {**BASE_FORMAT, "alignment": "两端对齐"}

    def _base(self, doc, p: bool, r: bool):
        cfg = self.pydantic_config
        ps = ParagraphStyle.from_config(cfg)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
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
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:  # 可选：仅当有差异时才添加批注
                self.add_comment(
                    doc=doc,
                    runs=run,
                    text=CharacterStyle.to_string(diff_result, target=self.NODE_LABEL),
                )
        if issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text="".join(str(dr) for dr in issues),
            )
