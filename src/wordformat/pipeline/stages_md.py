#! /usr/bin/env python
# @Time    : 2026/7/12
# @Author  : afish
# @File    : stages_md.py
"""Markdown → Docx 转换的 pipeline 阶段。

Markdown 仅提供文本内容和文档结构（标题层级、段落划分），
所有格式（字体、字号、段落样式）由 YAML 配置 + FormatNode 统一处理。
图片插入、表格创建等由对应 FormatNode 子类负责。
"""

from pathlib import Path

from docx import Document

from wordformat.config.dotdict import deep_merge
from wordformat.log_config import logger
from wordformat.markdown.parser import parse_markdown
from wordformat.math.omml import latex_to_omath, latex_to_omath_para
from wordformat.rules.node import FormatNode
from wordformat.utils import resolve_item_text

from .context import FormatContext

# 行内富文本默认配置（可由 YAML `inline` 段覆盖）
_DEFAULT_INLINE_CONFIG = {
    "code_font": {
        "chinese_font_name": "宋体",
        "english_font_name": "Consolas",
        "font_size": "五号",
    },
    "link": {"style": "Hyperlink", "color": "theme", "underline": True},
}
# 链接主题色（Word 默认超链接蓝）
_THEME_LINK_COLOR = "0563C1"


class LoadMarkdownStage:
    """读取 Markdown 文件到内存。"""

    def process(self, ctx: FormatContext) -> FormatContext:
        md_path = Path(ctx.md_path)
        if not md_path.exists():
            raise FileNotFoundError(f"Markdown 文件不存在: {md_path}")
        ctx.md_text = md_path.read_text(encoding="utf-8")
        logger.info(f"已加载 Markdown 文件: {md_path.resolve()}")
        return ctx


class MarkdownParseStage:
    """解析 Markdown 为扁平段落列表（文本 + inline_marks 行内标记）。"""

    def process(self, ctx: FormatContext) -> FormatContext:
        ctx.paragraphs = parse_markdown(ctx.md_text)
        logger.info(f"解析完成，共 {len(ctx.paragraphs)} 个段落")
        return ctx


class DocumentCreationStage:
    """从 FormatNode 树创建新的 .docx Document。

    普通节点创建一个段落并填充文本；含 inline_marks 的节点生成多 run 富文本。
    图片、表格由后续 FormattingExecutionStage 中各 FormatNode 子类处理；
    exists/sync_skip 表格不建占位段。
    """

    def process(self, ctx: FormatContext) -> FormatContext:
        ctx.document = Document()
        nodes = self._flatten_tree_nodes(ctx.root_node)
        inline_cfg = self._inline_config(ctx)
        logger.info(f"开始创建文档，共 {len(nodes)} 个节点")

        for node in nodes:
            value = node.value
            category = value.get("category", "") if isinstance(value, dict) else ""
            # 原文档已有表格（exists）或 sync 忽略的表格不占段落槽位；
            # 若仍建占位段，jd 会在正文中留下一个空段。
            if (
                category == "table_object"
                and isinstance(value, dict)
                and (value.get("exists") or value.get("sync_skip"))
            ):
                continue
            # jd 重建兼容：agent 编辑过的 JSON 用 replace 覆写文本（md 条目不带 replace）
            text = resolve_item_text(value) if isinstance(value, dict) else str(value)
            inline_marks = (
                value.get("inline_marks", []) if isinstance(value, dict) else []
            )
            # replace 是显式覆写：旧 inline_marks 的文本若已不一致，丢弃标记，
            # 否则会渲染出 replace 之前的旧文本。
            if (
                isinstance(value, dict)
                and isinstance(value.get("replace"), str)
                and inline_marks
            ):
                marks_text = "".join(
                    str(seg.get("text", ""))
                    for seg in inline_marks
                    if isinstance(seg, dict)
                )
                if marks_text.strip() != text.strip():
                    logger.warning(
                        f"条目 {category} 的 replace 与 inline_marks 文本不一致，"
                        "已以 replace 为准并忽略行内标记"
                    )
                    inline_marks = []
            para = ctx.document.add_paragraph()
            # figure_image: 路径留给 FigureImage._try_insert_image 处理，这里只建空段落
            if category == "figure_image":
                para.add_run("")
            elif category == "math_block":
                self._add_math_paragraph(para, text)
            elif self._has_math(inline_marks):
                self._add_math_runs(para, inline_marks)
            elif self._has_rich_marks(inline_marks):
                self._add_inline_runs(para, inline_marks, inline_cfg)
                # 标记内容级强调，后续字符格式化不把 bold/italic/underline 拉平
                node.protected_inline = True
            else:
                para.add_run(text.strip())
            node.paragraph = para

        return ctx

    @staticmethod
    def _inline_config(ctx: FormatContext) -> dict:
        """读取 YAML `inline` 段并与默认值深合并。"""
        config = ctx.config_model
        user_cfg = config.get("inline") if isinstance(config, dict) else None
        if isinstance(user_cfg, dict):
            return deep_merge(_DEFAULT_INLINE_CONFIG, user_cfg)
        return dict(_DEFAULT_INLINE_CONFIG)

    @staticmethod
    def _has_math(inline_marks: list[dict]) -> bool:
        return any(m.get("math") for m in inline_marks)

    @staticmethod
    def _has_rich_marks(inline_marks: list[dict]) -> bool:
        """是否含需要映射为多 run 的行内格式（粗/斜/删/代码/链接）。"""
        return any(
            m.get("bold")
            or m.get("italic")
            or m.get("code")
            or m.get("strikethrough")
            or m.get("link_url")
            for m in inline_marks
        )

    @staticmethod
    def _add_math_paragraph(para, latex: str):
        """创建块级公式段落（居中显示）。"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        omp = latex_to_omath_para(latex)
        if omp is not None:
            para._p.append(omp)
        else:
            para.add_run(latex)

    @staticmethod
    def _add_math_runs(para, inline_marks: list[dict]):
        """根据 inline_marks 创建段落，数学片段渲染为 OMML 元素。"""
        for seg in inline_marks:
            if seg.get("math"):
                latex = seg.get("text", "")
                if seg.get("math_display"):
                    # 块级公式在段落中：使用 oMathPara
                    omp = latex_to_omath_para(latex)
                    if omp is not None:
                        para._p.append(omp)
                    else:
                        para.add_run(latex)
                else:
                    om = latex_to_omath(latex)
                    if om is not None:
                        para._p.append(om)
                    else:
                        para.add_run(latex)
            else:
                para.add_run(seg.get("text", ""))

    @classmethod
    def _add_inline_runs(cls, para, inline_marks: list[dict], inline_cfg: dict):
        """根据 inline_marks 创建多 run 段落（粗/斜/删/代码/链接）。"""
        for seg in inline_marks:
            text = seg.get("text", "")
            if not text:
                continue
            link_url = seg.get("link_url")
            if link_url:
                cls._add_hyperlink(para, text, link_url, inline_cfg)
                continue
            run = para.add_run(text)
            if seg.get("bold"):
                run.bold = True
            if seg.get("italic"):
                run.italic = True
            if seg.get("strikethrough"):
                run.font.strike = True
            if seg.get("code"):
                cls._apply_code_font(run, inline_cfg)

    @staticmethod
    def _apply_code_font(run, inline_cfg: dict) -> None:
        """代码片段：字体取 inline.code_font，中文字体写 eastAsia。"""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt

        cfg = inline_cfg.get("code_font") or {}
        en_font = cfg.get("english_font_name", "Consolas")
        cn_font = cfg.get("chinese_font_name", "宋体")
        run.font.name = en_font
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:eastAsia"), cn_font)
        size = cfg.get("font_size")
        if size:
            from wordformat.style.defs import FontSize

            try:
                run.font.size = Pt(float(FontSize(size).rel_value))
            except (ValueError, TypeError):
                pass

    @staticmethod
    def _add_hyperlink(para, text: str, url: str, inline_cfg: dict) -> None:
        """插入 w:hyperlink（外部关系），样式/颜色/下划线取 inline.link 配置。"""
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        cfg = inline_cfg.get("link") or {}
        r_id = para.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)

        run = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        style_id = cfg.get("style")
        if style_id:
            rstyle = OxmlElement("w:rStyle")
            rstyle.set(qn("w:val"), str(style_id))
            rpr.append(rstyle)
        color = cfg.get("color")
        if color:
            color_val = _THEME_LINK_COLOR if color == "theme" else str(color)
            color_el = OxmlElement("w:color")
            color_el.set(qn("w:val"), color_val)
            rpr.append(color_el)
        if cfg.get("underline", True):
            underline = OxmlElement("w:u")
            underline.set(qn("w:val"), "single")
            rpr.append(underline)
        run.append(rpr)

        text_el = OxmlElement("w:t")
        text_el.text = text
        text_el.set(qn("xml:space"), "preserve")
        run.append(text_el)
        hyperlink.append(run)
        para._p.append(hyperlink)

    @staticmethod
    def _flatten_tree_nodes(root_node: FormatNode) -> list[FormatNode]:
        result: list[FormatNode] = []

        def dfs(node):
            for child in node.children:
                result.append(child)
                dfs(child)

        dfs(root_node)
        return result
