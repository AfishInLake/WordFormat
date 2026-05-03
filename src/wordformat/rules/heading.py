#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from docx.document import Document
from docx.enum.dml import MSO_COLOR_TYPE
from docx.oxml import OxmlElement
from loguru import logger  # 推荐添加日志，便于调试

from wordformat.config.datamodel import HeadingLevelConfig, NodeConfigRoot
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle
from wordformat.style.style_enum import FontColor


class BaseHeadingNode(FormatNode[HeadingLevelConfig]):
    """标题节点基类（复用1/2/3级标题的通用逻辑）"""

    LEVEL: int = 0  # 标题层级（1/2/3）
    NODE_TYPE: str = ""
    CONFIG_MODEL = HeadingLevelConfig

    def _get_level_config(self, root_config: NodeConfigRoot) -> HeadingLevelConfig:
        """根据层级获取对应配置"""
        level_config_map = {
            1: root_config.headings.level_1,
            2: root_config.headings.level_2,
            3: root_config.headings.level_3,
        }
        target_config = level_config_map.get(self.LEVEL, root_config.headings.level_1)
        return target_config

    def load_config(self, root_config: dict | NodeConfigRoot):
        """重载加载配置方法，自动匹配对应层级的配置"""
        try:
            if isinstance(root_config, dict):
                # 修复：使用单下划线 _config（匹配基类@property的底层属性）
                level_config_dict = root_config.get("headings", {}).get(
                    f"level_{self.LEVEL}", {}
                )
                self._config = level_config_dict  # 正确赋值给单下划线私有属性
                logger.debug(f"{self.LEVEL}级标题字典配置：{self._config}")
                self._pydantic_config = self.CONFIG_MODEL(
                    **self._config
                )  # 读取赋值后的 _config

            elif isinstance(root_config, NodeConfigRoot):
                # 修复：先赋值 _pydantic_config，再同步到 _config
                self._pydantic_config = self._get_level_config(root_config)
                self._config = (
                    self._pydantic_config.model_dump()
                )  # 读取底层 _pydantic_config
            else:
                raise TypeError(
                    f"配置类型不支持：{type(root_config)}，仅支持dict或NodeConfigRoot"
                )

        except Exception as e:
            logger.error(f"{self.LEVEL}级标题配置加载失败：{str(e)}")
            raise  # 抛出异常，避免后续使用错误配置

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

        # 格式化模式下：先强制应用 builtin_style_name，清除段落级别的显式格式覆盖
        # （如显式设置的 jc=both 会覆盖样式中的居中对齐）
        # 直接操作 XML 设置 pStyle，同时移除显式的 jc 等段落格式覆盖
        if not p:
            builtin_name = getattr(cfg, "builtin_style_name", None)
            if builtin_name:
                from docx.oxml.ns import qn
                pPr = self.paragraph._element.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    self.paragraph._element.insert(0, pPr)

                # 移除显式的对齐方式（让段落继承样式中的对齐）
                jc = pPr.find(qn("w:jc"))
                if jc is not None:
                    pPr.remove(jc)

                # 设置 pStyle
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    pStyle.set(qn("w:val"), builtin_name)
                else:
                    pStyle = OxmlElement("w:pStyle")
                    pStyle.set(qn("w:val"), builtin_name)
                    pPr.insert(0, pStyle)

        # 段落样式检查（完整使用所有配置字段）
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

    @staticmethod
    def _fix_style_definition_color(doc: Document, cfg: HeadingLevelConfig):
        """修正样式定义级别的字体颜色，清除主题色引用。

        当 Word 样式定义（如 Heading 1）使用了 themeColor 时，
        段落内的 run 即使没有显式设置颜色，也会继承样式中的主题色。
        仅修改 run 级别的颜色不够，必须同时修正样式定义本身。

        直接操作底层 XML 元素，确保 themeColor 属性被彻底移除。
        """
        style_name = getattr(cfg, "builtin_style_name", None)
        if not style_name:
            return

        try:
            style = doc.styles[style_name]
        except KeyError:
            logger.info(f"样式 '{style_name}' 不存在，跳过样式级别颜色修正")
            return

        # 直接操作样式定义的 XML 元素
        from docx.oxml.ns import qn
        style_element = style.element
        rPr = style_element.find(qn("w:rPr"))
        if rPr is None:
            logger.info(f"样式 '{style_name}' 无 rPr，跳过样式级别颜色修正")
            return

        color_elem = rPr.find(qn("w:color"))
        if color_elem is None:
            logger.info(f"样式 '{style_name}' 无 color 元素，跳过样式级别颜色修正")
            return

        # 检查是否有主题色属性
        has_theme = (
            color_elem.get(qn("w:themeColor")) is not None
            or color_elem.get(qn("w:themeTint")) is not None
            or color_elem.get(qn("w:themeShade")) is not None
        )
        if not has_theme:
            return

        font_color_str = getattr(cfg, "font_color", "黑色") or "黑色"
        try:
            fc = FontColor(font_color_str)
            rgb_tuple = fc.rel_value
            hex_color = f"{rgb_tuple[0]:02X}{rgb_tuple[1]:02X}{rgb_tuple[2]:02X}"

            # 移除旧 color 元素，创建新的干净 color 元素（无主题属性）
            rPr.remove(color_elem)
            new_color = OxmlElement("w:color")
            new_color.set(qn("w:val"), hex_color)
            rPr.append(new_color)

            logger.debug(
                f"已修正样式 '{style_name}' 的主题色 → #{hex_color}"
            )
        except Exception as e:
            logger.warning(f"修正样式 '{style_name}' 颜色失败: {e}")


# 各层级标题节点（无需重写check_format，直接复用基类逻辑）
class HeadingLevel1Node(BaseHeadingNode):
    """一级标题节点"""

    LEVEL = 1
    NODE_TYPE = "headings.level_1"


class HeadingLevel2Node(BaseHeadingNode):
    """二级标题节点"""

    LEVEL = 2
    NODE_TYPE = "headings.level_2"


class HeadingLevel3Node(BaseHeadingNode):
    """三级标题节点"""

    LEVEL = 3
    NODE_TYPE = "headings.level_3"
