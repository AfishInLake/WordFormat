#! /usr/bin/env python
# @Time    : 2026/1/20 18:30
# @Author  : afish
# @File    : set_some.py
"""
设置 段落/字体 属性
"""

from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Mm, Pt
from docx.text.paragraph import Paragraph
from loguru import logger


def run_set_font_name(run, font_name: str):
    """设置 Run 对象的东亚字体名称（python-docx Font 不支持 eastAsia）。"""
    rPr = run._element.get_or_add_rPr()
    rPr.get_or_add_rFonts().set(qn("w:eastAsia"), font_name)


def _paragraph_space_by_lines(
    paragraph: Paragraph, before_lines: float = 0.0, after_lines: float = 0.0
) -> None:
    """
    设置段落【段前/段后间距】为指定行数（Word原生行单位，非Pt磅值）
    核心：单次删旧建全新spacing节点，与行距完全解耦，支持单/双侧同时设置
    :param paragraph: 目标段落对象
    :param before_lines: 段前间距行数，浮点型（如0.1/0.7，>0时生效）
    :param after_lines: 段后间距行数，浮点型（如0.1/0.7，>0时生效）
    """
    pPr = paragraph._element.get_or_add_pPr()
    # 步骤1：读取已有属性（若存在spacing节点，保留未被修改的属性）
    existing_before = 0.0
    existing_after = 0.0
    # 查找已有spacing节点，避免直接删除
    for spacing in pPr.xpath(".//w:spacing"):
        # 读取已有段前属性，转换为行数值
        if spacing.get(qn("w:beforeLines")):
            existing_before = int(spacing.get(qn("w:beforeLines"))) / 100
        # 读取已有段后属性，转换为行数值
        if spacing.get(qn("w:afterLines")):
            existing_after = int(spacing.get(qn("w:afterLines"))) / 100
        # 先暂存，后续统一删除

    # 步骤2：删除所有旧spacing节点（保持原有核心逻辑，避免属性冲突）
    for spacing in pPr.xpath(".//w:spacing"):
        pPr.remove(spacing)

    # 步骤3：合并属性值（新设置值>0则用新值，否则用原有值）
    final_before = before_lines if before_lines > 0 else existing_before
    final_after = after_lines if after_lines > 0 else existing_after

    # 步骤4：新建spacing节点，同时设置合并后的双侧属性
    spacing = OxmlElement("w:spacing")
    if final_before > 0:
        spacing.set(qn("w:beforeLines"), str(int(final_before * 100)))
    if final_after > 0:
        spacing.set(qn("w:afterLines"), str(int(final_after * 100)))

    pPr.append(spacing)


def set_paragraph_space_before_by_lines(paragraph: Paragraph, lines: float) -> None:
    """
    设置段落【段前间距】为指定行数（Word 原生 w:beforeLines 行单位，非 Pt 磅值）
    :param paragraph: 目标段落对象
    :param lines: 段前间距行数，浮点型（如 0.0/1.0/0.5），直接映射 Word w:beforeLines
    """
    _paragraph_space_by_lines(paragraph, before_lines=lines)


def set_paragraph_space_after_by_lines(paragraph: Paragraph, lines: float) -> None:
    """
    设置段落【段后间距】为指定行数（Word 原生 w:afterLines 行单位，非 Pt 磅值）
    :param paragraph: 目标段落对象
    :param lines: 段后间距行数，浮点型（如 0.0/1.0/0.5），直接映射 Word w:afterLines
    """
    _paragraph_space_by_lines(paragraph, after_lines=lines)


class SetSpacing:
    """设置间距距函数"""

    @staticmethod
    def _clear_conflicting_attrs(paragraph: Paragraph, spacing_type: str) -> None:
        """清除与 spacing 冲突的 XML 属性（Autospacing 和 Lines）。

        当通过物理单位（pt/cm/mm/inch）设置间距时，需清除自动间距标志
        和行单位属性，否则 Word 会优先读取它们而忽略新设置的 twips 值。
        """
        pPr = paragraph._element.find(qn("w:pPr"))
        if pPr is None:
            return
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            return
        autospacing_key = qn(f"w:{spacing_type}Autospacing")
        if spacing.get(autospacing_key) is not None:
            spacing.attrib.pop(autospacing_key)
        lines_key = qn(f"w:{spacing_type}Lines")
        if spacing.get(lines_key) is not None:
            spacing.attrib.pop(lines_key)

    @staticmethod
    def _set_hang_on_pPr(pPr, spacing_type: str, value: float):
        """在任意 pPr 元素上设置「行」单位间距（段落或样式定义通用）。"""
        value = max(float(value), 0.0)
        if value > 10.0:
            logger.warning(f"间距{value}行过大，自动兜底为10行")
            value = 10.0
        lines_100unit = int(round(value * 100))

        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)

        spacing.set(qn(f"w:{spacing_type}Lines"), str(lines_100unit))
        spacing.set(qn(f"w:{spacing_type}"), "0")
        autospacing_key = qn(f"w:{spacing_type}Autospacing")
        if spacing.get(autospacing_key) is not None:
            spacing.attrib.pop(autospacing_key)

    @staticmethod
    def set_hang(paragraph: Paragraph, spacing_type: str, value: float):
        """
        核心通用设置逻辑：段前/段后间距统一处理（行→Lines属性×100）
        完全复刻Word逻辑：
        1. 优先设置w:XXLines属性（值=行×100，Word原生存储规则）
        2. 自动创建pPr/spacing缺失的XML节点，无空值报错
        3. 清除兜底的w:XX twips属性，避免与Lines属性冲突
        4. 支持浮点数行值（如0.5行→50，1.2行→120）
        """
        try:
            pPr = paragraph._element.get_or_add_pPr()
            SetSpacing._set_hang_on_pPr(pPr, spacing_type, value)
        except Exception as e:
            logger.error(f"设置段落{spacing_type}间距{value}行失败: {e}")

    @staticmethod
    def set_pt(paragraph: Paragraph, spacing_type: str, target_value: float):
        """
        使用 PT/磅 为单位设置间距（1磅=1PT）
        Args:
            paragraph: 目标段落对象
            spacing_type: 间距类型（"before"=段前，"after"=段后）
            target_value: PT/磅数值
        """
        SetSpacing._clear_conflicting_attrs(paragraph, spacing_type)
        if spacing_type == "before":
            paragraph.paragraph_format.space_before = Pt(target_value)
        else:
            paragraph.paragraph_format.space_after = Pt(target_value)

    @staticmethod
    def set_cm(paragraph: Paragraph, spacing_type: str, target_value: float):
        """
        使用 厘米(cm/CM) 为单位设置间距
        Args:
            paragraph: 目标段落对象
            spacing_type: 间距类型（"before"=段前，"after"=段后）
            target_value: 厘米数值
        """
        SetSpacing._clear_conflicting_attrs(paragraph, spacing_type)
        if spacing_type == "before":
            paragraph.paragraph_format.space_before = Cm(target_value)
        else:
            paragraph.paragraph_format.space_after = Cm(target_value)

    @staticmethod
    def set_inch(paragraph: Paragraph, spacing_type: str, target_value: float):
        """
        使用 英寸(inch/Inches) 为单位设置间距
        Args:
            paragraph: 目标段落对象
            spacing_type: 间距类型（"before"=段前，"after"=段后）
            target_value: 英寸数值
        """
        SetSpacing._clear_conflicting_attrs(paragraph, spacing_type)
        if spacing_type == "before":
            paragraph.paragraph_format.space_before = Inches(target_value)
        else:
            paragraph.paragraph_format.space_after = Inches(target_value)

    @staticmethod
    def set_mm(paragraph: Paragraph, spacing_type: str, target_value: float):
        """
        使用 毫米(mm/MM) 为单位设置间距
        Args:
            paragraph: 目标段落对象
            spacing_type: 间距类型（"before"=段前，"after"=段后）
            target_value: 毫米数值
        """
        SetSpacing._clear_conflicting_attrs(paragraph, spacing_type)
        if spacing_type == "before":
            paragraph.paragraph_format.space_before = Mm(target_value)
        else:
            paragraph.paragraph_format.space_after = Mm(target_value)


class SetLineSpacing:
    """
    设置段落行距
    """

    @staticmethod
    def _set_on_pPr(pPr, line_rule: str, line_val: float):
        """在任意 pPr 元素上设置行距（段落或样式定义通用）。

        line_rule: "auto" | "exact" | "atLeast"
        line_val: 行距值（auto 模式下为 240×倍数，exact/atLeast 模式下为 twips）
        """
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:lineRule"), line_rule)
        spacing.set(qn("w:line"), str(int(round(line_val))))

    @staticmethod
    def set_pt(paragraph: Paragraph, value: float):
        """
        使用 PT/磅 为单位设置行距
        Args:
            paragraph: 段落对象
            value: PT/磅数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Pt(value)

    @staticmethod
    def set_cm(paragraph: Paragraph, value: float):
        """
        使用 厘米(cm/CM) 为单位设置行距
        Args:
            paragraph: 段落对象
            value: 厘米数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Cm(value)

    @staticmethod
    def set_inch(paragraph: Paragraph, value: float):
        """
        使用 英寸(inch/Inches) 为单位设置行距
        Args:
            paragraph: 段落对象
            value: 英寸数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Inches(value)

    @staticmethod
    def set_mm(paragraph: Paragraph, value: float):
        """
        使用 毫米(mm/MM) 为单位设置行距
        Args:
            paragraph: 段落对象
            value: 毫米数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Mm(value)


class SetIndent:
    """
    缩进 文本之前(R)/ 文本之后(X):
    """

    @staticmethod
    def _set_char_on_pPr(pPr, indent_type: str, value: float):
        """在任意 pPr 元素上设置「字符」单位缩进（段落或样式定义通用）。"""
        if indent_type not in ("R", "X"):
            raise ValueError(
                f"无效的缩进类型: {indent_type}. 仅支持 'R'（左）或 'X'（右）"
            )

        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)

        char_attr = qn("w:leftChars") if indent_type == "R" else qn("w:rightChars")
        physical_attr = qn("w:left") if indent_type == "R" else qn("w:right")

        if value == 0:
            if char_attr in ind.attrib:
                del ind.attrib[char_attr]
            if physical_attr in ind.attrib:
                del ind.attrib[physical_attr]
        else:
            ind.set(char_attr, str(int(round(float(value) * 100))))

    @staticmethod
    def set_char(paragraph: Paragraph, indent_type: str, value: float) -> bool:
        """
        使用 字符 为单位设置缩进
        Args:
            paragraph: 段落对象
            indent_type: 缩进类型（"R"=文本之前缩进，"X"=文本之后缩进）
            value: 字符数值
        """
        try:
            pPr = paragraph._element.get_or_add_pPr()
            SetIndent._set_char_on_pPr(pPr, indent_type, value)
            return True
        except Exception as e:
            logger.error(f"设置字符缩进失败 (type={indent_type}, value={value}): {e}")
            return False

    @staticmethod
    def set_pt(paragraph: Paragraph, indent_type: str, value: float):
        """
        使用 PT/磅 为单位设置缩进
        Args:
            paragraph: 段落对象
            indent_type: 缩进类型（"R"=文本之前缩进，"X"=文本之后缩进）
            value: PT/磅数值
        """
        value = Pt(value)
        SetIndent._apply_indent(paragraph, indent_type, value)

    @staticmethod
    def set_cm(paragraph: Paragraph, indent_type: str, value: float):
        """
        使用 厘米(cm/CM) 为单位设置缩进
        Args:
            paragraph: 段落对象
            indent_type: 缩进类型（"R"=文本之前缩进，"X"=文本之后缩进）
            value: 厘米数值
        """
        value = Cm(value)
        SetIndent._apply_indent(paragraph, indent_type, value)

    @staticmethod
    def set_inch(paragraph: Paragraph, indent_type: str, value: float):
        """
        使用 英寸(inch/Inches) 为单位设置缩进
        Args:
            paragraph: 段落对象
            indent_type: 缩进类型（"R"=文本之前缩进，"X"=文本之后缩进）
            value: 英寸数值
        """
        value = Inches(value)
        SetIndent._apply_indent(paragraph, indent_type, value)

    @staticmethod
    def set_mm(paragraph: Paragraph, indent_type: str, value: float):
        """
        使用 毫米(mm/MM) 为单位设置缩进
        Args:
            paragraph: 段落对象
            indent_type: 缩进类型（"R"=文本之前缩进，"X"=文本之后缩进）
            value: 毫米数值
        """
        value = Mm(value)
        SetIndent._apply_indent(paragraph, indent_type, value)

    @staticmethod
    def _apply_indent(paragraph: Paragraph, indent_type: str, value):
        """内部方法：应用缩进值"""
        if indent_type == "R":
            # "R" 对应左缩进（文本之前）
            paragraph.paragraph_format.left_indent = value
        elif indent_type == "X":
            # "X" 对应右缩进（文本之后）
            paragraph.paragraph_format.right_indent = value
        else:
            raise ValueError(
                f"无效的缩进类型: {indent_type}. 请使用 'R'（左缩进）或 'X'（右缩进）"
            )


class SetFirstLineIndent:
    """
    缩进 特殊格式：首行缩进(>0)/悬挂缩进(<0)
    """

    @staticmethod
    def _clear_ind_on_pPr(pPr):
        """清除 pPr 中 <w:ind> 的首行/悬挂缩进属性，保留 left/right。"""
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return
        left = ind.get(qn("w:left"))
        right = ind.get(qn("w:right"))
        new_attrs = {}
        if left is not None:
            new_attrs[qn("w:left")] = left
        if right is not None:
            new_attrs[qn("w:right")] = right
        if new_attrs:
            ind.clear()
            for k, v in new_attrs.items():
                ind.set(k, v)
        else:
            pPr.remove(ind)

    @staticmethod
    def clear(paragraph: Paragraph):
        """
        清除段落的首行缩进设置（包括 firstLine 和 firstLineChars 字段），
        同时保留左缩进（left）和右缩进（right）
        """
        pPr = paragraph._element.get_or_add_pPr()
        SetFirstLineIndent._clear_ind_on_pPr(pPr)

    @staticmethod
    def _set_char_on_pPr(pPr, value: float, existing_left=None, existing_right=None):
        """在任意 pPr 元素上设置「字符」单位首行缩进（段落或样式定义通用）。

        value > 0 → 首行缩进（w:firstLineChars）
        value < 0 → 悬挂缩进（w:hangingChars）
        value == 0 → 清除缩进
        """
        # 保留 left / right（如果存在）
        attrs = {}
        if existing_left is not None:
            attrs["w:left"] = existing_left
        if existing_right is not None:
            attrs["w:right"] = existing_right

        if value == 0:
            attrs["w:firstLine"] = "0"
            attrs["w:firstLineChars"] = "0"
            attrs["w:hanging"] = "0"
        elif value > 0:
            attrs["w:firstLineChars"] = str(int(round(float(value) * 100)))
        else:
            attrs["w:hangingChars"] = str(int(round(float(abs(value)) * 100)))

        attr_str = " ".join(f'{k}="{v}"' for k, v in attrs.items())
        ind_xml = f'<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" {attr_str}/>'  # noqa E501
        new_ind = parse_xml(ind_xml)
        pPr.append(new_ind)

    @staticmethod
    def set_char(  # noqa C901
        paragraph: Paragraph, value: float | int
    ):
        """
        设置首行缩进（>0）或悬挂缩进（<0），支持字符单位。
        首行缩进 → w:firstLineChars；悬挂缩进 → w:hangingChars

        :param paragraph: 目标段落对象
        :param value: 缩进值（正数为首行缩进，负数为悬挂缩进，如 -2.2 表示悬挂 2.2 字符）
        :return: 设置成功返回 True，失败返回 False
        """
        SetFirstLineIndent.clear(paragraph)
        try:
            pPr = paragraph._element.get_or_add_pPr()
            existing_left = None
            existing_right = None
            ind = pPr.find(qn("w:ind"))
            if ind is not None:
                existing_left = ind.get(qn("w:left"))
                existing_right = ind.get(qn("w:right"))
                pPr.remove(ind)
            SetFirstLineIndent._set_char_on_pPr(
                pPr, value, existing_left, existing_right
            )
            return True
        except Exception as e:
            logger.error(f"设置首行缩进失败：{e}")
            return False

    @staticmethod
    def set_inch(paragraph: Paragraph, value: float):
        """
        设置首行缩进，支持英寸单位，写入Word原生XML字段（firstLine/firstLineChars）
        :param para: 段落对象
        :param value: 缩进值（英寸单位为浮点数，如0.5）
        :return: 设置成功返回True，失败返回False
        """
        SetFirstLineIndent.clear(paragraph)
        # 强制清除firstLineChars，避免优先级冲突
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:firstLineChars")):
            del ind.attrib[qn("w:firstLineChars")]
        paragraph.paragraph_format.first_line_indent = Inches(value)

    @staticmethod
    def set_mm(paragraph: Paragraph, value: float):
        """
        设置首行缩进，支持毫米单位，写入Word原生XML字段（firstLine/firstLineChars）
        :param para: 段落对象
        :param value: 缩进值（毫米单位为浮点数，如5.0）
        :return:
        """
        SetFirstLineIndent.clear(paragraph)
        # 强制清除firstLineChars，避免优先级冲突
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:firstLineChars")):
            del ind.attrib[qn("w:firstLineChars")]
        paragraph.paragraph_format.first_line_indent = Mm(value)

    @staticmethod
    def set_pt(paragraph: Paragraph, value: float):
        """
        设置首行缩进，支持磅单位，写入Word原生XML字段（firstLine/firstLineChars）
        :param para: 段落对象
        :param value: 缩进值（磅单位为浮点数，如0.5）
        :return:
        """
        SetFirstLineIndent.clear(paragraph)
        # 强制清除firstLineChars，避免优先级冲突
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:firstLineChars")):
            del ind.attrib[qn("w:firstLineChars")]
        paragraph.paragraph_format.first_line_indent = Pt(value)

    @staticmethod
    def set_cm(paragraph: Paragraph, value: float):
        """
        设置首行缩进，支持厘米单位，写入Word原生XML字段（firstLine/firstLineChars）
        :param para: 段落对象
        :param value: 缩进值（厘米单位为浮点数，如0.5）
        :return:
        """
        SetFirstLineIndent.clear(paragraph)
        # 强制清除firstLineChars，避免优先级冲突
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:firstLineChars")):
            del ind.attrib[qn("w:firstLineChars")]
        paragraph.paragraph_format.first_line_indent = Cm(value)
