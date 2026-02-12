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
from docx.text.run import Run
from loguru import logger


def run_set_font_name(run: Run, font_name: str):
    """
    设置 Run 对象的中文字体名称（修复空值问题，兼容无样式配置的 Run）

    参数:
        run (docx.text.run.Run): 要设置字体的 Run 对象。
        font_name (str): 字体名称，如 "Microsoft YaHei"、"SimSun"、"Arial" 等。
    """
    # 1. 获取 Run 对应的 XML 元素
    r = run._element
    # 2. 检查并创建 rPr 节点（字体样式根节点，不存在则创建）
    if r.rPr is None:
        # 手动创建 rPr 节点（docx 原生XML结构，兼容所有版本）
        rPr = parse_xml(
            r'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        r.append(rPr)
    # 3. 检查并创建 rFonts 节点（字体名称配置节点，不存在则创建）
    if r.rPr.rFonts is None:
        # 手动创建 rFonts 节点
        rFonts = parse_xml(
            r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        r.rPr.append(rFonts)
    # 4. 安全设置中文字体（eastAsia 对应中文/东亚字体）
    r.rPr.rFonts.set(qn("w:eastAsia"), font_name)


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


class _SetSpacing:
    """设置间距距函数"""

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
            # 校验行值：非负数，过大值兜底（避免排版异常）
            value = max(float(value), 0.0)
            if value > 10.0:
                logger.warning(f"间距{value}行过大，自动兜底为10行")
                value = 10.0

            # 核心：Word中Lines属性以「1/100行」为单位存储（如0.5行→50，1行→100）
            lines_100unit = int(round(value * 100))  # 四舍五入为整数，匹配Word存储

            p = paragraph._element
            # 步骤1：创建/获取pPr节点（段落属性根节点，无则创建）
            pPr = p.find(qn("w:pPr"))
            if pPr is None:
                pPr = parse_xml(
                    r'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
                )
                p.append(pPr)

            # 步骤2：创建/获取spacing节点（间距节点，无则创建）
            spacing = pPr.find(qn("w:spacing"))
            if spacing is None:
                spacing = parse_xml(
                    r'<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
                )
                pPr.append(spacing)

            # 步骤3：设置核心属性w:XXLines（优先级最高，与获取函数对应）
            spacing.set(qn(f"w:{spacing_type}Lines"), str(lines_100unit))

            # 步骤4：清除兜底的w:XX twips属性，避免冲突（Word会优先读取Lines，清除后更纯净）
            if spacing.get(qn(f"w:{spacing_type}")) is not None:
                spacing.attrib.pop(qn(f"w:{spacing_type}"))

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
        if spacing_type == "before":
            paragraph.paragraph_format.space_before = Mm(target_value)
        else:
            paragraph.paragraph_format.space_after = Mm(target_value)


class _SetLineSpacing:
    """
    设置段落行距
    """

    @staticmethod
    def set_pt(paragraph: Paragraph, target_value: float):
        """
        使用 PT/磅 为单位设置行距
        Args:
            paragraph: 段落对象
            target_value: PT/磅数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Pt(target_value)

    @staticmethod
    def set_cm(paragraph: Paragraph, target_value: float):
        """
        使用 厘米(cm/CM) 为单位设置行距
        Args:
            paragraph: 段落对象
            target_value: 厘米数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Cm(target_value)

    @staticmethod
    def set_inch(paragraph: Paragraph, target_value: float):
        """
        使用 英寸(inch/Inches) 为单位设置行距
        Args:
            paragraph: 段落对象
            target_value: 英寸数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Inches(target_value)

    @staticmethod
    def set_mm(paragraph: Paragraph, target_value: float):
        """
        使用 毫米(mm/MM) 为单位设置行距
        Args:
            paragraph: 段落对象
            target_value: 毫米数值
        """
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Mm(target_value)


class _SetIndent:
    """
    缩进 文本之前(R)/ 文本之后(X):
    """

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
            if indent_type not in ("R", "X"):
                logger.error(
                    f"无效的缩进类型: {indent_type}. 仅支持 'R'（左）或 'X'（右）"
                )
                return False

            p = paragraph._element
            pPr = p.get_or_add_pPr()

            # 获取或创建 <w:ind> 元素
            ind = pPr.find(qn("w:ind"))
            if ind is None:
                from docx.oxml.text.parfmt import CT_Ind

                ind = CT_Ind()
                pPr.append(ind)

            # 清除旧的字符缩进和物理缩进（根据类型）
            char_attr = qn("w:leftChars") if indent_type == "R" else qn("w:rightChars")
            physical_attr = qn("w:left") if indent_type == "R" else qn("w:right")

            # 设置新值
            if value == 0:
                # 删除属性
                if char_attr in ind.attrib:
                    del ind.attrib[char_attr]
                if physical_attr in ind.attrib:
                    del ind.attrib[physical_attr]
            else:
                # 1 字符 = 100 单位
                chars_int = int(round(float(value) * 100))
                ind.set(char_attr, str(chars_int))
                # 注意：不要设置 physical_attr，Word 会优先使用 *Chars

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
        _SetIndent._apply_indent(paragraph, indent_type, value)

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
        _SetIndent._apply_indent(paragraph, indent_type, value)

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
        _SetIndent._apply_indent(paragraph, indent_type, value)

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
        _SetIndent._apply_indent(paragraph, indent_type, value)

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


class _SetFirstLineIndent:
    """
    缩进 特殊格式：首行缩进(>0)/悬挂缩进(<0)
    """

    @staticmethod
    def clear(paragraph: Paragraph):
        """
        清除段落的首行缩进设置（包括 firstLine 和 firstLineChars 字段），
        同时保留左缩进（left）和右缩进（right）
        Args:
            paragraph: 段落对象
        """
        p = paragraph._element
        pPr = p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))

        if ind is None:
            # 没有缩进设置，无需操作
            return

        # 保留 left 和 right
        left = ind.get(qn("w:left"))
        right = ind.get(qn("w:right"))

        # 创建新的属性字典，只保留 left / right
        new_attrs = {}
        if left is not None:
            new_attrs[qn("w:left")] = left
        if right is not None:
            new_attrs[qn("w:right")] = right

        if new_attrs:
            # 有保留的属性，更新 <w:ind>
            # 先清空所有属性
            ind.clear()
            # 再设置保留的属性
            for k, v in new_attrs.items():
                ind.set(k, v)
        else:
            # 没有保留属性，直接删除 <w:ind>
            pPr.remove(ind)

    @staticmethod
    def set_char(  # noqa C901
        paragraph: Paragraph, value: float | int
    ):
        """
        设置首行缩进，支持字符/物理单位，写入Word原生XML字段（firstLineChars/firstLine）
        :param paragraph: 目标段落对象
        :param value: 缩进值（字符单位为浮点数/整数，如2/2.0；物理单位为整数pt，如24）
        :return: 设置成功返回True，失败返回False
        """
        _SetFirstLineIndent.clear(paragraph)
        try:
            if value < 0:
                logger.warning("缩进值不能为负，自动置为0")
                value = 0

            p = paragraph._element
            pPr = p.get_or_add_pPr()

            # 保留 left / right
            existing_left = None
            existing_right = None
            ind = pPr.find(qn("w:ind"))
            if ind is not None:
                existing_left = ind.get(qn("w:left"))
                existing_right = ind.get(qn("w:right"))
                pPr.remove(ind)

            attrs = {}
            if existing_left is not None:
                attrs["w:left"] = existing_left
            if existing_right is not None:
                attrs["w:right"] = existing_right

            if value == 0:
                # ✅ 现实主义方案：同时写 firstLine=0 和 firstLineChars=0
                # 虽然冗余，但能确保 WPS/Word 都清零
                attrs["w:firstLine"] = "0"
                attrs["w:firstLineChars"] = "0"
                attrs["w:hanging"] = "0"  # 顺手清 hanging
            else:
                chars_int = int(round(float(value) * 100))
                attrs["w:firstLineChars"] = str(chars_int)

            attr_str = " ".join(f'{k}="{v}"' for k, v in attrs.items())
            ind_xml = f'<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" {attr_str}/>'  # noqa E501

            new_ind = parse_xml(ind_xml)
            pPr.append(new_ind)
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
        _SetFirstLineIndent.clear(paragraph)
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
        _SetFirstLineIndent.clear(paragraph)
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
        _SetFirstLineIndent.clear(paragraph)
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
        _SetFirstLineIndent.clear(paragraph)
        # 强制清除firstLineChars，避免优先级冲突
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None and ind.get(qn("w:firstLineChars")):
            del ind.attrib[qn("w:firstLineChars")]
        paragraph.paragraph_format.first_line_indent = Cm(value)
