#! /usr/bin/env python
# @Time    : 2026/1/20 18:30
# @Author  : afish
# @File    : set_some.py
"""
设置 段落/字体 属性
"""

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt
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


def set_paragraph_space_before(paragraph: Paragraph, before_lines: float):
    """
    设置段落**段前间距**（单位：行），与paragraph_get_space_before逻辑完全对齐
    核心规则：Lines属性存储（×100）→ 自动创建缺失节点 → 兼容Word原生逻辑
    :param paragraph: 目标Paragraph对象
    :param before_lines: 段前间距（行），支持浮点数（如0.5/1/1.5）
    """
    _set_paragraph_spacing(paragraph, spacing_type="before", target_lines=before_lines)


def set_paragraph_space_after(paragraph: Paragraph, after_lines: float):
    """
    设置段落**段后间距**（单位：行），与paragraph_get_space_after逻辑完全对齐
    :param paragraph: 目标Paragraph对象
    :param after_lines: 段后间距（行），支持浮点数（如0.5/1/1.5）
    """
    _set_paragraph_spacing(paragraph, spacing_type="after", target_lines=after_lines)


def _set_paragraph_spacing(
    paragraph: Paragraph, spacing_type: str, target_lines: float
):
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
        target_lines = max(float(target_lines), 0.0)
        if target_lines > 10.0:
            logger.warning(f"间距{target_lines}行过大，自动兜底为10行")
            target_lines = 10.0

        # 核心：Word中Lines属性以「1/100行」为单位存储（如0.5行→50，1行→100）
        lines_100unit = int(round(target_lines * 100))  # 四舍五入为整数，匹配Word存储

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
        logger.error(f"设置段落{spacing_type}间距{target_lines}行失败: {e}")


def set_paragraph_first_line_indent(  # noqa C901
    paragraph: Paragraph,
    indent_chars: int,  # 你规定的缩进字符数（核心：自定义，如1/2/3）
    default_font_size_pt: float = 12.0,
):
    """
    为段落设置**自定义字符数**的首行缩进，与paragraph_get_first_line_indent计算逻辑完全一致
    核心互逆公式：目标缩进pt值 = 段落实际字体大小pt × 自定义缩进字符数
    Params:
        paragraph: python-docx的Paragraph对象（目标段落）
        indent_chars: 你规定的缩进字符数（必填，int类型，如1/2/3，支持正整数）
        default_font_size_pt: 默认字体大小（pt），段落无有效字体大小时使用，默认12.0pt（宋体小四）
    Return:
        bool: 设置成功返回True，失败返回False
    """
    try:
        # 校验缩进字符数：仅支持正整数（符合论文排版常规，避免无效值）
        if not isinstance(indent_chars, int) or indent_chars < 1:
            logger.error(f"缩进字符数必须为正整数，当前传入：{indent_chars}")
            return False

        # 步骤1：获取段落实际有效字体大小（pt）——与原有获取函数逻辑完全一致
        font_size_pt = default_font_size_pt
        # 优先级1：从段落Run中取第一个有效字体大小（遍历所有Run）
        if paragraph.runs and len(paragraph.runs) > 0:
            for run in paragraph.runs:
                if run.font and run.font.size:
                    if hasattr(run.font.size, "pt"):
                        current_pt = run.font.size.pt
                    else:
                        try:
                            current_pt = float(run.font.size)
                        except (ValueError, TypeError):
                            continue
                    if current_pt and current_pt > 0:
                        font_size_pt = current_pt
                        break
        # 优先级2：Run无有效值时，从段落样式中获取
        if (
            font_size_pt == default_font_size_pt
            and hasattr(paragraph, "style")
            and paragraph.style
        ):
            try:
                style_font = paragraph.style.font
                if hasattr(style_font, "size") and style_font.size:
                    if hasattr(style_font.size, "pt"):
                        style_pt = style_font.size.pt
                    else:
                        style_pt = float(style_font.size)
                    if style_pt and style_pt > 0:
                        font_size_pt = style_pt
            except AttributeError:
                pass

        # 步骤2：核心计算——自定义字符数 → 对应pt值（与获取函数互逆）
        # 原有获取逻辑：字符数 = 缩进pt / 字体pt → 反向推导：缩进pt = 字符数 × 字体pt
        target_indent_pt = font_size_pt * indent_chars

        # 步骤3：设置段落首行缩进（python-docx原生支持pt值直接赋值）
        paragraph.paragraph_format.first_line_indent = Pt(int(target_indent_pt))

        return True
    except Exception as e:
        logger.error(f"设置段落{indent_chars}字符首行缩进失败: {e}")
        return False
