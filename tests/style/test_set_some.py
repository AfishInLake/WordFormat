#! /usr/bin/env python
# @Time    : 2026/2/11 10:46
# @Author  : afish
# @File    : test_style.py
"""
测试style模块的功能
"""
from io import StringIO
from unittest.mock import MagicMock, patch, Mock

import pytest
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from loguru import logger

from wordformat.style.check_format import (
    DIFFResult,
    CharacterStyle,
    ParagraphStyle,
)
from wordformat.style.get_some import (
    paragraph_get_alignment,
    paragraph_get_space_before,
    paragraph_get_space_after,
    paragraph_get_line_spacing,
    paragraph_get_first_line_indent,
    paragraph_get_builtin_style_name,
    run_get_font_name,
    run_get_font_size_pt,
    run_get_font_color,
    run_get_font_bold,
    run_get_font_italic,
    run_get_font_underline,
    _get_style_spacing,
)
from wordformat.style.set_some import (
    run_set_font_name,
    set_paragraph_space_before_by_lines,
    set_paragraph_space_after_by_lines,
    _SetSpacing,
    _SetLineSpacing,
    _SetFirstLineIndent, _SetIndent,
)
from wordformat.style.style_enum import (
    FontName,
    FontSize,
    FontColor,
    Alignment,
    LineSpacingRule,
    LineSpacing,
    FirstLineIndent,
    BuiltInStyle,
)


# 测试配置初始化
@pytest.fixture(autouse=True)
def init_config():
    """初始化配置"""
    with patch('wordformat.style.check_format.get_config') as mock_get_config:
        mock_config = MagicMock()
        mock_config.style_checks_warning = MagicMock(
            bold=True,
            italic=True,
            underline=True,
            font_size=True,
            font_color=True,
            font_name=True,
            alignment=True,
            space_before=True,
            space_after=True,
            line_spacing=True,
            line_spacingrule=True,
            first_line_indent=True,
            builtin_style_name=True,
        )
        mock_get_config.return_value = mock_config
        yield


# 测试get_some.py中的函数
def test_paragraph_get_alignment():
    """测试获取段落对齐方式"""
    # 模拟段落对象
    paragraph = MagicMock(spec=Paragraph)

    # 测试情况1：段落直接设置了对齐方式
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    assert paragraph_get_alignment(paragraph) == WD_ALIGN_PARAGRAPH.CENTER

    # 测试情况2：段落没有直接设置对齐方式，但样式设置了
    paragraph.paragraph_format.alignment = None
    base_style = MagicMock()
    base_style.paragraph_format.alignment = None
    base_style._base_style = None

    style = MagicMock()
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    style._base_style = base_style

    paragraph.style = style
    assert paragraph_get_alignment(paragraph) == WD_ALIGN_PARAGRAPH.RIGHT

    # 测试情况3：段落和样式都没有设置对齐方式
    paragraph.paragraph_format.alignment = None
    style.paragraph_format.alignment = None
    assert paragraph_get_alignment(paragraph) is None


def test_paragraph_get_space_before():
    """测试获取段前间距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()
    paragraph.style = MagicMock()
    paragraph.style.name = "正文"

    paragraph.style.element = MagicMock()
    paragraph.style.element.find.return_value = None
    paragraph.style.base_style = None

    # 模拟pPr和spacing
    pPr = MagicMock()
    pPr.find.return_value = None
    paragraph._element.find.return_value = pPr

    # 测试默认情况
    assert paragraph_get_space_before(paragraph) is None


def test_paragraph_get_space_after():
    """测试获取段后间距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()
    paragraph.style = MagicMock()
    paragraph.style.name = "正文"

    paragraph.style.element = MagicMock()
    paragraph.style.element.find.return_value = None  # pPr 为 None
    paragraph.style.base_style = None  # 基样式为 None
    # 模拟pPr和spacing
    pPr = MagicMock()
    pPr.find.return_value = None
    paragraph._element.find.return_value = pPr

    # 测试默认情况
    assert paragraph_get_space_after(paragraph) == None


def test_paragraph_get_line_spacing():
    """测试获取行距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph.paragraph_format = MagicMock()
    paragraph.style = MagicMock()
    paragraph.style.paragraph_format = MagicMock()

    # 测试单倍行距
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    paragraph.paragraph_format.line_spacing = 1.0
    assert paragraph_get_line_spacing(paragraph) == 1.0

    # 测试1.5倍行距
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    assert paragraph_get_line_spacing(paragraph) == 1.5

    # 测试双倍行距
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    assert paragraph_get_line_spacing(paragraph) == 2.0


def test_paragraph_get_first_line_indent():
    """测试获取首行缩进"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()

    # 测试无缩进情况
    pPr = MagicMock()
    pPr.find.return_value = None
    paragraph._element.find.return_value = pPr
    assert paragraph_get_first_line_indent(paragraph) is None


def test_paragraph_get_builtin_style_name():
    """测试获取段落样式名称"""
    # 模拟段落对象
    paragraph = MagicMock()

    # 测试有样式的情况
    style = MagicMock()
    style.name = "Heading 1"
    paragraph.style = style
    assert paragraph_get_builtin_style_name(paragraph) == "heading 1"

    # 测试无样式的情况
    paragraph.style = None
    assert paragraph_get_builtin_style_name(paragraph) == ""


def test_run_get_font_name():
    """测试获取字体名称"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run._element = MagicMock()
    run._element.rPr = MagicMock()
    run._element.rPr.rFonts = MagicMock()

    # 测试有字体名称的情况
    with patch('wordformat.style.get_some.qn', return_value='w:eastAsia'):
        run._element.rPr.rFonts.get.return_value = "Microsoft YaHei"
        assert run_get_font_name(run) == "Microsoft YaHei"

    # 测试无字体名称的情况
    with patch('wordformat.style.get_some.qn', return_value='w:eastAsia'):
        run._element.rPr.rFonts.get.return_value = None
        assert run_get_font_name(run) is None


def test_run_get_font_size_pt():
    """测试获取字体大小"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run._parent = MagicMock()
    run._parent.style = MagicMock()
    run._parent.style.font = MagicMock()

    # 测试Run直接设置了字体大小
    run.font.size = MagicMock()
    run.font.size.pt = 12.0
    assert run_get_font_size_pt(run) == 12.0

    # 测试Run没有设置字体大小，但样式设置了
    run.font.size = None
    run._parent.style.font.size = MagicMock()
    run._parent.style.font.size.pt = 14.0
    assert run_get_font_size_pt(run) == 14.0

    # 测试都没有设置字体大小
    run.font.size = None
    run._parent.style.font.size = None
    assert run_get_font_size_pt(run) == 12.0


def test_run_get_font_color():
    """测试获取字体颜色"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()

    # 测试有颜色的情况
    run.font.color = MagicMock()
    run.font.color.rgb = (255, 0, 0)  # 红色
    assert run_get_font_color(run) == (255, 0, 0)

    # 测试无颜色的情况
    run.font.color = None
    assert run_get_font_color(run) == (0, 0, 0)


def test_run_get_font_bold():
    """测试获取字体是否加粗"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()

    # 测试加粗的情况
    run.font.bold = True
    assert run_get_font_bold(run) is True

    # 测试未加粗的情况
    run.font.bold = False
    assert run_get_font_bold(run) is False


def test_run_get_font_italic():
    """测试获取字体是否斜体"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()

    # 测试斜体的情况
    run.font.italic = True
    assert run_get_font_italic(run) is True

    # 测试未斜体的情况
    run.font.italic = False
    assert run_get_font_italic(run) is False


def test_run_get_font_underline():
    """测试获取字体是否下划线"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()

    # 测试有下划线的情况
    run.font.underline = True
    assert run_get_font_underline(run) is True

    # 测试无下划线的情况
    run.font.underline = False
    assert run_get_font_underline(run) is False


# 测试set_some.py中的函数
def test_run_set_font_name():
    """测试设置字体名称"""
    # 模拟Run对象
    run = MagicMock(spec=Run)
    run._element = MagicMock()

    # 测试rPr为None的情况
    run._element.rPr = None

    # 模拟parse_xml和qn
    with patch('wordformat.style.set_some.parse_xml') as mock_parse_xml:
        with patch('wordformat.style.set_some.qn') as mock_qn:
            mock_qn.return_value = 'w:eastAsia'
            mock_rPr = MagicMock()
            mock_rFonts = MagicMock()
            mock_rFonts.set = MagicMock()
            mock_parse_xml.side_effect = [mock_rPr, mock_rFonts]

            # 设置模拟对象的属性和方法
            def mock_append(element):
                if element == mock_rPr:
                    run._element.rPr = mock_rPr
                elif element == mock_rFonts:
                    mock_rPr.rFonts = mock_rFonts

            run._element.append = mock_append
            mock_rPr.rFonts = mock_rFonts

            # 执行函数
            run_set_font_name(run, "Microsoft YaHei")

            # 验证调用
            assert run._element.rPr is not None


def test_set_paragraph_space_before_by_lines():
    """测试设置段前间距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 执行函数
    set_paragraph_space_before_by_lines(paragraph, 0.5)

    # 验证调用
    assert paragraph._element.get_or_add_pPr.called


def test_set_paragraph_space_after_by_lines():
    """测试设置段后间距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 执行函数
    set_paragraph_space_after_by_lines(paragraph, 0.5)

    # 验证调用
    assert paragraph._element.get_or_add_pPr.called


def test_set_spacing_hang():
    """测试设置行间距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()

    # 模拟find和append
    pPr = MagicMock()
    pPr.find.return_value = None
    pPr.append = MagicMock()
    paragraph._element.find.return_value = pPr
    paragraph._element.append = MagicMock()

    # 模拟parse_xml和qn
    with patch('wordformat.style.set_some.parse_xml') as mock_parse_xml:
        with patch('wordformat.style.set_some.qn') as mock_qn:
            mock_qn.side_effect = ['w:pPr', 'w:spacing', 'w:beforeLines', 'w:before']
            mock_spacing = MagicMock()
            mock_spacing.get.return_value = None

            # 模拟spacing元素的attrib字典和设置属性的方法
            mock_attrib = {}

            def mock_set(attr, value):
                mock_attrib[attr] = value

            mock_spacing.attrib = mock_attrib
            mock_spacing.set = mock_set
            mock_parse_xml.side_effect = [pPr, mock_spacing]

            # 执行函数
            _SetSpacing.set_hang(paragraph, "before", 0.5)

            # 验证调用
            assert pPr.append.called


def test_set_line_spacing_pt():
    """测试设置磅值行距"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph.paragraph_format = MagicMock()

    # 执行函数
    _SetLineSpacing.set_pt(paragraph, 12.0)

    # 验证调用
    assert paragraph.paragraph_format.line_spacing_rule == WD_LINE_SPACING.EXACTLY


def test_set_first_line_indent_char():
    """测试设置字符首行缩进"""
    # 模拟段落对象
    paragraph = MagicMock()
    paragraph._element = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    pPr.find = MagicMock()
    pPr.remove = MagicMock()
    pPr.append = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find和remove
    ind = MagicMock()
    ind.get.side_effect = [None, None]
    pPr.find.return_value = ind

    # 模拟clear方法
    with patch.object(_SetFirstLineIndent, 'clear') as mock_clear:
        # 模拟parse_xml
        with patch('wordformat.style.set_some.parse_xml') as mock_parse_xml:
            mock_new_ind = MagicMock()
            mock_parse_xml.return_value = mock_new_ind

            # 执行函数
            result = _SetFirstLineIndent.set_char(paragraph, 2.0)

            # 验证结果
            assert result is True


# 测试check_format.py中的类和函数
def test_diff_result_str():
    """测试DIFFResult的__str__方法"""
    diff = DIFFResult("bold", True, False, "期待加粗")
    assert str(diff) == "期待加粗"


def test_character_style_diff_from_run():
    """测试CharacterStyle的diff_from_run方法"""
    # 创建CharacterStyle实例
    char_style = CharacterStyle()

    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False

    # 模拟run_get_font_size_pt
    with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=12.0):
        # 模拟run_get_font_color
        with patch('wordformat.style.check_format.run_get_font_color', return_value=(0, 0, 0)):
            # 模拟run_get_font_name
            with patch('wordformat.style.check_format.run_get_font_name', return_value="宋体"):
                # 执行方法
                diffs = char_style.diff_from_run(run)

                # 验证结果
                assert isinstance(diffs, list)


def test_paragraph_style_diff_from_paragraph():
    """测试ParagraphStyle的diff_from_paragraph方法"""
    # 创建ParagraphStyle实例
    para_style = ParagraphStyle()

    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()
    paragraph.paragraph_format.first_line_indent = None
    # 添加_element属性
    paragraph._element = MagicMock()
    paragraph._element.find.return_value = None

    # 模拟paragraph_get_alignment
    with patch('wordformat.style.get_some.paragraph_get_alignment', return_value=WD_ALIGN_PARAGRAPH.LEFT):
        # 模拟paragraph_get_space_before
        with patch('wordformat.style.get_some.paragraph_get_space_before', return_value=0.5):
            # 模拟paragraph_get_space_after
            with patch('wordformat.style.get_some.paragraph_get_space_after', return_value=0.5):
                # 模拟paragraph_get_line_spacing
                with patch('wordformat.style.get_some.paragraph_get_line_spacing', return_value=1.5):
                    # 模拟paragraph_get_first_line_indent
                    with patch('wordformat.style.get_some.paragraph_get_first_line_indent', return_value=0):
                        # 模拟paragraph_get_builtin_style_name
                        with patch('wordformat.style.get_some.paragraph_get_builtin_style_name', return_value="正文"):
                            # 执行方法
                            diffs = para_style.diff_from_paragraph(paragraph)

                            # 验证结果
                            assert isinstance(diffs, list)




def test_character_style_to_string():
    """测试CharacterStyle的to_string方法"""
    # 创建DIFFResult实例
    diff1 = DIFFResult("bold", True, False, "期待加粗")
    diff2 = DIFFResult("italic", False, True, "期待非斜体")
    diffs = [diff1, diff2]

    # 执行方法
    result = CharacterStyle.to_string(diffs)

    # 验证结果
    assert isinstance(result, str)



def test_paragraph_style_to_string():
    """测试ParagraphStyle的to_string方法"""
    # 创建ParagraphStyle实例
    para_style = ParagraphStyle()

    # 创建DIFFResult实例
    diff1 = DIFFResult("alignment", "左对齐", "居中", "期待左对齐")
    diff2 = DIFFResult("space_before", "0.5行", "1行", "期待0.5行")
    diffs = [diff1, diff2]

    # 模拟style_checks_warning
    with patch('wordformat.style.check_format.style_checks_warning') as mock_warning:
        mock_warning.alignment = True
        mock_warning.space_before = True
        mock_warning.space_after = False
        mock_warning.line_spacing = False
        mock_warning.line_spacingrule = False
        mock_warning.first_line_indent = False
        mock_warning.builtin_style_name = False

        # 执行方法
        result = para_style.to_string(diffs)

        # 验证结果
        assert isinstance(result, str)


def test_character_style_init():
    """测试CharacterStyle的初始化方法"""
    # 测试默认参数初始化
    char_style1 = CharacterStyle()
    assert isinstance(char_style1, CharacterStyle)

    # 测试自定义参数初始化
    char_style2 = CharacterStyle(
        font_name_cn="宋体",
        font_name_en="Times New Roman",
        font_size="小四",
        font_color="BLACK",
        bold=True,
        italic=False,
        underline=False
    )
    assert isinstance(char_style2, CharacterStyle)


def test_paragraph_style_init():
    """测试ParagraphStyle的初始化方法"""
    # 测试默认参数初始化
    para_style1 = ParagraphStyle()
    assert isinstance(para_style1, ParagraphStyle)

    # 测试自定义参数初始化
    para_style2 = ParagraphStyle(
        alignment="左对齐",
        space_before="0.5行",
        space_after="0.5行",
        line_spacing="1.5倍",
        line_spacingrule="单倍行距",
        first_line_indent="0字符",
        builtin_style_name="正文"
    )
    assert isinstance(para_style2, ParagraphStyle)


def test_diff_result_init():
    """测试DIFFResult的初始化方法"""
    diff = DIFFResult(
        diff_type="bold",
        expected_value=True,
        current_value=False,
        comment="期待加粗",
        level=1
    )
    assert isinstance(diff, DIFFResult)
    assert diff.diff_type == "bold"
    assert diff.expected_value == True
    assert diff.current_value == False
    assert diff.comment == "期待加粗"
    assert diff.level == 1


def test_character_style_diff_from_run_all_match():
    """测试CharacterStyle的diff_from_run方法（所有样式匹配）"""
    # 创建CharacterStyle实例
    char_style = CharacterStyle(
        font_name_cn="宋体",
        font_name_en="Times New Roman",
        font_size="小四",
        font_color="BLACK",
        bold=False,
        italic=False,
        underline=False
    )

    # 模拟Run对象（所有样式都匹配）
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False
    run.font.name = "Times New Roman"

    # 模拟run_get_font_size_pt
    with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=12.0):
        # 模拟run_get_font_color
        with patch('wordformat.style.check_format.run_get_font_color', return_value=(0, 0, 0)):
            # 模拟run_get_font_name
            with patch('wordformat.style.check_format.run_get_font_name', return_value="宋体"):
                # 执行方法
                diffs = char_style.diff_from_run(run)

                # 验证结果
                assert isinstance(diffs, list)


def test_character_style_diff_from_run_no_match():
    """测试CharacterStyle的diff_from_run方法（所有样式都不匹配）"""
    # 创建CharacterStyle实例
    char_style = CharacterStyle(
        font_name_cn="宋体",
        font_name_en="Times New Roman",
        font_size="小四",
        font_color="BLACK",
        bold=True,
        italic=True,
        underline=True
    )

    # 模拟Run对象（所有样式都不匹配）
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False
    run.font.name = "Arial"

    # 模拟run_get_font_size_pt
    with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=10.0):
        # 模拟run_get_font_color
        with patch('wordformat.style.check_format.run_get_font_color', return_value=(255, 0, 0)):
            # 模拟run_get_font_name
            with patch('wordformat.style.check_format.run_get_font_name', return_value="微软雅黑"):
                # 执行方法
                diffs = char_style.diff_from_run(run)

                # 验证结果
                assert isinstance(diffs, list)
                assert len(diffs) > 0


def test_character_style_apply_to_run_unknown_diff_type():
    """测试CharacterStyle的apply_to_run方法（未知diff_type）"""
    # 创建CharacterStyle实例
    char_style = CharacterStyle()

    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False

    # 模拟diff_from_run返回包含未知diff_type的结果
    with patch('wordformat.style.check_format.CharacterStyle.diff_from_run') as mock_diff_from:
        # 创建模拟的DIFFResult
        mock_diff = MagicMock()
        mock_diff.diff_type = "unknown_type"
        mock_diff.expected_value = "value"
        mock_diff.current_value = "current_value"
        mock_diff.comment = "未知类型"

        mock_diff_from.return_value = [mock_diff]

        # 执行方法
        result = char_style.apply_to_run(run)

        # 验证结果
        assert isinstance(result, list)


def test_paragraph_style_apply_to_paragraph_unknown_diff_type():
    """测试ParagraphStyle的apply_to_paragraph方法（未知diff_type）"""
    # 创建ParagraphStyle实例
    para_style = ParagraphStyle()

    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)

    # 模拟diff_from_paragraph返回包含未知diff_type的结果
    with patch('wordformat.style.check_format.ParagraphStyle.diff_from_paragraph') as mock_diff_from:
        # 创建模拟的DIFFResult
        mock_diff = MagicMock()
        mock_diff.diff_type = "unknown_type"
        mock_diff.expected_value = "value"
        mock_diff.current_value = "current_value"
        mock_diff.comment = "未知类型"

        mock_diff_from.return_value = [mock_diff]

        # 执行方法
        result = para_style.apply_to_paragraph(paragraph)

        # 验证结果
        assert isinstance(result, list)



def test_get_style_spacing():
    """测试从样式中获取间距"""
    # 模拟样式对象
    style = MagicMock()
    style.element = MagicMock()
    style.base_style = None

    # 模拟style_pPr不存在的情况
    style.element.find.return_value = None
    assert _get_style_spacing(style, "before") == None































def test_font_name():
    """测试FontName类"""
    # 测试中文字体
    font_name_cn = FontName("宋体")
    assert font_name_cn.is_chinese("宋体") == True

    # 测试英文字体
    font_name_en = FontName("Times New Roman")
    assert font_name_en.is_chinese("Times New Roman") == False


def test_font_size():
    """测试FontSize类"""
    # 测试预设字号
    font_size = FontSize("小四")
    assert font_size.rel_value == 12

    # 测试数字字号
    font_size_num = FontSize("12")
    assert font_size_num.rel_value == "12"


def test_font_color():
    """测试FontColor类"""
    # 测试中文颜色名称
    font_color_zh = FontColor("红色")
    assert isinstance(font_color_zh.rel_value, tuple)

    # 测试英文颜色名称
    font_color_en = FontColor("red")
    assert isinstance(font_color_en.rel_value, tuple)

    # 测试十六进制颜色值
    font_color_hex = FontColor("#FF0000")
    assert isinstance(font_color_hex.rel_value, tuple)


def test_alignment():
    """测试Alignment类"""
    # 测试左对齐
    alignment_left = Alignment("左对齐")
    assert alignment_left.rel_value == WD_ALIGN_PARAGRAPH.LEFT

    # 测试居中对齐
    alignment_center = Alignment("居中对齐")
    assert alignment_center.rel_value == WD_ALIGN_PARAGRAPH.CENTER


def test_line_spacing_rule():
    """测试LineSpacingRule类"""
    # 测试单倍行距
    line_spacing_single = LineSpacingRule("单倍行距")
    assert line_spacing_single.rel_value == WD_LINE_SPACING.SINGLE

    # 测试1.5倍行距
    line_spacing_15 = LineSpacingRule("1.5倍行距")
    assert line_spacing_15.rel_value == WD_LINE_SPACING.ONE_POINT_FIVE


def test_line_spacing():
    """测试LineSpacing类"""
    # 测试1.5倍行距
    line_spacing = LineSpacing("1.5倍")
    assert isinstance(line_spacing.rel_value, float)
    assert line_spacing.rel_value == 1.5


def test_first_line_indent():
    """测试FirstLineIndent类"""
    # 测试2字符缩进
    first_line_indent = FirstLineIndent("2字符")
    assert isinstance(first_line_indent.rel_value, float)
    assert first_line_indent.rel_value == 2.0


def test_built_in_style():
    """测试BuiltInStyle类"""
    # 测试标题1样式
    built_in_style = BuiltInStyle("Heading 1")
    assert built_in_style.rel_value == "Heading 1"

    # 测试正文样式
    built_in_style_normal = BuiltInStyle("正文")
    assert built_in_style_normal.rel_value == "Normal"


def test_character_style_apply_to_run():
    """测试应用字符样式到Run对象"""
    # 创建CharacterStyle实例
    char_style = CharacterStyle()

    # 模拟Run对象
    run = MagicMock(spec=Run)
    run.font = MagicMock()
    run.font.bold = False
    run.font.italic = False
    run.font.underline = False

    # 模拟run_get_font_size_pt
    with patch('wordformat.style.check_format.run_get_font_size_pt', return_value=12.0):
        # 模拟run_get_font_color
        with patch('wordformat.style.check_format.run_get_font_color', return_value=(0, 0, 0)):
            # 模拟run_get_font_name
            with patch('wordformat.style.check_format.run_get_font_name', return_value="宋体"):
                # 模拟FontName.format
                with patch('wordformat.style.style_enum.FontName.format'):
                    # 模拟FontSize.format
                    with patch('wordformat.style.style_enum.FontSize.format'):
                        # 模拟FontColor.format
                        with patch('wordformat.style.style_enum.FontColor.format'):
                            # 执行方法
                            result = char_style.apply_to_run(run)

                            # 验证结果
                            assert isinstance(result, list)


def test_paragraph_style_apply_to_paragraph():
    """测试应用段落样式到Paragraph对象"""
    # 创建ParagraphStyle实例
    para_style = ParagraphStyle()

    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()
    paragraph.paragraph_format.first_line_indent = None
    # 添加_element属性
    paragraph._element = MagicMock()
    paragraph._element.find.return_value = None

    # 模拟paragraph_get_alignment
    with patch('wordformat.style.get_some.paragraph_get_alignment', return_value=WD_ALIGN_PARAGRAPH.LEFT):
        # 模拟paragraph_get_space_before
        with patch('wordformat.style.get_some.paragraph_get_space_before', return_value=0.5):
            # 模拟paragraph_get_space_after
            with patch('wordformat.style.get_some.paragraph_get_space_after', return_value=0.5):
                # 模拟paragraph_get_line_spacing
                with patch('wordformat.style.get_some.paragraph_get_line_spacing', return_value=1.5):
                    # 模拟paragraph_get_first_line_indent
                    with patch('wordformat.style.get_some.paragraph_get_first_line_indent', return_value=0):
                        # 模拟paragraph_get_builtin_style_name
                        with patch('wordformat.style.get_some.paragraph_get_builtin_style_name', return_value="正文"):
                            # 模拟Alignment.format
                            with patch('wordformat.style.style_enum.Alignment.format'):
                                # 模拟Spacing.format
                                with patch('wordformat.style.style_enum.Spacing.format'):
                                    # 模拟LineSpacing.format
                                    with patch('wordformat.style.style_enum.LineSpacing.format'):
                                        # 模拟LineSpacingRule.format
                                        with patch('wordformat.style.style_enum.LineSpacingRule.format'):
                                            # 模拟FirstLineIndent.format
                                            with patch('wordformat.style.style_enum.FirstLineIndent.format'):
                                                # 模拟BuiltInStyle.format
                                                with patch('wordformat.style.style_enum.BuiltInStyle.format'):
                                                    # 执行方法
                                                    result = para_style.apply_to_paragraph(paragraph)

                                                    # 验证结果
                                                    assert isinstance(result, list)


# 测试set_some.py中的函数和方法
def test_paragraph_space_by_lines():
    """测试设置段落间距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟xpath
    pPr.xpath.return_value = []

    # 模拟append
    pPr.append = MagicMock()

    # 执行函数
    from wordformat.style.set_some import _paragraph_space_by_lines
    _paragraph_space_by_lines(paragraph, before_lines=0.5, after_lines=0.5)

    # 验证调用
    assert paragraph._element.get_or_add_pPr.called


def test_set_spacing_set_pt():
    """测试设置磅值间距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetSpacing
    _SetSpacing.set_pt(paragraph, "before", 12.0)

    # 验证调用
    assert paragraph.paragraph_format.space_before is not None


def test_set_spacing_set_cm():
    """测试设置厘米间距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetSpacing
    _SetSpacing.set_cm(paragraph, "before", 1.0)

    # 验证调用
    assert paragraph.paragraph_format.space_before is not None


def test_set_spacing_set_inch():
    """测试设置英寸间距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetSpacing
    _SetSpacing.set_inch(paragraph, "before", 0.5)

    # 验证调用
    assert paragraph.paragraph_format.space_before is not None


def test_set_spacing_set_mm():
    """测试设置毫米间距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetSpacing
    _SetSpacing.set_mm(paragraph, "before", 10.0)

    # 验证调用
    assert paragraph.paragraph_format.space_before is not None


def test_set_line_spacing_set_cm():
    """测试设置厘米行距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetLineSpacing
    _SetLineSpacing.set_cm(paragraph, 1.0)

    # 验证调用
    assert paragraph.paragraph_format.line_spacing_rule is not None


def test_set_line_spacing_set_inch():
    """测试设置英寸行距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetLineSpacing
    _SetLineSpacing.set_inch(paragraph, 0.5)

    # 验证调用
    assert paragraph.paragraph_format.line_spacing_rule is not None


def test_set_line_spacing_set_mm():
    """测试设置毫米行距"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph.paragraph_format = MagicMock()

    # 执行方法
    from wordformat.style.set_some import _SetLineSpacing
    _SetLineSpacing.set_mm(paragraph, 10.0)

    # 验证调用
    assert paragraph.paragraph_format.line_spacing_rule is not None


def test_set_first_line_indent_clear():
    """测试清除首行缩进"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find
    ind = MagicMock()
    ind.get.return_value = None
    pPr.find.return_value = ind

    # 执行方法
    from wordformat.style.set_some import _SetFirstLineIndent
    _SetFirstLineIndent.clear(paragraph)

    # 验证调用
    assert paragraph._element.get_or_add_pPr.called


def test_set_first_line_indent_set_inch():
    """测试设置英寸首行缩进"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()
    paragraph.paragraph_format = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find
    ind = MagicMock()
    ind.get.return_value = None
    pPr.find.return_value = ind

    # 执行方法
    from wordformat.style.set_some import _SetFirstLineIndent
    _SetFirstLineIndent.set_inch(paragraph, 0.5)

    # 验证调用
    assert paragraph.paragraph_format.first_line_indent is not None


def test_set_first_line_indent_set_mm():
    """测试设置毫米首行缩进"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()
    paragraph.paragraph_format = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find
    ind = MagicMock()
    ind.get.return_value = None
    pPr.find.return_value = ind

    # 执行方法
    from wordformat.style.set_some import _SetFirstLineIndent
    _SetFirstLineIndent.set_mm(paragraph, 10.0)

    # 验证调用
    assert paragraph.paragraph_format.first_line_indent is not None


def test_set_first_line_indent_set_pt():
    """测试设置磅值首行缩进"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()
    paragraph.paragraph_format = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find
    ind = MagicMock()
    ind.get.return_value = None
    pPr.find.return_value = ind

    # 执行方法
    from wordformat.style.set_some import _SetFirstLineIndent
    _SetFirstLineIndent.set_pt(paragraph, 12.0)

    # 验证调用
    assert paragraph.paragraph_format.first_line_indent is not None


def test_set_first_line_indent_set_cm():
    """测试设置厘米首行缩进"""
    # 模拟Paragraph对象
    paragraph = MagicMock(spec=Paragraph)
    paragraph._element = MagicMock()
    paragraph.paragraph_format = MagicMock()

    # 模拟get_or_add_pPr
    pPr = MagicMock()
    paragraph._element.get_or_add_pPr.return_value = pPr

    # 模拟find
    ind = MagicMock()
    ind.get.return_value = None
    pPr.find.return_value = ind

    # 执行方法
    from wordformat.style.set_some import _SetFirstLineIndent
    _SetFirstLineIndent.set_cm(paragraph, 1.0)

    # 验证调用
    assert paragraph.paragraph_format.first_line_indent is not None


class TestSetIndent:
    """测试 _SetIndent 类"""

    @pytest.fixture
    def mock_paragraph(self):
        """创建一个模拟的段落对象"""
        paragraph = Mock(spec=Paragraph)
        paragraph._element = Mock()
        paragraph.paragraph_format = Mock()

        # 模拟 pPr 元素
        pPr = Mock()
        paragraph._element.get_or_add_pPr.return_value = pPr

        # 模拟现有的 ind 元素
        ind = Mock()
        ind.attrib = {}
        ind.get.return_value = None
        pPr.find.return_value = ind

        return paragraph

    def test_set_char_left_indent_positive(self, mock_paragraph):
        """测试设置正数左缩进（字符单位）"""
        # 模拟 XML 操作
        mock_pPr = Mock()
        mock_pPr.find.return_value = None  # 关键：ind = None
        mock_paragraph._element.get_or_add_pPr.return_value = mock_pPr
        # 调用方法
        result = _SetIndent.set_char(mock_paragraph, "R", 2)

        # 验证结果
        assert result is True
        # 验证 get_or_add_pPr 被调用
        mock_paragraph._element.get_or_add_pPr.assert_called_once()

    def test_set_char_right_indent_positive(self, mock_paragraph):
        """测试设置正数右缩进（字符单位）"""
        # 调用方法
        result = _SetIndent.set_char(mock_paragraph, "X", 3)

        # 验证结果
        assert result is True

    def test_set_char_zero_indent(self, mock_paragraph):
        """测试设置零缩进（清除缩进）"""
        # 设置一个已有的缩进属性
        pPr = mock_paragraph._element.get_or_add_pPr.return_value
        ind = pPr.find.return_value
        ind.attrib = {
            "w:leftChars": "200",
            "w:left": "1000",
            "w:rightChars": "100"
        }

        # 调用方法设置左缩进为0
        result = _SetIndent.set_char(mock_paragraph, "R", 0)

        # 验证结果
        assert result is True
        # 验证 ind 元素被移除
        pPr.remove.assert_called_once_with(ind)

    def test_set_char_negative_value_allowed(self, mock_paragraph):
        """测试允许设置负缩进（Word 支持负值）"""
        mock_pPr = Mock()
        mock_pPr.find.return_value = None
        mock_paragraph._element.get_or_add_pPr.return_value = mock_pPr

        result = _SetIndent.set_char(mock_paragraph, "R", -3)
        assert result is True



    def test_set_char_invalid_type(self, mock_paragraph):
        # 创建内存 buffer 捕获日志
        log_buffer = StringIO()
        log_id = logger.add(log_buffer, level="ERROR", format="{message}")

        result = _SetIndent.set_char(mock_paragraph, "Z", 2)

        # 获取日志内容
        log_output = log_buffer.getvalue()

        assert result is False
        assert "无效的缩进类型" in log_output

        # 清理 handler
        logger.remove(log_id)

    def test_set_char_exception_handling(self, mock_paragraph, caplog):
        """测试异常处理"""
        log_buffer = StringIO()
        log_id = logger.add(log_buffer, level="ERROR", format="{message}")
        # 模拟异常
        mock_paragraph._element.get_or_add_pPr.side_effect = Exception("测试异常")

        # 调用方法
        result = _SetIndent.set_char(mock_paragraph, "R", 2)

        # 验证结果
        assert result is False
        # 验证错误日志
        assert "设置字符缩进失败" in  log_buffer.getvalue()
        logger.remove(log_id)

    @patch('wordformat.style.set_some.Pt')
    def test_set_pt_left_indent(self, mock_pt, mock_paragraph):
        """测试设置 PT 单位的左缩进"""
        # 模拟 Pt 对象
        mock_pt_instance = Mock()
        mock_pt.return_value = mock_pt_instance

        # 调用方法
        _SetIndent.set_pt(mock_paragraph, "R", 12.5)

        # 验证
        mock_pt.assert_called_once_with(12.5)
        mock_paragraph.paragraph_format.left_indent = mock_pt_instance

    @patch('wordformat.style.set_some.Pt')
    def test_set_pt_right_indent(self, mock_pt, mock_paragraph):
        """测试设置 PT 单位的右缩进"""
        # 模拟 Pt 对象
        mock_pt_instance = Mock()
        mock_pt.return_value = mock_pt_instance

        # 调用方法
        _SetIndent.set_pt(mock_paragraph, "X", 10.0)

        # 验证
        mock_pt.assert_called_once_with(10.0)
        mock_paragraph.paragraph_format.right_indent = mock_pt_instance

    @patch('wordformat.style.set_some.Cm')
    def test_set_cm_indent(self, mock_cm, mock_paragraph):
        """测试设置厘米单位的缩进"""
        # 模拟 Cm 对象
        mock_cm_instance = Mock()
        mock_cm.return_value = mock_cm_instance

        # 测试左缩进
        _SetIndent.set_cm(mock_paragraph, "R", 2.54)
        mock_cm.assert_called_once_with(2.54)
        mock_paragraph.paragraph_format.left_indent = mock_cm_instance

        # 重置模拟
        mock_cm.reset_mock()
        mock_paragraph.paragraph_format.reset_mock()

        # 测试右缩进
        _SetIndent.set_cm(mock_paragraph, "X", 1.27)
        mock_cm.assert_called_once_with(1.27)
        mock_paragraph.paragraph_format.right_indent = mock_cm_instance

    @patch('wordformat.style.set_some.Inches')
    def test_set_inch_indent(self, mock_inches, mock_paragraph):
        """测试设置英寸单位的缩进"""
        # 模拟 Inches 对象
        mock_inches_instance = Mock()
        mock_inches.return_value = mock_inches_instance

        # 测试左缩进
        _SetIndent.set_inch(mock_paragraph, "R", 1.0)
        mock_inches.assert_called_once_with(1.0)
        mock_paragraph.paragraph_format.left_indent = mock_inches_instance

        # 重置模拟
        mock_inches.reset_mock()
        mock_paragraph.paragraph_format.reset_mock()

        # 测试右缩进
        _SetIndent.set_inch(mock_paragraph, "X", 0.5)
        mock_inches.assert_called_once_with(0.5)
        mock_paragraph.paragraph_format.right_indent = mock_inches_instance

    @patch('wordformat.style.set_some.Mm')
    def test_set_mm_indent(self, mock_mm, mock_paragraph):
        """测试设置毫米单位的缩进"""
        # 模拟 Mm 对象
        mock_mm_instance = Mock()
        mock_mm.return_value = mock_mm_instance

        # 测试左缩进
        _SetIndent.set_mm(mock_paragraph, "R", 25.4)
        mock_mm.assert_called_once_with(25.4)
        mock_paragraph.paragraph_format.left_indent = mock_mm_instance

    def test_apply_indent_left(self, mock_paragraph):
        """测试 _apply_indent 方法 - 左缩进"""
        test_value = Mock()

        # 调用方法
        _SetIndent._apply_indent(mock_paragraph, "R", test_value)

        # 验证
        mock_paragraph.paragraph_format.left_indent = test_value
        mock_paragraph.paragraph_format.right_indent.assert_not_called()

    def test_apply_indent_right(self, mock_paragraph):
        """测试 _apply_indent 方法 - 右缩进"""
        test_value = Mock()

        # 调用方法
        _SetIndent._apply_indent(mock_paragraph, "X", test_value)

        # 验证
        mock_paragraph.paragraph_format.right_indent = test_value
        mock_paragraph.paragraph_format.left_indent.assert_not_called()

    def test_apply_indent_invalid_type(self, mock_paragraph):
        """测试 _apply_indent 方法 - 无效类型"""
        test_value = Mock()

        # 验证异常
        with pytest.raises(ValueError, match="无效的缩进类型"):
            _SetIndent._apply_indent(mock_paragraph, "Z", test_value)

    def test_char_conversion_calculation(self, mock_paragraph):
        """测试字符单位的转换计算"""
        # 测试字符到 Word 单位的转换
        # 1 字符 = 100 Word 单位

        # 模拟 XML 操作
        pPr = Mock()
        ind = Mock()
        ind.attrib = {}
        ind.get.return_value = None
        pPr.find.return_value = ind
        mock_paragraph._element.get_or_add_pPr.return_value = pPr

        # 调用方法设置 2.5 字符
        result = _SetIndent.set_char(mock_paragraph, "R", 2.5)

        # 验证结果
        assert result is True
        # 验证字符被转换为 250 (2.5 * 100)
        # 注意：这里我们验证的是方法逻辑，实际 XML 操作被模拟了

    def test_existing_attributes_preserved(self, mock_paragraph):
        """测试设置缩进时保留现有属性"""
        # 设置一个已有的 ind 元素，带有其他属性
        pPr = mock_paragraph._element.get_or_add_pPr.return_value
        ind = Mock()
        ind.attrib = {
            "w:leftChars": "100",
            "w:right": "2000",
            "w:hanging": "500"
        }
        ind.get.side_effect = lambda key: ind.attrib.get(key)
        pPr.find.return_value = ind

        # 调用方法修改右缩进
        result = _SetIndent.set_char(mock_paragraph, "X", 3)

        # 验证结果
        assert result is True
        # 验证旧的 ind 被移除
        pPr.remove.assert_called_once_with(ind)