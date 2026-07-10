"""文本工具：CJK 字符检测、编号文字、题注解析。"""

import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def get_paragraph_numbering_text(paragraph: Paragraph) -> str:
    """
    提取段落的自动编号文字（如"第一章"、"1.1"、"1."等）。

    Word 的自动编号存储在 numbering.xml 中，para.text 不包含编号文字。
    本函数从段落的 XML 中读取 numId 和 ilvl，查找对应的 lvlText 模板，
    然后根据当前编号计数器替换占位符，生成实际的编号文字。

    Args:
        paragraph: docx 段落对象

    Returns:
        编号文字字符串，无编号时返回空字符串
    """
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return ""
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return ""

    numId_elem = numPr.find(qn("w:numId"))
    ilvl_elem = numPr.find(qn("w:ilvl"))
    if numId_elem is None:
        return ""

    num_id = numId_elem.get(qn("w:val"))
    ilvl = int(ilvl_elem.get(qn("w:val"))) if ilvl_elem is not None else 0

    if num_id is None or num_id == "0":
        return ""

    # 获取 numbering part
    try:
        numbering_part = paragraph.part.numbering_part
    except (AttributeError, KeyError, NotImplementedError):
        return ""

    numbering_elm = numbering_part._element

    # 查找 num -> abstractNum -> lvl -> lvlText
    num_elem = None
    for el in numbering_elm.findall(qn("w:num")):
        if el.get(qn("w:numId")) == num_id:
            num_elem = el
            break
    if num_elem is None:
        return ""

    abstract_num_id_ref = num_elem.find(qn("w:abstractNumId"))
    if abstract_num_id_ref is None:
        return ""

    abstract_num_id = abstract_num_id_ref.get(qn("w:val"))

    abstract_num = None
    for el in numbering_elm.findall(qn("w:abstractNum")):
        if el.get(qn("w:abstractNumId")) == abstract_num_id:
            abstract_num = el
            break
    if abstract_num is None:
        return ""

    lvl = None
    for el in abstract_num.findall(qn("w:lvl")):
        if el.get(qn("w:ilvl")) == str(ilvl):
            lvl = el
            break
    if lvl is None:
        return ""

    lvl_text_elem = lvl.find(qn("w:lvlText"))
    if lvl_text_elem is None:
        return ""

    lvl_text = lvl_text_elem.get(qn("w:val"), "")
    num_fmt_elem = lvl.find(qn("w:numFmt"))
    num_fmt = (
        num_fmt_elem.get(qn("w:val"), "decimal")
        if num_fmt_elem is not None
        else "decimal"
    )

    # 计算当前级别的编号值
    # 需要遍历文档中所有使用同一 abstractNum 的段落来计数
    level_counters = _count_numbering_levels(numbering_elm, abstract_num_id, paragraph)

    current_num = level_counters.get(ilvl, 1)

    # 根据格式化类型转换数字
    _format_number(current_num, num_fmt)

    # 替换 lvlText 中的占位符
    # %1 -> 当前级别, %2 -> 下一级别, etc.
    result = lvl_text
    for lvl_idx, lvl_val in sorted(level_counters.items()):
        placeholder = f"%{lvl_idx + 1}"
        if placeholder in result:
            fmt_for_level = _get_level_fmt(abstract_num, lvl_idx)
            result = result.replace(placeholder, _format_number(lvl_val, fmt_for_level))

    return result


def _count_numbering_levels(
    numbering_elm, abstract_num_id: str, target_paragraph: Paragraph
) -> dict[int, int]:
    """
    遍历文档段落，计算目标段落所在编号链的各级计数器值。

    返回 {ilvl: count} 字典，如 {0: 1, 1: 2} 表示一级编号为1，二级编号为2
    """
    counters = {}
    target_element = target_paragraph._element

    # 找到所有引用同一 abstractNum 的 numId
    num_ids = set()
    for num_elem in numbering_elm.findall(qn("w:num")):
        abstract_num_id_ref = num_elem.find(qn("w:abstractNumId"))
        if (
            abstract_num_id_ref is not None
            and abstract_num_id_ref.get(qn("w:val")) == abstract_num_id
        ):
            num_ids.add(num_elem.get(qn("w:numId")))

    if not num_ids:
        return counters

    # 遍历文档 body 中的所有段落
    body = target_paragraph._element.getparent()  # body element
    if body is None:
        return counters

    # 获取 isRestart 标志
    def get_is_restart(lvl_element):
        """检查该级别是否在每个上级编号重启"""
        if lvl_element is None:
            return True
        restart = lvl_element.find(qn("w:isLgl") if False else qn("w:lvlRestart"))
        # lvlRestart 不存在时默认行为：下级在上级重启
        return restart is None

    for para_elm in body.findall(qn("w:p")):
        pPr = para_elm.find(qn("w:pPr"))
        if pPr is None:
            continue
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            continue
        numId_elem = numPr.find(qn("w:numId"))
        if numId_elem is None:
            continue
        para_num_id = numId_elem.get(qn("w:val"))
        if para_num_id not in num_ids:
            continue

        ilvl_elem = numPr.find(qn("w:ilvl"))
        para_ilvl = int(ilvl_elem.get(qn("w:val"))) if ilvl_elem is not None else 0

        # 增加当前级别计数
        counters[para_ilvl] = counters.get(para_ilvl, 0) + 1

        # 重置下级计数
        for key in list(counters.keys()):
            if key > para_ilvl:
                del counters[key]

        # 检查是否到达目标段落
        if para_elm is target_element:
            break

    return counters


def _format_number(num: int, num_fmt: str) -> str:
    """根据 numFmt 将数字格式化为对应字符串"""
    format_map = {
        "decimal": str,
        "upperRoman": lambda n: _to_roman(n).upper(),
        "lowerRoman": lambda n: _to_roman(n).lower(),
        "upperLetter": lambda n: chr(ord("A") + n - 1) if 1 <= n <= 26 else str(n),
        "lowerLetter": lambda n: chr(ord("a") + n - 1) if 1 <= n <= 26 else str(n),
        "chineseCountingThousand": lambda n: _to_chinese_num(n),
        "ideographTraditional": lambda n: _to_chinese_num(n),
        "chineseCounting": lambda n: _to_chinese_num(n),
    }
    formatter = format_map.get(num_fmt, str)
    return formatter(num)


def _get_level_fmt(abstract_num, ilvl: int) -> str:
    """获取指定级别的 numFmt"""
    lvl = None
    for el in abstract_num.findall(qn("w:lvl")):
        if el.get(qn("w:ilvl")) == str(ilvl):
            lvl = el
            break
    if lvl is not None:
        num_fmt_elem = lvl.find(qn("w:numFmt"))
        if num_fmt_elem is not None:
            return num_fmt_elem.get(qn("w:val"), "decimal")
    return "decimal"


def _to_roman(num: int) -> str:
    """将整数转换为罗马数字"""
    if num <= 0:
        return "0"
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syms = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman = ""
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman += syms[i]
            num -= val[i]
        i += 1
    return roman


def _to_chinese_num(num: int) -> str:
    """将整数转换为中文数字（一到一百）"""
    if num <= 0:
        return str(num)
    digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    if num < 10:
        return digits[num]
    if num < 100:
        tens = num // 10
        ones = num % 10
        if tens == 1:
            result = "十"
        else:
            result = digits[tens] + "十"
        if ones > 0:
            result += digits[ones]
        return result
    if num == 100:
        return "一百"
    return str(num)


def _from_roman(roman: str) -> int:
    """罗马数字转整数，如 'I'→1, 'IV'→4, 'X'→10"""
    roman = roman.strip().lower()
    if not roman:
        raise ValueError("Empty roman numeral")
    roman_map = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
    result = 0
    prev_val = 0
    for char in reversed(roman):
        val = roman_map.get(char)
        if val is None:
            raise ValueError(f"Invalid roman numeral: '{roman}'")
        if val < prev_val:
            result -= val
        else:
            result += val
        prev_val = val
    return result


def _from_chinese_num(chinese: str) -> int:
    """中文数字转整数，如 '一'→1, '十二'→12, '一百'→100, '一百二十三'→123"""
    chinese = chinese.strip()
    if not chinese:
        raise ValueError("Empty chinese numeral")

    digit_map = {
        "零": 0,
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "壹": 1,
        "贰": 2,
        "叁": 3,
        "肆": 4,
        "伍": 5,
        "陆": 6,
        "柒": 7,
        "捌": 8,
        "玖": 9,
    }
    unit_map = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000}

    # 纯数字字符
    if all(c in digit_map for c in chinese):
        result = 0
        for c in chinese:
            result = result * 10 + digit_map[c]
        return result

    result = 0
    section = 0
    for char in chinese:
        if char in digit_map:
            section = digit_map[char]
        elif char in unit_map:
            unit = unit_map[char]
            if section == 0:
                section = 1
            result += section * unit
            section = 0
        else:
            raise ValueError(f"Invalid chinese numeral character: '{char}'")
    result += section
    return result


def _make_caption_result(
    label, ch_text, ch_num, sep, num_text, num_num, name, is_continued
):
    """构建题注解析结果字典。"""
    return {
        "label": label,
        "chapter_text": ch_text,
        "chapter_num": ch_num,
        "separator": sep,
        "number_text": num_text,
        "number_num": num_num,
        "name": name.strip(),
        "is_continued": is_continued,
    }


def _try_parse_num(text: str, parser):
    """安全解析数字，失败返回 None。"""
    try:
        return parser(text)
    except ValueError:
        return None


def parse_caption_text(text: str) -> dict | None:
    """解析题注文本为结构化组件。

    支持格式：[续][标签][章节号][分隔符][题注编号] [题注名称]
    """
    text = text.strip()
    if not text:
        return None

    is_continued = text.startswith("续")
    if is_continued:
        text = text[1:].strip()
    if not text:
        return None

    SEP = r"[.\-:—–]"
    LABEL = r"([图表])\s*"
    CN = r"[一二三四五六七八九十百千零壹贰叁肆伍陆柒捌玖拾佰仟]+"
    ROMAN = r"[IVXLCDMivxlcdm]+"

    # 按优先级依次尝试各模式: (pattern, ch_parser, num_parser)
    patterns = [
        (rf"^{LABEL}(\d+)({SEP})(\d+)[\s　]+(.+)$", int, int),
        (rf"^{LABEL}({CN})({SEP})(\d+)[\s　]+(.+)$", _from_chinese_num, int),
        (rf"^{LABEL}({ROMAN})({SEP})(\d+)[\s　]+(.+)$", _from_roman, int),
        (
            rf"^{LABEL}(\d+)({SEP})({CN}|{ROMAN})[\s　]+(.+)$",
            int,
            lambda n: _from_chinese_num(n)
            if any(c in _digit_map for c in n)
            else _from_roman(n),
        ),
    ]

    for pattern, ch_parser, num_parser in patterns:
        m = re.match(pattern, text)
        if m:
            label, ch_text, sep, num_text, name = m.groups()
            return _make_caption_result(
                label,
                ch_text,
                _try_parse_num(ch_text, ch_parser),
                sep,
                num_text,
                _try_parse_num(num_text, num_parser),
                name,
                is_continued,
            )

    return None


# 供 parse_caption_text 内部使用的 digit_map
_digit_map = {
    "零": 0,
    "一": 1,
    "二": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "壹": 1,
    "贰": 2,
    "叁": 3,
    "肆": 4,
    "伍": 5,
    "陆": 6,
    "柒": 7,
    "捌": 8,
    "玖": 9,
}


# ---------------------------------------------------------------------------
# CJK 字符检测工具
# ---------------------------------------------------------------------------


def is_chinese_char(ch: str) -> bool:
    """判定单个字符是否为 CJK 统一表意文字。"""
    return "一" <= ch <= "鿿"


def count_chinese_chars(text: str) -> int:
    """统计文本中 CJK 字符数量。"""
    return sum(1 for ch in text if is_chinese_char(ch))


def extract_chinese_chars(text: str) -> str:
    """提取文本中的纯 CJK 字符序列。"""
    return "".join(ch for ch in text if is_chinese_char(ch))


def has_chinese(text: str) -> bool:
    """文本中是否包含至少一个 CJK 字符。"""
    return any(is_chinese_char(ch) for ch in text)
