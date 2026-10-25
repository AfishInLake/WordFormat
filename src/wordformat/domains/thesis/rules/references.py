#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : references.py

import re
from typing import Callable, List, Tuple

from wordformat.config.dotdict import BASE_FORMAT, deep_merge
from wordformat.rules.node import FormatNode
from wordformat.structure.registry import register
from wordformat.utils import has_chinese

# ===== GB/T 7714 著录格式校验（只报批注，不改文字） =====
# 文献类型标识（GB/T 7714-2015 表1）。载体标识（/OL、/CD…）不列白名单：
# 标准里枚举不全，实际错误几乎都出在主类型上。
_GBT_DOC_TYPES = {"M", "C", "G", "N", "J", "D", "R", "S", "P", "DB", "CP", "EB"}
# 类型标识形如 [J] / [EB/OL]；纯数字的 [1] 是条目序号，不能当类型标识
_TYPE_MARKER_RE = re.compile(r"\[([A-Z]{1,3}(?:/[A-Z]{1,3})?)\]")
_YEAR_RE = re.compile(r"(?:1[89]|20)\d{2}")
# 卷(期) 的期号部分：允许 "20,25(3)"，也允许只有 "(3)"
_ISSUE_RE = re.compile(r"[\(（]\s*\d{1,3}\s*[\)）]")
_PAGES_RE = re.compile(r"[:：]\s*\d+(?:\s*[-–—]\s*\d+)?")
# 出版项：出版地:出版者,出版年
_PUB_ITEM_RE = re.compile(r"[^:：]{1,30}[:：]\s*[^,，]{2,60},\s*(?:1[89]|20)\d{2}")
# 完整日期：GB/T 7714 著录日期一律用 YYYY-MM-DD，不接受中文年月日
_FULL_DATE_RE = re.compile(r"(?:1[89]|20)\d{2}\s*-\s*\d{1,2}\s*-\s*\d{1,2}")
# 报纸版次：(3)，括号内只写数字，不写“版”
_EDITION_RE = re.compile(r"[\(（]\s*\d{1,3}\s*[\)）]")
# 标准编号：GB/T 7714-2015、ISO 9001、IEEE 802.11 等
_STANDARD_NO_RE = re.compile(
    r"\b(?:GB|DB\d*|YY|JB|QB|SJ|HG|CJ|JG|TB|SY|NB|MH|WS|"
    r"ISO|IEC|IEEE|ANSI|ASTM|JIS|DIN|BS|EN)"
    r"(?:\s*/\s*[A-Z]{1,2})?\s*\d[\d.\-—]*",
    re.IGNORECASE,
)
# 专利号：CN101234567A、ZL 2010 1 0123456.7 或“专利号: …”
_PATENT_NO_RE = re.compile(
    r"(?:CN|ZL|US|EP|JP|WO|KR)\s*\d[\dA-Z.]*"
    r"|专利号\s*[:：]?\s*[A-Z]{0,3}\s*\d[\dA-Z.]*",
    re.IGNORECASE,
)
# 电子资源的（更新日期）/[引用日期]
_CITATION_DATE_RE = re.compile(
    r"[\[(（]\s*(?:1[89]|20)\d{2}-\d{1,2}(?:-\d{1,2})?\s*[\])）]"
)
_ACCESS_PATH_RE = re.compile(r"(?:https?://|www\.|DOI[:：])", re.IGNORECASE)
_ELLIPSIS_RE = re.compile(r"\.{3,}|…+")
_FULL_WIDTH_PUNCT = "：，；（）．"

# (问题类型, 现状, 规范)
_Issue = Tuple[str, str, str]
# 校验函数统一签名：(完整条目, [类型标识] 之后的著录部分, 类型标识)，返回问题列表
_Checker = Callable[[str, str, str], List[_Issue]]
# (配置键, 缺省开关, 校验函数)
_FieldCheck = Tuple[str, bool, _Checker]

_MARKER_MISSING: _Issue = (
    "类型标识缺失",
    "无 [J]/[M]/[D] 等标识",
    "文献类型标识 [J]/[M]/[C]/[D]/[EB/OL]…",
)


def _check_journal_tail(_text: str, tail: str, code: str) -> List[_Issue]:
    """[J] 期刊：出版年必查；页码必查（联机版可无页码）。

    卷(期) 与页码至少有一项，但只有卷号没有期号（如 “2020, 43: 100-110”）
    是合法写法，因此期号本身不单独报错。
    """
    if not _YEAR_RE.search(tail):
        return [("出版年缺失", "无 4 位出版年", "出版年,卷(期):页码")]
    if not _ISSUE_RE.search(tail) and not _PAGES_RE.search(tail):
        return [("卷期页缺失", "既无卷(期)也无页码", "出版年,卷(期):页码")]
    if "/OL" not in code and not _PAGES_RE.search(tail):
        return [("页码缺失", "无起始页码", "出版年,卷(期):页码")]
    return []


def _check_publisher_tail(_text: str, tail: str, _code: str) -> List[_Issue]:
    """专著/论文集/学位论文/汇编/报告：需有“出版地:出版者,出版年”。"""
    if _PUB_ITEM_RE.search(tail):
        return []
    return [("出版项缺失", "无“出版地:出版者,出版年”", "出版地:出版者,出版年")]


def _check_carrier_tail(text: str, tail: str, code: str) -> List[_Issue]:
    """出版项校验：联机（/OL）文献可省略出版项，离线载体仍需著录。"""
    if "/OL" in code:
        return []
    return _check_publisher_tail(text, tail, code)


def _check_newspaper_tail(_text: str, tail: str, _code: str) -> List[_Issue]:
    """[N] 报纸：需有“报纸名, 出版日期(版次)”。"""
    issues: List[_Issue] = []
    date = _FULL_DATE_RE.search(tail)
    if date is None:
        issues.append(
            ("出版日期缺失", "无“YYYY-MM-DD”出版日期", "报纸名, 出版日期(版次)")
        )
    elif not tail[: date.start()].strip(" ,，.。"):
        issues.append(("报纸名缺失", "出版日期前无报纸名", "报纸名, 出版日期(版次)"))
    if not _EDITION_RE.search(tail):
        issues.append(("版次缺失", "无“(版次)”", "出版日期(版次)"))
    return issues


def _check_standard_tail(text: str, tail: str, _code: str) -> List[_Issue]:
    """[S] 标准：标准编号必查，并需有出版项。"""
    issues: List[_Issue] = []
    if not _STANDARD_NO_RE.search(text):
        issues.append(
            ("标准号缺失", "无 GB/T、ISO 等标准编号", "标准编号, 标准名称[S]")
        )
    if not _PUB_ITEM_RE.search(tail):
        issues.append(
            ("出版项缺失", "无“出版地:出版者,出版年”", "出版地:出版者,出版年")
        )
    return issues


def _check_patent_tail(text: str, tail: str, _code: str) -> List[_Issue]:
    """[P] 专利：专利号必查，并需有公告/公开日期。"""
    issues: List[_Issue] = []
    if not _PATENT_NO_RE.search(text):
        issues.append(("专利号缺失", "无 CN/ZL 等专利号", "专利题名: 专利号[P]"))
    if not _FULL_DATE_RE.search(tail):
        issues.append(("公告日期缺失", "无“YYYY-MM-DD”日期", "公告日期或公开日期"))
    return issues


# 按主类型分派的著录项校验，覆盖 GB/T 7714-2015 表1 的全部文献类型
_PUBLISHER_CHECKS: Tuple[_FieldCheck, ...] = (
    ("check_publisher_item", True, _check_carrier_tail),
)
_TYPE_CHECKS: dict[str, Tuple[_FieldCheck, ...]] = {
    "J": (("check_journal_item", True, _check_journal_tail),),
    "N": (("check_newspaper_item", True, _check_newspaper_tail),),
    "S": (("check_standard_item", True, _check_standard_tail),),
    "P": (("check_patent_item", True, _check_patent_tail),),
    "M": _PUBLISHER_CHECKS,
    "C": _PUBLISHER_CHECKS,
    "G": _PUBLISHER_CHECKS,
    "R": _PUBLISHER_CHECKS,
    "D": _PUBLISHER_CHECKS,
    "DB": (("check_publisher_item", True, _check_carrier_tail),),
    "CP": (("check_publisher_item", True, _check_carrier_tail),),
    "EB": (("check_publisher_item", True, _check_carrier_tail),),
}


def _check_electronic_item(text: str, _tail: str, code: str) -> List[_Issue]:
    """联机网络文献（/OL）：需引用日期与获取路径。"""
    if "/OL" not in code:
        return []
    issues: List[_Issue] = []
    if not _CITATION_DATE_RE.search(text):
        issues.append(
            ("引用日期缺失", "无（更新日期）或[引用日期]", "(更新或修改日期)[引用日期]")
        )
    if not _ACCESS_PATH_RE.search(text):
        issues.append(("获取路径缺失", "无网址或 DOI", "获取和访问路径或 DOI"))
    return issues


def _check_author_ellipsis(text: str, _tail: str, _code: str) -> List[_Issue]:
    """著者省略：超过 3 人应列前 3 人后加“，等”或“, et al”，不用省略号。"""
    issues: List[_Issue] = []
    if _ELLIPSIS_RE.search(text):
        issues.append(
            (
                "著者省略错误",
                "出现省略号",
                "著者超 3 人时列前 3 人后加“，等”或“, et al”",
            )
        )
    if "et al" in text.lower() and has_chinese(text):
        issues.append(
            (
                "著者省略错误",
                "中文著者后接 et al",
                "中文条目用“，等”，英文条目用“, et al”",
            )
        )
    return issues


def _check_terminator(text: str, _tail: str, _code: str) -> List[_Issue]:
    """条目末尾句点：各校分歧大，默认关闭。"""
    if text.rstrip().endswith("."):
        return []
    return [("结尾符缺失", "末尾无半角句点", "著录条目以“.”结束")]


def _check_punctuation_width(text: str, _tail: str, _code: str) -> List[_Issue]:
    """著录符号（. , : ( )）应用半角。

    题名内部的中文标点本身是合法全角，无法逐项区分，故默认关闭、由配置开启。
    """
    found = sorted({ch for ch in text if ch in _FULL_WIDTH_PUNCT})
    if not found:
        return []
    return [
        (
            "标点问题",
            f"混用全角符号“{'、'.join(found)}”",
            "GB/T 7714 著录符号用半角 . , : ( )",
        )
    ]


# (配置键, 缺省, 校验函数)
_EXTRA_CHECKS: Tuple[_FieldCheck, ...] = (
    ("check_electronic_item", True, _check_electronic_item),
    ("check_author_ellipsis", True, _check_author_ellipsis),
    ("check_terminator", False, _check_terminator),
    ("check_punctuation_width", False, _check_punctuation_width),
)


def _marker_issue(code: str) -> _Issue:
    return (
        "类型标识错误",
        f"[{code}]",
        "文献类型标识只能取 " + "、".join(sorted(_GBT_DOC_TYPES)),
    )


def check_gbt7714_entry(text: str, rule_cfg) -> List[_Issue]:
    """按 GB/T 7714 校验一条著录，返回 [(问题类型, 现状, 规范), ...]。

    rule_cfg 为节点配置 rules.gbt_7714 子项；各子项可单独写 false 关闭，
    缺省值见 ReferenceEntry.DEFAULTS（结尾句点与全角标点默认关闭）。
    """
    match = _TYPE_MARKER_RE.search(text)
    if match is None:
        return [_MARKER_MISSING] if rule_cfg.get("require_type_marker", True) else []
    code = match.group(1)
    main_type = code.split("/")[0]
    if rule_cfg.get("require_type_marker", True) and main_type not in _GBT_DOC_TYPES:
        return [_marker_issue(code)]

    tail = text[match.end() :]
    issues = _type_specific_issues(main_type, text, tail, code, rule_cfg)
    for cfg_key, default, checker in _EXTRA_CHECKS:
        if rule_cfg.get(cfg_key, default):
            issues += checker(text, tail, code)
    return issues


def _type_specific_issues(
    main_type: str, text: str, tail: str, code: str, rule_cfg
) -> List[_Issue]:
    """按主类型逐项校验著录字段，类型与字段的对应关系见 _TYPE_CHECKS。"""
    # （更新日期）/[引用日期] 不是出版日期，先剔除，避免把引用日期当成出版年
    tail = _CITATION_DATE_RE.sub(" ", tail)
    issues: List[_Issue] = []
    for cfg_key, default, checker in _TYPE_CHECKS.get(main_type, ()):
        if rule_cfg.get(cfg_key, default):
            issues += checker(text, tail, code)
    return issues


@register("references_title", level=1)
class References(FormatNode):
    """参考文献节点"""

    NODE_TYPE = "references.title"
    NODE_LABEL = "参考文献标题"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": ["居中对齐", "两端对齐", "左对齐"],
                "space_before": ["0行", "0.5行", "1行"],
                "space_after": ["0行", "0.5行", "1行"],
                "first_line_indent": "0字符",
            },
            "font": {
                "chinese_font_name": "黑体",
                # 单行标题字号各校不一：三号/小二/小三/五号均常见
                "font_size": ["三号", "小二", "小三", "五号"],
                "bold": False,
            },
        },
    )

    def _para_checks_to_skip(self) -> set[str]:
        # 单行标题的行距无视觉意义
        return {"first_line_indent", "line_spacing_rule", "line_spacing"}


@register("references_content")
class ReferenceEntry(FormatNode):
    """参考文献条目节点"""

    NODE_TYPE = "references.entry"
    NODE_LABEL = "参考文献条目"
    DEFAULTS = deep_merge(
        BASE_FORMAT,
        {
            "paragraph": {
                "alignment": "两端对齐",
                # 各校参考文献排版分歧较大：悬挂缩进 -2 / 首行缩进 +2 / 无缩进
                # 都是常见且合规的写法，列表表示任一命中即可。
                "first_line_indent": ["-2字符", "0字符", "2字符"],
            },
            "font": {
                "chinese_font_name": "宋体",
                "font_size": ["五号", "小四"],
            },
            "rules": {
                "gbt_7714": {
                    "enabled": True,
                    "require_type_marker": True,
                    "check_journal_item": True,
                    "check_publisher_item": True,
                    "check_newspaper_item": True,
                    "check_standard_item": True,
                    "check_patent_item": True,
                    "check_electronic_item": True,
                    "check_author_ellipsis": True,
                    # 以下两项各校做法分歧大，默认关闭，按需开启
                    "check_terminator": False,
                    "check_punctuation_width": False,
                }
            },
        },
    )
    RULES = {"gbt_7714": "_check_gbt_7714"}

    def _check_gbt_7714(self, doc, rule_cfg, p: bool = False):
        """GB/T 7714 著录项校验：只报批注，不改文字（check/apply 都报）。"""
        if self.paragraph is None or not self.paragraph.runs:
            return
        text = self.paragraph.text.strip()
        if not text:
            return
        from wordformat.style.comments import format_comment

        for label, actual, standard in check_gbt7714_entry(text, rule_cfg):
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=format_comment(self.NODE_LABEL, label, f"当前{actual}", standard),
            )
