"""段落内 run 级文本改写。

JSON 条目里的 ``paragraph`` / ``replace`` 只是整段文本。直接「清 run 重写」会连带删掉
run 内的加粗、斜体、超链接和图片，所以这里按字符对齐把新文本分配回**原有 run 序列**：
只有真正发生变化的字符被改写，变化处的字符继承其起点所在 run 的格式。

可写 run 的判定刻意保守（见 :func:`_is_textual_run`）：含 ``w:drawing``/``w:pict`` 的
run、域代码 run（``w:fldChar``/``w:instrText``）一律排除。``Run.text`` 赋值会清空 run
的其余内容子元素，写进图片 run 就等于删图；写进域 run 则会毁掉 Word 的自动页码与
交叉引用。文本基准 ``paragraph_text()`` 与 python-docx 的 ``Paragraph.text`` 取同一片
段（正文 run + 超链接内 run），否则一次无意义的改动会被判定成整段替换。
"""

from __future__ import annotations

import difflib

from docx.oxml.ns import qn
from docx.text.run import Run

# run 里代表「可见文字」的子元素
_TEXT_TAGS = (
    qn("w:t"),
    qn("w:tab"),
    qn("w:br"),
    qn("w:cr"),
    qn("w:noBreakHyphen"),
    qn("w:sym"),
)
# 出现在 run 内就不能改写该 run 的结构（图片/域/公式/OLE）
_PROTECTED_TAGS = (
    qn("w:drawing"),
    qn("w:pict"),
    qn("w:object"),
    qn("w:fldChar"),
    qn("w:instrText"),
    qn("m:oMath"),
    qn("m:oMathPara"),
)
# 段落直接子级中会包裹文字 run 的容器；其余容器（如 mc:AlternateContent 兼容副本）
# 一律不进入可写序列，避免把同一份文字改写两次
_TEXT_CONTAINERS = (qn("w:hyperlink"), qn("w:smartTag"))


def _is_textual_run(r_el) -> bool:
    """run 是否承载可改写的纯文字。"""
    if any(r_el.find(tag) is not None for tag in _PROTECTED_TAGS):
        return False
    return any(r_el.find(tag) is not None for tag in _TEXT_TAGS)


def _textual_runs(p_el) -> list:
    """按文档序收集段落的可写 run 元素（含超链接/智能标签内的 run）。"""
    out = []
    for child in p_el:
        if child.tag == qn("w:r"):
            if _is_textual_run(child):
                out.append(child)
        elif child.tag in _TEXT_CONTAINERS:
            out.extend(r for r in child.findall(qn("w:r")) if _is_textual_run(r))
    return out


def paragraph_text(para) -> str:
    """可写 run 拼接出的段落文字（与 ``Paragraph.text`` 同片段）。"""
    return "".join(Run(el, para).text for el in _textual_runs(para._p))


def _distribute(segments: list[str], text: str) -> list[str] | None:
    """把新文本按字符对齐分配回各 run，返回与 ``segments`` 等长的新文本列表。

    没有可承载文本的 run 时返回 None，由调用方补一个 run。
    """
    count = len(segments)
    if count == 0:
        return None
    old = "".join(segments)
    if not old:
        return [text] + [""] * (count - 1)

    owner: list[int] = []  # 每个旧字符属于哪个 run
    for idx, seg in enumerate(segments):
        owner.extend([idx] * len(seg))

    buckets = [""] * count
    matcher = difflib.SequenceMatcher(None, old, text, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i1, i2):
                buckets[owner[k]] += old[k]
            continue
        if tag == "delete" or j2 <= j1:
            continue
        chunk = text[j1:j2]
        # 新增文字跟随其前一个字符所在 run；整段替换落在区段起始 run
        if tag == "insert":
            target = owner[i1 - 1] if i1 > 0 else owner[0]
        else:
            target = owner[i1] if i1 < len(owner) else owner[-1]
        buckets[target] += chunk
    return buckets


def set_paragraph_text(para, text: str) -> bool:
    """就地改写段落文字，保留原有 run 边界与格式。

    Returns:
        True 表示确实改写了某个 run。
    """
    elements = _textual_runs(para._p)
    segments = [Run(el, para).text for el in elements]
    visible = "".join(segments)
    if visible == text:
        return False
    # gj 导出的 paragraph 会去掉首尾空白；sync 以 JSON 为准时不应因此抹掉
    # 封面字段、签名行、公式编号用于定位的缩进空格——仅首尾空白差异视为无变化。
    # 例外：原段落无实际内容（纯空白）时仍允许清空。
    # 需要真实增删首尾空格时用显式 replace 字段。
    if visible.strip() and visible.strip() == text.strip():
        return False

    distributed = _distribute(segments, text)
    if distributed is None:
        if text:
            para.add_run(text)
        return bool(text)

    changed = False
    for el, new_seg in zip(elements, distributed, strict=True):
        run = Run(el, para)
        if run.text != new_seg:
            run.text = new_seg
            changed = True
    return changed
