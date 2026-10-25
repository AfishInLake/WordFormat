#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

import re

from docx import Document
from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.settings import BATCH_SIZE
from wordformat.utils import (
    get_paragraph_numbering_text,
    para_contains_image,
    para_contains_math,
    parse_caption_text,
)

# ===== 分级置信度阈值（C2）=====
# 低于阈值不再硬砍为 body_text，而是保留原标签并标记 needs_review 供人工复核。
# 2026-09 校准：以 6 篇留出集（1798 段，教师强制 PREV）上 int8 模型的每类正确段
# p10 为基准分档——保证每类误标率 ≤~10%，小摘要类阈值降到正确分布下沿。
# 注意与 WordFormatUI/src/composables/useTagHelpers.js 的 CONF_THRESHOLDS 保持一致。
CONF_THRESHOLDS = {
    # 0.6 档：主要段落类（正确段 p10 ≥ 0.64）
    "body_text": 0.6,
    "heading_mulu": 0.6,
    "caption_figure": 0.6,
    "heading_level_1": 0.6,
    "heading_level_2": 0.6,
    "heading_level_3": 0.6,
    "other": 0.6,
    # 0.55 档：内容/标题类（正确段 p10 0.55~0.61）
    "references_content": 0.55,
    "acknowledgements_content": 0.55,
    "keywords_english": 0.55,
    "abstract_chinese_content": 0.55,
    "acknowledgements_title": 0.55,
    # 0.5 档：短标题/题注类（正确段 p10 0.50~0.58）
    "document_title": 0.5,
    "caption_table": 0.5,
    "abstract_chinese_title_content": 0.5,
    # 0.4 档：小样本标题/关键词类
    "keywords_chinese": 0.4,
    "abstract_chinese_title": 0.4,
    "references_title": 0.4,
    # 0.35 档：天然低置信但验证准确的小类（英文摘要/附录标题）
    "abstract_english_title": 0.35,
    "abstract_english_title_content": 0.35,
    "abstract_english_content": 0.35,
    "heading_fulu": 0.35,
    "default": 0.5,
}

# ===== 章节状态机（C5）：章节起点标签 → 章节状态 =====
SECTION_START = {
    "abstract_chinese_title": "abstract_cn",
    "abstract_english_title": "abstract_en",
    "references_title": "references",
    "acknowledgements_title": "acknowledgements",
}

# 章节状态 → 内容标签（使用后端真实注册的类别名，避免节点被丢弃）
SECTION_CONTENT = {
    "abstract_cn": "abstract_chinese_content",
    "abstract_en": "abstract_english_content",
    "references": "references_content",
    "acknowledgements": "acknowledgements_content",
}

# 正文标题：遇到则退出特殊章节，避免状态泄漏到下一章节
BODY_HEADINGS = {
    "heading_level_1",
    "heading_level_2",
    "heading_level_3",
    "heading_fulu",
}

# 参考文献条目编号：[1] / 1.
REF_ITEM_PATTERN = re.compile(r"^\s*(\[\d+\]|\d+\.)\s")

# 页脚 / AI 生成声明
FOOTER_PATTERN = re.compile(r"由\s*AI\s*生成", re.IGNORECASE)

# 状态机置信度门控：只有足够可信的章节标题才触发“内容标签”传播。
# 避免源码附录里被误判的低置信标题（如 "class X:"→abstract_english_title 0.45）
# 把连续几十行代码放大成 *_content。真实的摘要/参考文献/致谢标题通常 >0.9。
STATE_MIN_CONF = 0.6

# 附录（heading_fulu）区内需要中性化为 body_text 的前置/后置类标签：
# 附录通常是源码或截图，不含摘要/关键词/致谢等结构，
# 模型对代码行的误判会污染识别准确率与文档树（产生伪标题节点）。
APPENDIX_NEUTRALIZE = {
    "abstract_chinese_title",
    "abstract_english_title",
    "abstract_chinese_title_content",
    "abstract_english_title_content",
    "abstract_chinese_content",
    "abstract_english_content",
    "keywords_chinese",
    "keywords_english",
    "acknowledgements_title",
    "acknowledgements_content",
}

# 目录标题（模型词表无 heading_mulu，用规则补齐）
TOC_PATTERN = re.compile(r"^目\s*录$")

# ===== 目录条目识别（规则）=====
# 目录条目既不是正文也不是标题：样式为 toc 1/2/3，或文本以“制表符 + 页码”结尾
# （未套用 toc 样式、被静态化的目录）。若不排除，目录项会被当成标题参与层级
# 计数与章节状态传播，导致题注章节号和正文格式出现大量误报。
TOC_ENTRY_STYLE_PATTERN = re.compile(r"^toc\s*\d+$", re.IGNORECASE)
TOC_ENTRY_TAIL_PATTERN = re.compile(r"\t\s*(?:\d+|[ivxlcdm]+)\s*$", re.IGNORECASE)
TOC_TITLE_PATTERN = re.compile(r"^(?:目\s*录|图\s*目\s*录|表\s*目\s*录)$")

# 附录标题：模型对小标题/代码区标题置信度低，常误判为一级标题。
# 命中后整段区域按附录处理（子树不参与格式化），避免源码被逐行当正文检查。
APPENDIX_HEADING_PATTERN = re.compile(
    r"^(?:附\s*录(?:\s*[A-Za-z0-9一二三四五六七八九十]+)?"
    r"(?:\s*[：:].*)?|程\s*序\s*代\s*码|源\s*代\s*码|源\s*程\s*序|附\s*件)\s*$"
)


def _is_toc_entry_text(text: str) -> bool:
    return bool(TOC_ENTRY_TAIL_PATTERN.search(text))


# ===== 独立公式行识别（结构规则）=====
# 模型词表没有公式类，且 paragraph.text 不含 OMML 文字，公式段读起来往往是
# 空段或只剩右侧编号。这里按 XML 结构短路判定，不走 AI 推理。
# 编号形态：（2-1）/(2.1)/[3-12]，允许前后空白与制表符；必须整段只有编号，
# 含其他文字的段落（如"由式（2-1）可得"）仍交给模型分类，避免吞掉正文。
EQUATION_NUMBER_PATTERN = re.compile(
    r"^[\s\t]*[（(\[]\s*\d{1,2}\s*[-–—.．]\s*\d{1,3}\s*[)）\]][\s\t]*$"
)


def _is_display_formula(para, text: str) -> bool:
    """段落是否为独立公式行。

    可见文字（para.text 不含 OMML 的 m:t）只剩公式编号时判定为公式行；
    "由式（2-1）可得"这类带正文的段落仍交给模型分类。
    """
    return para_contains_math(para) and bool(EQUATION_NUMBER_PATTERN.match(text))


def _structural_item(para, text: str) -> dict:
    """按 XML 结构短路判定的段落条目（图片 / 公式 / 空段）。"""
    if para_contains_image(para):
        category, note = "figure_image", "图片段落"
    elif para_contains_math(para):
        category, note = "math_block", "公式段落"
    else:
        category, note = "body_text", "空段落"
    return {
        "category": category,
        "score": 1.0,
        "comment": note,
        "paragraph": text,
        "needs_review": False,
    }


# ===== 标题编号层级修正（正则校准）=====
# 模型对标题层级泛化不足，常把 X.Y 二级标题误判为三级（或反之）。
# 此处以编号格式为强约束：仅对模型已预测为 heading_level_* 的段落推断
# 编号层级，与预测不符时以编号为准；无编号/无法确定时保持原判，避免误伤正文。
HEADING_LEVEL_CATS = {"heading_level_1", "heading_level_2", "heading_level_3"}
# 第X章 / 第1章（中文数字或阿拉伯数字）
_HEADING_CHAPTER_RE = re.compile(r"^第\s*[一二三四五六七八九十百千零0-9]+\s*章")
# 数字编号：1 / 1.1 / 1.1.1（编号后必须跟顿号/点/空格+标题文字，防止误吞正文数字）
_HEADING_NUM_RE = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,2})(?![.\d])(?:[、.．]?\s*\S)")
# 中文数字编号：一、/ 二、…（须跟顿号/点/空格，防止"一是…"误判）
_HEADING_CN_NUM_RE = re.compile(r"^[一二三四五六七八九十]{1,3}(?=[、.．]|\s+\S)")

# ===== 摘要页标题回补（位置规则）=====
# 合并式摘要正文（*_title_content）→ 其上方独立标题行应升格为的标题类别
ABSTRACT_CONTENT_TO_TITLE = {
    "abstract_chinese_title_content": "abstract_chinese_title",
    "abstract_english_title_content": "abstract_english_title",
}
# 可升格为摘要标题的前置标签（封面保护后的 other、正文、被误判的关键词）
ABSTRACT_TITLE_PREV_CATS = {
    "other",
    "body_text",
    "keywords_english",
    "keywords_chinese",
}
# 封面/声明字段特征：命中则不升格（避免把姓名/学院/签名/日期行误判为标题）
COVER_MARKER_PATTERN = re.compile(
    r"(签\s*名|日\s*期|学\s*号|学\s*院|专\s*业|指导教师|职\s*称|声\s*明|导\s*师|班\s*级|姓\s*名)"
)

# ===== 摘要内容位置门控 =====
# 摘要开场句与绪论首段在词面上同构（“随着…发展”“目前…”），仅凭文本难以区分；
# 摘要内容段只允许紧跟摘要标题/上一段摘要内容出现，否则回退为正文。
ABSTRACT_CONTENT_GATE_CATS = {"abstract_chinese_content", "abstract_english_content"}
ABSTRACT_CONTENT_GATE_ALLOWED_PREV = {
    "abstract_chinese_title",
    "abstract_chinese_title_content",
    "abstract_chinese_content",
    "abstract_english_title",
    "abstract_english_title_content",
    "abstract_english_content",
}
# 摘要页标题最大长度（中文标题短，英文标题放宽；超长多为正文句子）
ABSTRACT_TITLE_MAXLEN = 120


class DocxBase:
    def __init__(self, docx_file, configpath):
        self.re_dict = {}
        self.docx_file = docx_file
        self.document = Document(docx_file)
        """
        以下注释掉的代码用于未来加载配置文件
        """
        # init_config(configpath)
        # try:
        #     self.config_model = get_config()  # 首次调用：触发load()
        #     self.config = self.config_model.model_dump()
        #     logger.info("配置文件验证通过")
        # except Exception as e:
        #     logger.error(f"配置加载失败: {str(e)}")
        #     raise

    def parse(self) -> list[dict]:
        # 收集所有段落（含空段）：空段按结构判定（图片/公式/空段），
        # 独立公式行短路，其余走 AI 推理
        all_paras = list(self.document.paragraphs)
        result: list = [None] * len(all_paras)
        text_indices = []
        texts_for_ai = []
        for idx, para in enumerate(all_paras):
            text = para.text.strip()
            if not text:
                result[idx] = _structural_item(para, "")
                continue
            if TOC_ENTRY_STYLE_PATTERN.match(_paragraph_style_name(para)):
                result[idx] = _toc_entry_item(text)
                continue
            if _is_display_formula(para, text):
                # 编号文字要保留，否则 jd 重建会丢掉 (2-1)
                result[idx] = _structural_item(para, text)
                continue
            numbering_text = get_paragraph_numbering_text(para)
            full_text = f"{numbering_text} {text}" if numbering_text else para.text
            texts_for_ai.append(full_text)
            text_indices.append(idx)

        # 对非空段进行批量 AI 推理（跨批次维护 PREV 上下文状态链）
        prev_label = None
        for i in range(0, len(texts_for_ai), BATCH_SIZE):
            batch_texts = texts_for_ai[i : i + BATCH_SIZE]
            batch_indices = text_indices[i : i + BATCH_SIZE]

            try:
                batch_results, prev_label = onnx_batch_infer(batch_texts, prev_label)
            except Exception as e:
                logger.error(f"批量推理失败，降级到单条处理: {e}")
                batch_results = []
                for t in batch_texts:
                    r = onnx_single_infer(t, prev_label)
                    batch_results.append(r)
                    if r["label"]:
                        prev_label = r["label"]

            for idx, text, pred in zip(
                batch_indices, batch_texts, batch_results, strict=False
            ):
                response = {
                    "category": pred["label"],
                    "score": pred["score"],
                    "comment": f"置信度：{pred['score']:.4f}",
                    "paragraph": text,
                }
                _apply_threshold(response)
                result[idx] = response

        assert all(r is not None for r in result), "存在未处理的段落"

        # 后处理顺序很重要：
        # 1) 文档标题 → 2) 英文摘要标题兜底 → 3) 摘要合并段/封面保护
        # → 4) 摘要页标题回补 → 5) 页脚识别 → 6) 目录标题 → 7) 附录区中性化
        # → 8) 章节状态机（置信度门控）→ 9) 序列修正 → 10) 参考文献区位置门控
        # → 11) 标题编号正则校准
        _fix_document_title(result)
        _fix_abstract_en_title(result)
        _fix_known_categories(result)
        _fix_abstract_titles(result)
        _gate_abstract_content(result)
        _apply_footer(result)
        _fix_toc(result)
        _fix_toc_entries(result, all_paras)
        _fix_appendix_heading(result)
        _neutralize_appendix(result)
        _fix_title_categories(result)
        _apply_section_state(result)
        _fix_sequence(result)
        _fix_keyword_categories(result)
        _fix_references_content(result)
        _fix_heading_levels(result)
        return self._interleave_tables(result)

    def _interleave_tables(self, result: list[dict]) -> list[dict]:
        """把文档中的表格按文档序导出为 table_object 条目。

        ``document.paragraphs`` 只含 body 直接子级段落，表格单元格不在其中，
        原表格在 gj JSON 中完全不可见。此处按 body XML 顺序遍历 w:p / w:tbl，
        在每个表格前记录已出现的段落数，把表格条目插入对应位置。

        表格条目为**独立项**：不占用段落槽位，带 ``exists=True`` 供 zip 侧
        跳过段落绑定（建树阶段不绑定原表格）；``table_rows`` 为全部单元格
        文本（行优先，合并单元格按 python-docx 语义拍平）。
        """
        from docx.oxml.ns import qn

        body = self.document.element.body
        tables = []
        p_count = 0
        for el in body.iterchildren():
            if el.tag == qn("w:p"):
                p_count += 1
            elif el.tag == qn("w:tbl"):
                tables.append((p_count, el))
        if not tables:
            return result

        out = list(result)
        for k, (pos, tbl_el) in enumerate(tables):
            out.insert(
                min(pos + k, len(out)),
                {
                    "category": "table_object",
                    "paragraph": "",
                    "score": 1.0,
                    "comment": "原文档表格（内容只读，不参与格式检查）",
                    "table_rows": self._extract_table_rows(tbl_el),
                    "exists": True,
                    "needs_review": False,
                },
            )
        return out

    def _extract_table_rows(self, tbl_el) -> list[list[str]]:
        """提取表格全部单元格文本（行优先）。"""
        from docx.table import Table

        table = Table(tbl_el, self.document._body)
        return [[cell.text for cell in row.cells] for row in table.rows]


def _apply_threshold(item: dict) -> None:
    """分级置信度（C2）：低于阈值保留原标签并加 needs_review 标记，不再硬砍 body_text。"""
    label = item["category"]
    score = item["score"]
    threshold = CONF_THRESHOLDS.get(label, CONF_THRESHOLDS["default"])
    if score < threshold:
        item["needs_review"] = True
        item["comment"] = (
            f"原预测 {label}，置信度 {score:.4f} < {threshold}，建议人工复核"
        )
    else:
        item["needs_review"] = False


def _fix_document_title(result: list[dict]) -> None:
    """文档标题检测（C3）：第一段 + 短 + 无句末标点 + 后续紧跟摘要 → document_title。"""
    if not result:
        return
    first = result[0]
    if first["category"] != "body_text":
        return
    text = (first.get("paragraph") or "").strip()
    if not text or len(text) > 40:
        return
    if text.endswith(("。", "！", "？", "；")):
        return
    for j in range(1, min(8, len(result))):
        t = (result[j].get("paragraph") or "").strip()
        if re.match(r"^(摘\s*要|Abstract|ABSTRACT)", t):
            first["category"] = "document_title"
            first["comment"] = "文档标题（位置规则）"
            first["score"] = 1.0
            first["needs_review"] = False
            return


def _fix_abstract_en_title(result: list[dict]) -> None:
    """英文摘要标题兜底（C4）：独立成段的 Abstract 直接覆盖。"""
    for item in result:
        t = (item.get("paragraph") or "").strip()
        if t in ("Abstract", "ABSTRACT", "abstract", "ABSTRACT:"):
            if item["category"] != "abstract_english_title":
                item["category"] = "abstract_english_title"
                item["comment"] = "英文摘要标题（规则覆盖）"
                item["score"] = 1.0
                item["needs_review"] = False


def _apply_footer(result: list[dict]) -> None:
    """页脚识别：AI 生成声明等段落 → footer（不参与格式化）。"""
    for item in result:
        if item["category"] != "body_text":
            continue
        text = (item.get("paragraph") or "").strip()
        if text and FOOTER_PATTERN.search(text):
            item["category"] = "footer"
            item["comment"] = "页脚/AI 生成声明（规则识别）"
            item["score"] = 1.0
            item["needs_review"] = False


def _fix_toc(result: list[dict]) -> None:
    """目录标题（规则补齐）：模型词表无 heading_mulu，独立成段的“目录”直接覆盖。"""
    for item in result:
        t = (item.get("paragraph") or "").strip()
        if TOC_PATTERN.match(t) and item["category"] != "heading_mulu":
            item["category"] = "heading_mulu"
            item["comment"] = "目录标题（规则覆盖）"
            item["score"] = 1.0
            item["needs_review"] = False


def _paragraph_style_name(para) -> str:
    style = getattr(para, "style", None)
    return (getattr(style, "name", "") or "").strip()


def _toc_entry_item(text: str) -> dict:
    return {
        "category": "toc_entry",
        "score": 1.0,
        "comment": "目录条目（规则识别，跳过格式化）",
        "paragraph": text,
        "needs_review": False,
    }


def _fix_toc_entries(result: list[dict], paragraphs) -> None:
    """把目录条目标记为 ``toc_entry``，与正文/标题彻底隔离。

    两类来源：
    1. 段落样式为 ``toc 1`` / ``toc 2`` / ``toc 3``（Word 自动目录）；
    2. 静态化目录：``目录`` 标题之后、以“制表符 + 页码”结尾的连续段落。
       目录块遇到第一个非条目段落即结束，避免误伤正文。
    """
    in_toc = False
    for item, para in zip(result, paragraphs, strict=False):
        text = (para.text or "").strip()
        if TOC_ENTRY_STYLE_PATTERN.match(_paragraph_style_name(para)):
            item.update(_toc_entry_item(text))
            in_toc = True
            continue
        if not text:
            continue
        if TOC_TITLE_PATTERN.match(text):
            in_toc = True
            continue
        if in_toc and _is_toc_entry_text(text):
            item.update(_toc_entry_item(text))
            continue
        in_toc = False


def _heading_number_level(text: str) -> int | None:
    """从段落编号格式推断标题层级（1/2/3）；无编号或无法确定返回 None。"""
    t = (text or "").strip()
    if not t:
        return None
    if _HEADING_CHAPTER_RE.match(t) or _HEADING_CN_NUM_RE.match(t):
        return 1
    m = _HEADING_NUM_RE.match(t)
    if not m:
        return None
    return len(m.group(1).split("."))


def _fix_references_content(result: list[dict]) -> None:
    """参考文献内容位置门控：references_content 只能出现在首个参考文献标题之后。

    正文综述段常因包含 [n] 引用标记被模型误判为参考文献条目（低置信度），
    位置是强约束——参考文献条目不可能出现在参考文献标题之前。
    """
    first_title = None
    for i, item in enumerate(result):
        if item["category"] == "references_title":
            first_title = i
            break
    if first_title is None:
        return
    for i in range(first_title):
        item = result[i]
        if item["category"] != "references_content":
            continue
        item["category"] = "body_text"
        item["comment"] = (
            "参考文献区位置门控（原：references_content，位于参考文献标题前）"
        )
        item["score"] = 1.0
        item["needs_review"] = False


def _fix_heading_levels(result: list[dict]) -> None:
    """标题编号正则校准：模型常混淆标题层级（如把 X.Y 二级标题认成三级），
    以段落编号格式为准修正预测层级。仅处理已被预测为 heading_level_* 的段落；
    无编号或编号层级无法确定时保持原判，避免误伤正文段落。"""
    for item in result:
        cat = item["category"]
        if cat not in HEADING_LEVEL_CATS:
            continue
        level = _heading_number_level(item.get("paragraph") or "")
        if level is None:
            continue
        expected = f"heading_level_{level}"
        if expected == cat:
            continue
        item["category"] = expected
        item["comment"] = f"标题编号正则修正：{cat} → {expected}"
        item["score"] = 1.0
        item["needs_review"] = False


def _fix_appendix_heading(result: list[dict]) -> None:
    """附录标题规则补齐：``附录``/``程序代码``/``源代码`` 等小标题直接标为
    ``heading_fulu``。

    模型对代码区标题置信度低（常见 0.6~0.65），会误判为一级标题，使后续
    源码段落挂进正文并被逐行检查。只对短标题且当前为标题/正文类的段落生效。
    """
    for item in result:
        text = (item.get("paragraph") or "").strip()
        if not text or len(text) > 20:
            continue
        if item["category"] not in HEADING_LEVEL_CATS | {"body_text", "other"}:
            continue
        if not APPENDIX_HEADING_PATTERN.match(text):
            continue
        item["category"] = "heading_fulu"
        item["comment"] = "附录标题（规则识别）"
        item["score"] = 1.0
        item["needs_review"] = False


def _neutralize_appendix(result: list[dict]) -> None:
    """附录区中性化：可信 heading_fulu（附录）之后，把摘要/关键词/致谢等前置类
    标签强制回 body_text，并把低置信度的伪标题（源码里的 ``end`` 等）降为正文。

    附录通常放源码或截图，不含这些结构；模型会把代码行（import/class/:param 等）
    误判为英文摘要/关键词，若不中性化会拉低准确率并污染文档树（产生伪标题节点）。
    区域从可信 heading_fulu 起，到下一个可信 heading_level_1 止（附录一般在末尾）。
    """
    in_appendix = False
    for item in result:
        cat = item["category"]
        score = item.get("score", 0.0)
        if cat == "heading_fulu":
            if score >= STATE_MIN_CONF:
                in_appendix = True
            continue
        if cat == "heading_level_1" and score >= STATE_MIN_CONF:
            in_appendix = False
            continue
        if not in_appendix:
            continue
        if cat in APPENDIX_NEUTRALIZE or (
            cat in HEADING_LEVEL_CATS and score < STATE_MIN_CONF
        ):
            item["category"] = "body_text"
            item["comment"] = f"附录区中性化：{cat} → body_text"
            item["score"] = 1.0
            item["needs_review"] = False


def _apply_section_state(result: list[dict]) -> None:
    """章节状态机（C5）：把特殊章节内的 body_text 修正为对应内容标签。

    仅处理 body_text；遇到任何其他已识别标签（标题/关键词/题注等）立即退出
    当前章节，避免状态泄漏到下一章节。参考文献条目还需匹配编号模式。
    """
    current_section = None

    for item in result:
        label = item["category"]

        # 1. 章节起点：仅当标题足够可信才进入状态（门控，防止源码附录伪标题放大错误）
        if label in SECTION_START:
            if item.get("score", 0.0) >= STATE_MIN_CONF:
                current_section = SECTION_START[label]
            else:
                current_section = None  # 低置信标题不触发传播，并结束当前章节
            continue

        # 2. 非 body_text 的其他标签 → 退出特殊章节
        if label != "body_text":
            current_section = None
            continue

        # 3. 特殊章节内的非空 body_text → 修正为对应内容标签
        text = (item.get("paragraph") or "").strip()
        if not current_section or not text:
            continue
        new_label = SECTION_CONTENT.get(current_section)
        if not new_label:
            continue
        if current_section == "references" and not REF_ITEM_PATTERN.match(text):
            continue
        item["category"] = new_label
        item["comment"] = f"章节状态修正：{current_section} → {new_label}"
        item["score"] = 1.0
        item["needs_review"] = False


def _mark_cover_as_other(result: list[dict]) -> None:
    """封面保护：摘要之前的 body_text 标为 other（跳过格式化）。

    保留已由位置规则识别出的 document_title 等特殊标签。
    """
    abstract_start = None
    for i, item in enumerate(result):
        text = (item.get("paragraph") or "").strip()
        if re.match(r"^(摘\s*要|Abstract|ABSTRACT)", text):
            abstract_start = i
            break
    if abstract_start is None:
        return
    for i in range(abstract_start):
        if result[i]["category"] == "body_text":
            result[i]["category"] = "other"
            result[i]["comment"] = "摘要前封面/声明内容，跳过格式化（原：body_text）"
            result[i]["score"] = 1.0
            result[i]["needs_review"] = False


# 代码块残留被模型当成标题：正文里嵌入的源码行（如 MATLAB 的 end）常以
# 低置信度命中 heading_level_*，会把章节计数器带偏（题注章节号随之错位）。
CODE_HEADING_PATTERN = re.compile(
    r"^(?:end|return|break|continue|else|case|default|\{|\}|\)|\])\s*[;.]?$",
    re.IGNORECASE,
)

# 伪一级标题：以“第X章——”起头的整句（常出现在正文叙述中）、以及 1) 2）这类
# 列表序号行。它们被模型判为 heading_level_1 会把题注章节号带偏。
PSEUDO_HEADING_PATTERNS = (
    re.compile(r"^第\s*[0-9一二三四五六七八九十百零〇]+\s*章\s*[—–\-－~～]"),
    re.compile(r"^\d+\s*[)）]"),
    # 图/表引用句、被拆出的题注残片（“如5.1 实物照片”“图2.1 系统框图”）
    re.compile(r"^\s*(?:如|见)\s*[图表]?\s*\d+\s*[.．\-—]\s*\d"),
    re.compile(r"^\s*[图表]\s*\d+\s*[.．\-—]\s*\d"),
)
# 以“第一章/第2章”开头但很长的标题是正文叙述（“第一章介绍研究目的……”）
_CHAPTER_LABEL_RE = re.compile(r"^第\s*[0-9一二三四五六七八九十百零〇]+\s*章")
_CHAPTER_LABEL_MAXLEN = 25
# 章内中文序号小节（一、方案确定 / 二 筛子的材料）
_CN_ITEM_RE = re.compile(r"^\s*[一二三四五六七八九十]+\s*[、.．\s　]")
# 真章节名（无编号时也不能当小节降级）
_CHAPTER_NAME_PREFIXES = (
    "绪论",
    "引言",
    "概述",
    "结论",
    "结束语",
    "总结",
    "致谢",
    "参考文献",
    "摘要",
    "目录",
    "附录",
)
# 纯数字编号的一级标题（“2 系统总体设计”），用于向下看二级标题编号是否同源
_NUMERIC_L1_RE = re.compile(r"^\s*(\d{1,2})(?:[、.．\s　])\s*\S")
_NUMERIC_CHILD_RE = re.compile(r"^\s*(\d{1,2})[.．]\d")
# 数字后直接跟中文的章节标题（“3核心工作部件设计”）
_DIGIT_CHAPTER_RE = re.compile(r"^\s*(\d{1,2})(?=[一-鿿])")
_LIST_ITEM_RE = re.compile(r"^\s*\d+\s*[)）]")


def _is_chapter_heading_text(text: str) -> bool:
    """粗判文本是否像一级章节标题（用于定位封面内容的边界）。"""
    if not text:
        return False
    if _CHAPTER_LABEL_RE.match(text) or _DIGIT_CHAPTER_RE.match(text):
        return True
    if _NUMERIC_L1_RE.match(text) and not _LIST_ITEM_RE.match(text):
        return True
    m = _CN_ITEM_RE.match(text)
    if m:
        rest = text[m.end() :].strip()
        return any(rest.startswith(p) for p in _CHAPTER_NAME_PREFIXES)
    return any(text.startswith(p) for p in _CHAPTER_NAME_PREFIXES)


def _find_first_chapter_index(result: list[dict]) -> int | None:
    for idx, item in enumerate(result):
        if item["category"] not in HEADING_LEVEL_CATS:
            continue
        text = (item.get("paragraph") or "").strip()
        if _is_chapter_heading_text(text):
            return idx
    return None


def _caption_chapter_conflict(result: list[dict], idx: int, number: int) -> bool:
    """候选章节标题后紧跟的题注章号与候选编号不一致。

    真章节通常在首个二级标题前不会出现题注；若出现且章号为 M≠N，说明该候选
    实际是当前章内的小节（“3 温度传感器模块”后面紧跟“图2.3 …”）。
    一旦先遇到二级标题，说明标题层级已经向下展开，不再据此判断。
    """
    for other in result[idx + 1 :]:
        cat = other["category"]
        if cat in ("heading_level_1", "heading_level_2"):
            return False
        if cat in ("caption_figure", "caption_table"):
            parsed = parse_caption_text((other.get("paragraph") or "").strip())
            chapter = parsed.get("chapter_num") if parsed else None
            if chapter is not None and chapter != number:
                return True
    return False


def _fix_known_categories(result: list[dict]) -> None:  # noqa C901
    """用已知文本模式修正摘要分类，并保护封面/声明内容。"""
    _mark_cover_as_other(result)
    # 第一个真章节之前的标题属于封面题名/声明页，不按一级标题校验格式。
    first_chapter_idx = _find_first_chapter_index(result)
    if first_chapter_idx:
        for item in result[:first_chapter_idx]:
            if item["category"] in HEADING_LEVEL_CATS:
                item["category"] = "other"
                item["comment"] = "封面/摘要前标题，跳过格式化"
                item["score"] = 1.0
                item["needs_review"] = False
    # 正文叙述句以“图5.1展示了……”开头时会被误判为题注；题注名很短，
    # 超过 60 字或带叙述动词/句号的长句多为正文，降级后不再占用题注顺序号。
    # 以句号收尾的短名词性题注（“图3.3 系统流程图。”）是正常写法，不能一律降级。
    for item in result:
        if item["category"] not in ("caption_figure", "caption_table"):
            continue
        text = (item.get("paragraph") or "").strip()
        narrative = len(text) > 60 or (
            re.search(r"[。！？!?]$", text)
            and (
                re.search(r"(所示|如图|见图|给出了|展示了|表示为|说明|描述|如下)", text)
                or re.match(
                    r"^[图表]\s*\d+(?:\s*[.．\-—]\s*\d+)?\s*(?:为|是|给出)",
                    text,
                )
            )
        )
        if narrative:
            item["category"] = "body_text"
            item["comment"] = "长句降级为正文（疑似题注叙述句）"
            item["score"] = 1.0
            item["needs_review"] = False
    # 预先判定哪些一级标题是“真·阿拉伯数字章节”（2 系统总体设计）。
    # “1 加速度传感器模块”这类章内小节编号会与后续二级标题编号不同源。
    numeric_chapter_ids: set[int] = set()
    for idx, item in enumerate(result):
        if item["category"] != "heading_level_1":
            continue
        text = (item.get("paragraph") or "").strip()
        if "章" in text[:6]:
            continue
        m = _NUMERIC_L1_RE.match(text)
        if not m:
            continue
        # 只看本章范围内的二级标题：没有二级标题的章节（“2 系统总体方案设计”）
        # 是正常的，不能因为下一章的 3.1 编号不同就把本章标题降级。
        chapter_end = next(
            (
                j
                for j in range(idx + 1, len(result))
                if result[j]["category"] == "heading_level_1"
            ),
            len(result),
        )
        next_l2 = next(
            (
                (other.get("paragraph") or "").strip()
                for other in result[idx + 1 : chapter_end]
                if other["category"] == "heading_level_2"
            ),
            None,
        )
        child = _NUMERIC_CHILD_RE.match(next_l2 or "")
        if not child or child.group(1) == m.group(1):
            numeric_chapter_ids.add(idx)
    first_numeric_chapter = min(numeric_chapter_ids) if numeric_chapter_ids else None

    last_numeric_chapter: int | None = None
    demoted: set[int] = set()
    for idx, item in enumerate(result):
        if item["category"] not in HEADING_LEVEL_CATS:
            continue
        text = (item.get("paragraph") or "").strip()

        def demote(reason: str, item=item, idx=idx) -> None:
            item["category"] = "body_text"
            item["comment"] = reason
            item["score"] = 1.0
            item["needs_review"] = False
            demoted.add(idx)

        if any(pat.match(text) for pat in PSEUDO_HEADING_PATTERNS):
            demote("伪标题（叙述句/列表序号）降级为正文")
            continue
        if item["category"] == "heading_level_1":
            # 以句号结尾的是叙述句（“……实物图如图2.4所示。”）
            if re.search(r"[。！？.!?]$", text):
                demote("叙述句降级为正文（句末标点）")
                continue
            # “第一章介绍研究目的与意义……”超长的是正文叙述，不是标题
            if _CHAPTER_LABEL_RE.match(text) and len(text) > _CHAPTER_LABEL_MAXLEN:
                demote("长句降级为正文（第X章开头的叙述句）")
                continue
            # 文档用阿拉伯数字编章时，“一、方案确定”是章内小节
            if numeric_chapter_ids and _CN_ITEM_RE.match(text):
                demote("中文序号小节降级为正文")
                continue
            # 已进入数字编章后，短小且无编号的一级标题多为小节名
            if (
                first_numeric_chapter is not None
                and idx > first_numeric_chapter
                and len(text) <= 15
                and not _NUMERIC_L1_RE.match(text)
                and not _CHAPTER_LABEL_RE.match(text)
                and not _CN_ITEM_RE.match(text)
                and not any(text.startswith(p) for p in _CHAPTER_NAME_PREFIXES)
                and any(j > idx for j in numeric_chapter_ids)
            ):
                demote("无编号短标题降级为正文（章内小节）")
                continue
        # 章内小节重新从 1 编号时（“1 加速度传感器模块”），模型常误判为
        # 一级标题。真章节编号必须递增，且标题后紧跟的题注章号必须自洽。
        l1_number: int | None = None
        if item["category"] == "heading_level_1":
            m_num = _NUMERIC_L1_RE.match(text) or _DIGIT_CHAPTER_RE.match(text)
            if m_num and "章" not in text[:6]:
                candidate = int(m_num.group(1))
                if (
                    last_numeric_chapter is not None
                    and candidate <= last_numeric_chapter
                ):
                    demote("章内编号小节降级为正文（章节编号未递增）")
                    continue
                if _caption_chapter_conflict(result, idx, candidate):
                    demote("章内编号小节降级为正文（题注章号与编号不符）")
                    continue
                l1_number = candidate
        # 真章节标题下方的首个二级标题编号必然同源（2 → 2.x）。
        if item["category"] == "heading_level_1" and idx not in numeric_chapter_ids:
            m = _NUMERIC_L1_RE.match(text)
            if m and "章" not in text[:6]:
                demote("章内编号小节降级为正文（编号与二级标题不同源）")
                continue
        if l1_number is not None:
            last_numeric_chapter = l1_number
        if item.get("score", 0.0) >= STATE_MIN_CONF:
            continue
        if CODE_HEADING_PATTERN.match(text):
            item["category"] = "body_text"
            item["comment"] = "代码行降级为正文（低置信度标题）"
            item["score"] = 1.0
            item["needs_review"] = False
    # 章节标题被模型判为正文（“3核心工作部件设计”“2 系统总体方案设计”）：
    # 编号不与已有章节号重复、标题较短时补回一级标题，避免按正文字号校验。
    existing_chapters = set()
    for item in result:
        if item["category"] != "heading_level_1":
            continue
        m = _NUMERIC_L1_RE.match((item.get("paragraph") or "").strip())
        if m:
            existing_chapters.add(int(m.group(1)))
    for idx, item in enumerate(result):
        if item["category"] != "body_text" or idx in demoted:
            continue
        text = (item.get("paragraph") or "").strip()
        if (
            not text
            or len(text) > 30
            or _LIST_ITEM_RE.match(text)
            or _NUMERIC_CHILD_RE.match(text)
        ):
            continue
        if re.search(r"[。！？.!?；;，,]$", text):
            continue
        m = _NUMERIC_L1_RE.match(text) or _DIGIT_CHAPTER_RE.match(text)
        if not m:
            continue
        try:
            number = int(m.group(1))
        except (TypeError, ValueError):
            continue
        rest = text[m.end() :].strip()
        if len(re.findall(r"[一-鿿]", rest)) < 4 or number in existing_chapters:
            continue
        item["category"] = "heading_level_1"
        item["comment"] = "章节标题补识别（编号与已有章节号不重复）"
        item["score"] = 1.0
        item["needs_review"] = False
        existing_chapters.add(number)

    abstract_patterns = [
        (r"^(摘要)\s*$", "abstract_chinese_title"),
        (r"^(Abstract)\s*$", "abstract_english_title"),
        (r"^摘要\s*[:：]", "abstract_chinese_title_content"),
        (r"^Abstract\s*[:：]?", "abstract_english_title_content"),
    ]
    # 摘要合并段修正：仅处理仍为 body_text 的段落
    for item in result:
        if item["category"] != "body_text":
            continue
        text = (item.get("paragraph") or "").strip()
        if not text:
            continue
        for pat, cat in abstract_patterns:
            if re.match(pat, text):
                item["category"] = cat
                item["comment"] = f"模式修正为 {cat}"
                item["score"] = 1.0
                item["needs_review"] = False
                break


def _fix_abstract_titles(result: list[dict]) -> None:  # noqa C901
    """摘要页标题回补：紧邻“摘要：/Abstract:”正文上方的独立标题行升格为摘要标题。

    模型常把摘要页的中/英文标题误判为 other（封面保护）/body_text/keywords_english，
    导致标题拿不到标题格式。此处从合并式摘要正文（*_title_content）向上找最近非空段，
    若其为标题样式（不太长、无句末/冒号标点、非摘要关键词行、非封面声明字段）则升格。
    """

    def title_key(value: str) -> str:
        return re.sub(r"\s+", "", value)

    # 摘要在封面之后单独成页时，页首常重复一遍论文题名（字号同封面）。
    # 这类行不是“摘要”标题，不应按摘要标题字号校验。
    cover_title_indices: dict[str, set[int]] = {}
    for i, item in enumerate(result):
        if item["category"] not in ("other", "document_title"):
            continue
        key = title_key(item.get("paragraph") or "")
        if key:
            cover_title_indices.setdefault(key, set()).add(i)

    for i, item in enumerate(result):
        target = ABSTRACT_CONTENT_TO_TITLE.get(item["category"])
        if not target:
            continue
        # 向上找最近的非空段（跳过空段/图片段）
        j = i - 1
        while j >= 0 and not (result[j].get("paragraph") or "").strip():
            j -= 1
        if j < 0:
            continue
        prev = result[j]
        prev_cat = prev["category"]
        prev_text = (prev.get("paragraph") or "").strip()
        if prev_cat not in ABSTRACT_TITLE_PREV_CATS:
            continue
        if len(prev_text) > ABSTRACT_TITLE_MAXLEN:
            continue
        if prev_text.endswith(("。", "！", "？", "；", ".", "!", "?", ";", "：", ":")):
            continue
        if re.match(r"^(摘\s*要|Abstract|关键词|keywords?)", prev_text, re.IGNORECASE):
            continue
        if COVER_MARKER_PATTERN.search(prev_text):
            continue
        prev["category"] = target
        prev["comment"] = f"摘要页标题（位置规则，原：{prev_cat}）"
        prev["score"] = 1.0
        prev["needs_review"] = False

    # 升格后的摘要标题若与封面题名相同，说明这是摘要页页首重复的论文题名，
    # 应跳过格式校验，而不是按“摘要”标题要求字号。
    for idx, item in enumerate(result):
        if item["category"] not in ("abstract_chinese_title", "abstract_english_title"):
            continue
        key = title_key((item.get("paragraph") or "").strip())
        if len(key) < 4:
            continue
        # 需要另一处封面/题名行同文，才判定为重复题名（该行自身不算）
        if not any(i != idx for i in cover_title_indices.get(key, ())):
            continue
        item["category"] = "document_title"
        item["comment"] = "摘要页重复的论文题名（按文档题名处理，不做格式校验）"
        item["score"] = 1.0
        item["needs_review"] = False


def _gate_abstract_content(result: list[dict]) -> None:
    """摘要内容段位置门控：前一段不是摘要标题/摘要内容时，回退为正文。

    摘要开场句与绪论首段词面同构（“随着…发展”“目前…”），模型仅凭文本无法
    区分；而摘要内容在文档中的位置是强约束——它只能紧跟摘要标题或上一段
    摘要内容。prev 不在允许集合内时回退为 body_text（绪论误吸入摘要）。
    """
    for i, item in enumerate(result):
        orig = item["category"]
        if orig not in ABSTRACT_CONTENT_GATE_CATS:
            continue
        text = (item.get("paragraph") or "").strip()
        # 合并式摘要段（“Abstract：正文……”“摘要：正文……”）自带标题前缀，
        # 不能因前一段不是摘要标题而回退为正文。
        merged = re.match(r"^(摘\s*要|Abstract)\s*[:：]", text, re.IGNORECASE)
        if merged:
            zh = "摘" in merged.group(1)
            item["category"] = (
                "abstract_chinese_title_content"
                if zh
                else "abstract_english_title_content"
            )
            item["comment"] = f"合并式摘要段（原：{orig}）"
            item["score"] = 1.0
            item["needs_review"] = False
            continue
        j = i - 1
        while j >= 0 and not (result[j].get("paragraph") or "").strip():
            j -= 1
        if j >= 0 and result[j]["category"] in ABSTRACT_CONTENT_GATE_ALLOWED_PREV:
            continue
        item["category"] = "body_text"
        prev_cat = result[j]["category"] if j >= 0 else None
        item["comment"] = f"摘要内容位置门控（原：{orig}，prev={prev_cat}）"
        item["score"] = 1.0  # 位置规则已确认修正，重置分数避免前端按 body_text 阈值误标
        item["needs_review"] = False


def _looks_like_thesis_title_en(text: str) -> bool:
    """英文论文题名（非 Abstract 标题行）。

    AbstractTitleEN 节点同时承担英文论文题名；这类题名较长、含英文单词、
    且不以 Abstract/Keywords 开头、不以句末标点结尾，不能按“文本必须等于
    abstract”降级，否则会被当成正文去校验正文格式。
    """
    return (
        len(text) >= 20
        and re.search(r"[A-Za-z]", text) is not None
        and not re.match(r"^(abstract|keywords?)\b", text, re.IGNORECASE)
        # 英文题名常以句点收尾，句点不算叙述句标志
        and not re.search(r"[。！？!?]$", text)
    )


def _fix_title_categories(result: list[dict]) -> None:  # noqa C901
    """按标题文本校订参考文献/致谢/摘要标题的识别结果。

    模型会把“感谢父母……”（致谢首段）、“冯勋.基于……”（参考文献条目）、
    英文论文题名等误判为对应章节的标题，置信度往往只有 0.5~0.6。标题的
    文本形式是强约束：去掉空白后必须恰为“致谢”“参考文献”“摘要”或
    “Abstract”。先按文本补出真标题，再把不匹配的伪标题降级到该章节正文，
    避免用标题字号（小二/三号）去要求普通正文，产生成片误报。
    """
    title_cats = {
        "acknowledgements_title": "致谢",
        "references_title": "参考文献",
        "abstract_chinese_title": "摘要",
        "abstract_english_title": "abstract",
    }

    def _normalized(text: str) -> str:
        return re.sub(r"\s+", "", (text or "")).rstrip(":：")

    # 合并式摘要（“摘要：正文……”“Abstract：正文……”）被模型标成标题时，
    # 转入 *_title_content 混合节点，标题前缀按标题样式、其余按正文样式。
    merged_pat = re.compile(r"^(摘\s*要|Abstract)\s*[:：]", re.IGNORECASE)
    merged_cats = {
        "abstract_chinese_title",
        "abstract_english_title",
        "body_text",
        "other",
    }
    for item in result:
        if item["category"] not in merged_cats:
            continue
        text = (item.get("paragraph") or "").strip()
        m = merged_pat.match(text)
        if not m:
            continue
        zh = "摘" in m.group(1)
        orig = item["category"]
        new_cat = (
            "abstract_chinese_title_content" if zh else "abstract_english_title_content"
        )
        item["category"] = new_cat
        item["comment"] = f"合并式摘要标题（原：{orig}）"
        item["score"] = 1.0
        item["needs_review"] = False

    # 第一遍：文本命中标题模式 → 直接认定为标题（TOC 条目已在前面排除）
    for item in result:
        norm = _normalized(item.get("paragraph") or "")
        for cat, label in title_cats.items():
            if norm.lower() == label:
                if item["category"] != cat:
                    item["category"] = cat
                    item["comment"] = f"标题文本修正：{item['category']} → {cat}"
                    item["score"] = 1.0
                    item["needs_review"] = False
                break

    # 第二遍：标题类别但文本不符 → 降级到对应正文
    ack_title_idx = next(
        (
            i
            for i, item in enumerate(result)
            if item["category"] == "acknowledgements_title"
        ),
        None,
    )
    for idx, item in enumerate(result):
        cat = item["category"]
        label = title_cats.get(cat)
        if label is None:
            continue
        text = (item.get("paragraph") or "").strip()
        if _normalized(text).lower() == label:
            continue
        if cat == "references_title":
            new_cat = (
                "references_content" if REF_ITEM_PATTERN.match(text) else "body_text"
            )
        elif cat == "acknowledgements_title":
            new_cat = (
                "acknowledgements_content"
                if ack_title_idx is not None and idx > ack_title_idx
                else "body_text"
            )
        elif cat == "abstract_english_title" and _looks_like_thesis_title_en(text):
            continue
        else:
            new_cat = "body_text"
        item["category"] = new_cat
        item["comment"] = f"标题文本不符，降级：{cat} → {new_cat}"
        item["score"] = 1.0
        item["needs_review"] = False

    # 合并式摘要转入 *_title_content 后，紧邻其上的独立标题行（英文论文题名等）
    # 需要再跑一次位置规则才能升格为摘要标题。
    _fix_abstract_titles(result)


def _fix_sequence(result: list[dict]) -> None:  # noqa C901
    """用段落类别间的合法转移关系修正序列错误。

    例如："参考文献"后面的 body_text 不可能是正文，应该是参考文献条目。
    """
    n = len(result)
    if n == 0:
        return

    def _text(i):
        return (result[i].get("paragraph") or "").strip()

    def _cat(i):
        return result[i].get("category", "")

    def _set(i, cat, reason):
        old = result[i]["category"]
        result[i]["category"] = cat
        result[i]["comment"] = f"{reason}（原：{old}）"
        result[i]["score"] = 1.0
        result[i]["needs_review"] = False

    # 规则1：关键词后面的 body_text → 如果是英文关键词附近，修正为英文关键词
    i = 0
    while i < n:
        if "keywords_chinese" in _cat(i) and i + 1 < n:
            j = i + 1
            while j < n and _cat(j) in ("body_text",):
                if re.search(r"Keywords?|KEY\s*WORDS", _text(j), re.IGNORECASE):
                    _set(j, "keywords_english", "关键词序列修正：英文关键词")
                    break
                j += 1
        i += 1

    # 规则2：独立"参考文献"行 → references_title
    for i in range(n):
        t = _text(i)
        if re.match(r"^参考文献\s*$", t) and _cat(i) != "references_title":
            _set(i, "references_title", "序列修正：参考文献标题")

    # 规则3：独立"致谢"行 → acknowledgements_title
    for i in range(n):
        t = _text(i)
        if re.match(r"^致\s*谢\s*$", t) and _cat(i) != "acknowledgements_title":
            _set(i, "acknowledgements_title", "序列修正：致谢标题")

    # 规则4：keywords + 后面紧跟 keyword-like 内容 → 标记为关键词
    for i in range(n):
        if "keywords" in _cat(i) and i + 1 < n and _cat(i + 1) == "body_text":
            t = _text(i + 1)
            # 含分号分隔的短词 → 关键词
            if re.search(r"[；;,]", t) and len(t) < 200:
                if "keyword" not in _cat(i + 1):
                    _set(i + 1, _cat(i), "序列修正：关键词延续")

    # 规则5（C6）：参考文献条目（references_title 之后，直到下一个顶层标题）
    in_refs = False
    for i in range(n):
        cat = _cat(i)
        if cat == "references_title":
            in_refs = True
            continue
        if cat in BODY_HEADINGS or cat == "acknowledgements_title":
            in_refs = False
            continue
        if in_refs and cat == "body_text" and REF_ITEM_PATTERN.match(_text(i)):
            _set(i, "references_content", "序列修正：参考文献条目")

    # 规则6：致谢正文（acknowledgements_title 之后，直到下一个顶层标题）
    # 模型可能漏判“致谢”标题，规则3 才补上，此时状态机已执行完毕，
    # 故在此把致谢正文补为 acknowledgements_content（与参考文献规则5 对称）。
    # 置信度门控（与状态机 STATE_MIN_CONF 一致）：源码里被误判的低置信
    # acknowledgements_title（如 "{"@0.39）不得触发，否则会把代码扫成致谢正文。
    in_ack = False
    for i in range(n):
        cat = _cat(i)
        if cat == "acknowledgements_title":
            in_ack = result[i].get("score", 0.0) >= STATE_MIN_CONF
            continue
        if cat in BODY_HEADINGS or cat == "references_title":
            in_ack = False
            continue
        if in_ack and cat == "body_text" and _text(i):
            _set(i, "acknowledgements_content", "序列修正：致谢正文")


_KEYWORD_LABEL_RE = re.compile(
    r"^\s*(?:关\s*键\s*词|key\s*words?|keywords)\s*[:：]?\s*", re.IGNORECASE
)


def _looks_like_keyword_list(text: str) -> bool:
    """关键词段应是短词列表，而非功能说明/代码行。

    模型会把 "K1（菜单键）：进入设置菜单；" 这类按键说明、以及源码行
    误判为关键词；判定后降级为正文，避免“关键词数量过少/末尾标点”误报。
    """
    has_label = bool(_KEYWORD_LABEL_RE.match(text))
    body = _KEYWORD_LABEL_RE.sub("", text).strip()
    if not body:
        return True
    if re.match(r"^(?://|/\*|#)", body):
        return False
    if re.search(r"[。！？!?：:（）(){}]", body):
        return False
    if re.search(r"\s=\s|[;{}]\s*$", body) and not re.search(r"[；;]\s*\S", body):
        return False
    if has_label:
        # 标签明确时即使不足两个词也保留，交给关键词数量规则报错
        return True
    parts = [p.strip() for p in re.split(r"[；;、,，]", body) if p.strip()]
    return len(parts) >= 2


def _fix_keyword_categories(result: list[dict]) -> None:
    """非词表的 keywords_* 段落降级为正文。"""
    for item in result:
        if item["category"] not in ("keywords_chinese", "keywords_english"):
            continue
        text = (item.get("paragraph") or "").strip()
        if not text or _looks_like_keyword_list(text):
            continue
        item["category"] = "body_text"
        item["comment"] = "非关键词列表，降级为正文"
        item["score"] = 1.0
        item["needs_review"] = False
