#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

import re

from docx import Document
from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.settings import BATCH_SIZE
from wordformat.utils import get_paragraph_numbering_text, para_contains_image

# ===== 分级置信度阈值（C2）=====
# 低于阈值不再硬砍为 body_text，而是保留原标签并标记 needs_review 供人工复核。
CONF_THRESHOLDS = {
    "abstract_chinese_title": 0.3,
    "abstract_english_title": 0.3,
    "abstract_chinese_title_content": 0.4,
    "abstract_english_title_content": 0.4,
    "references_title": 0.4,
    "acknowledgements_title": 0.4,
    "heading_fulu": 0.4,
    "keywords_chinese": 0.4,
    "keywords_english": 0.4,
    "default": 0.6,
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
        # 收集所有段落（含空段），空段/图片段直接标记，不走 AI 推理
        all_paras = list(self.document.paragraphs)
        text_indices = []
        texts_for_ai = []
        for idx, para in enumerate(all_paras):
            text = para.text.strip()
            if not text:
                continue
            numbering_text = get_paragraph_numbering_text(para)
            full_text = f"{numbering_text} {text}" if numbering_text else para.text
            texts_for_ai.append(full_text)
            text_indices.append(idx)

        result = [None] * len(all_paras)
        for i, para in enumerate(all_paras):
            if i not in text_indices:
                has_image = para_contains_image(para)
                result[i] = {
                    "category": "figure_image" if has_image else "body_text",
                    "score": 1.0,
                    "comment": "图片段落" if has_image else "空段落",
                    "paragraph": "",
                    "needs_review": False,
                }

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
        # → 8) 章节状态机（置信度门控）→ 9) 序列修正
        _fix_document_title(result)
        _fix_abstract_en_title(result)
        _fix_known_categories(result)
        _fix_abstract_titles(result)
        _gate_abstract_content(result)
        _apply_footer(result)
        _fix_toc(result)
        _neutralize_appendix(result)
        _apply_section_state(result)
        _fix_sequence(result)
        return result


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


def _neutralize_appendix(result: list[dict]) -> None:
    """附录区中性化：可信 heading_fulu（附录）之后，把摘要/关键词/致谢等前置类
    标签强制回 body_text。

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
        if in_appendix and cat in APPENDIX_NEUTRALIZE:
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


def _fix_known_categories(result: list[dict]) -> None:
    """用已知文本模式修正摘要分类，并保护封面/声明内容。"""
    _mark_cover_as_other(result)
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


def _fix_abstract_titles(result: list[dict]) -> None:
    """摘要页标题回补：紧邻“摘要：/Abstract:”正文上方的独立标题行升格为摘要标题。

    模型常把摘要页的中/英文标题误判为 other（封面保护）/body_text/keywords_english，
    导致标题拿不到标题格式。此处从合并式摘要正文（*_title_content）向上找最近非空段，
    若其为标题样式（不太长、无句末/冒号标点、非摘要关键词行、非封面声明字段）则升格。
    """
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
        j = i - 1
        while j >= 0 and not (result[j].get("paragraph") or "").strip():
            j -= 1
        if j >= 0 and result[j]["category"] in ABSTRACT_CONTENT_GATE_ALLOWED_PREV:
            continue
        item["category"] = "body_text"
        prev_cat = result[j]["category"] if j >= 0 else None
        item["comment"] = f"摘要内容位置门控（原：{orig}，prev={prev_cat}）"
        item["needs_review"] = False


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
