#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

import re

from docx import Document
from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.settings import BATCH_SIZE
from wordformat.utils import get_paragraph_numbering_text, get_paragraph_xml_fingerprint


def _para_contains_image(para) -> bool:
    """检查段落是否包含内联图片（w:drawing）。"""
    from docx.oxml.ns import qn

    for r in para._element.findall(qn("w:r")):
        if r.find(qn("w:drawing")) is not None:
            return True
    return False


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
                has_image = _para_contains_image(para)
                result[i] = {
                    "category": "figure_image" if has_image else "body_text",
                    "score": 1.0,
                    "comment": "图片段落" if has_image else "空段落",
                    "paragraph": "",
                    "fingerprint": get_paragraph_xml_fingerprint(para),
                }

        # 对非空段进行批量 AI 推理
        for i in range(0, len(texts_for_ai), BATCH_SIZE):
            batch_texts = texts_for_ai[i : i + BATCH_SIZE]
            batch_indices = text_indices[i : i + BATCH_SIZE]

            try:
                batch_results = onnx_batch_infer(batch_texts)
            except Exception as e:
                logger.error(f"批量推理失败，降级到单条处理: {e}")
                batch_results = [onnx_single_infer(t) for t in batch_texts]

            for idx, text, pred in zip(
                batch_indices, batch_texts, batch_results, strict=False
            ):
                tag = pred["label"]
                score = pred["score"]
                para_obj = all_paras[idx]
                response = {
                    "category": tag,
                    "score": score,
                    "comment": f"置信度：{score:.4f}",
                    "paragraph": text,
                    "fingerprint": get_paragraph_xml_fingerprint(para_obj),
                }
                if score < 0.6:
                    response["category"] = "body_text"
                    response["comment"] = (
                        f"原预测标签为 '{tag}'，置信度 {score:.4f} < 0.6，已强制设为 'body_text'"
                    )
                result[idx] = response

        assert all(r is not None for r in result), "存在未处理的段落"
        # 后处理：已知模式的段落强制修正分类
        _fix_known_categories(result)
        return result


def _fix_known_categories(result: list[dict]) -> None:
    """用已知文本模式修正常见 AI 分类错误。"""
    abstract_patterns = [
        (r"^(摘要)\s*$", "abstract_chinese_title"),
        (r"^(Abstract)\s*$", "abstract_english_title"),
        (r"^摘要\s*[:：]", "abstract_chinese_title_content"),
        (r"^Abstract\s*[:：]?", "abstract_english_title_content"),
    ]
    # 找到第一个摘要段落的位置，之前的内容全部标为 other（封面/声明）
    abstract_start = None
    for i, item in enumerate(result):
        text = (item.get("paragraph") or "").strip()
        if re.match(r"^(摘要|Abstract)", text):
            abstract_start = i
            break
    if abstract_start is not None:
        for i in range(abstract_start):
            result[i]["category"] = "other"
            result[i]["comment"] = (
                f"摘要前内容，跳过检查（原：{result[i]['category']}）"
            )
            result[i]["score"] = 1.0
    # 摘要修正
    for item in result:
        if item["category"] == "other":
            continue
        text = (item.get("paragraph") or "").strip()
        if item["category"] == "body_text":
            for pat, cat in abstract_patterns:
                if re.match(pat, text):
                    item["category"] = cat
                    item["comment"] = f"模式修正为 {cat}"
                    break
    # 序列修正：用类别相邻关系纠正明显不合理分类
    _fix_sequence(result)


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
