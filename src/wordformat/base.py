#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

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
        return result
