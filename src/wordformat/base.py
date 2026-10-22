#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

from docx import Document
from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.domains.postprocess import get_postprocess
from wordformat.settings import BATCH_SIZE
from wordformat.utils import get_paragraph_numbering_text, para_contains_image


class DocxBase:
    def __init__(self, docx_file, configpath=None, domain: str = "thesis"):
        self.re_dict = {}
        self.docx_file = docx_file
        self.domain = domain
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
        # 收集所有段落（含空段），空段/图片段直接标记
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

        # 对非空段进行批量推理（跨批次维护 PREV 上下文状态链）
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
                result[idx] = response

        assert all(r is not None for r in result), "存在未处理的段落"

        # 领域分类后处理：阈值应用 + 领域规则链由领域模块整链执行并注册到
        # 框架注册表（domains/postprocess.py）。领域未装配（未注册）时跳过，
        # 顺序约束与阈值常量随论文规则位于 domains/thesis/postprocess.py
        postprocessor = get_postprocess(self.domain)
        if postprocessor is not None:
            postprocessor(result)
        return result
