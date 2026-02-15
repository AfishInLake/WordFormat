#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.config.config import get_config, init_config
from wordformat.settings import BATCH_SIZE
from wordformat.utils import get_paragraph_xml_fingerprint


class DocxBase:
    def __init__(self, docx_file, configpath):
        from docx import Document

        self.re_dict = {}
        self.docx_file = docx_file
        self.document = Document(docx_file)
        init_config(configpath)
        try:
            self.config_model = get_config()  # 首次调用：触发load()
            self.config = self.config_model.model_dump()
            logger.info("配置文件验证通过")
        except Exception as e:
            logger.error(f"配置加载失败: {str(e)}")
            raise

    def parse(self) -> list[dict]:
        paragraphs = []
        paragraph_objects = []  # 保存原始 paragraph 对象，用于 fingerprint 和提前终止
        # 第一步：收集所有非空段落及其对象
        for para in self.document.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(para.text)
                paragraph_objects.append(para)

        result = []

        # 第二步：按批次（每批 128 条）进行批量推理
        batch_size = int(BATCH_SIZE)
        for i in range(0, len(paragraphs), batch_size):
            batch_texts = paragraphs[i : i + batch_size]
            batch_paras = paragraph_objects[i : i + batch_size]

            try:
                # 批量推理（关键修改点）
                batch_results = onnx_batch_infer(batch_texts)
            except Exception as e:
                # 整个 batch 推理失败？降级为逐条处理（可选）
                logger.error(f"批量推理失败，降级到单条处理: {e}")
                batch_results = [onnx_single_infer(text) for text in batch_texts]

            # 第三步：逐条处理 batch 结果，应用后处理逻辑
            for _j, (text, para_obj, pred) in enumerate(
                zip(batch_texts, batch_paras, batch_results, strict=False)
            ):
                tag = pred["预测标签"]
                score = pred["预测概率"]

                response = {
                    "category": tag,
                    "score": score,
                    "comment": f"置信度：{score:.4f}",
                    "paragraph": text,
                    "fingerprint": get_paragraph_xml_fingerprint(para_obj),
                }

                # 置信度过低处理
                if score < 0.6:
                    response["category"] = "body_text"
                    response["comment"] = (
                        f"原预测标签为 '{tag}'，置信度 {score:.4f} < 0.6，已强制设为 'body_text'"
                    )

                logger.info(response)
                result.append(response)

                # 遇到 heading_fulu 则提前终止整个解析
                if response["category"] == "heading_fulu":
                    return result  # 👈 立即返回，不再处理后续段落

        return result
