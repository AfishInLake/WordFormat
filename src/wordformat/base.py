#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py

from docx import Document
from loguru import logger

from wordformat.agent.onnx_infer import onnx_batch_infer, onnx_single_infer
from wordformat.media import ImageRegistry
from wordformat.settings import BATCH_SIZE
from wordformat.utils import get_paragraph_numbering_text, get_paragraph_xml_fingerprint


def _has_drawing(para) -> bool:
    """检测段落是否包含 w:drawing 元素。"""

    for r in para._element:
        tag = r.tag.split("}")[-1] if "}" in r.tag else r.tag
        if tag != "r":
            continue
        for child in r:
            ctag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if ctag == "drawing":
                return True
    return False


def _extract_image_info(para) -> dict:
    """从图片段落提取声明式数据（blob + sha256 + 尺寸），不再返回 raw XML。"""
    if not _has_drawing(para):
        return {}
    return ImageRegistry.extract_from_paragraph(para)


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

    def parse(self, image_dir: str | None = None) -> list[dict]:
        paragraphs = []
        paragraph_objects = []
        result = []
        # 收集所有非空段落及纯图片段落
        for para in self.document.paragraphs:
            text = para.text.strip()
            has_drawing = _has_drawing(para)
            if not text and not has_drawing:
                continue  # 跳过空段落（无文本无图片）
            # 拼接自动编号文字（para.text 不包含编号）
            numbering_text = get_paragraph_numbering_text(para)
            if numbering_text:
                full_text = f"{numbering_text} {text}" if text else numbering_text
            else:
                full_text = para.text
            paragraphs.append(full_text)
            paragraph_objects.append(para)

        # 按批次BATCH_SIZE进行批量推理
        for i in range(0, len(paragraphs), BATCH_SIZE):
            batch_texts = paragraphs[i : i + BATCH_SIZE]
            batch_paras = paragraph_objects[i : i + BATCH_SIZE]

            try:
                batch_results = onnx_batch_infer(batch_texts)
            except Exception as e:
                logger.error(f"批量推理失败，降级到单条处理: {e}")
                batch_results = [onnx_single_infer(text) for text in batch_texts]

            # batch 结果，应用后处理逻辑
            for _j, (full_text, para_obj, pred) in enumerate(
                zip(batch_texts, batch_paras, batch_results, strict=False)
            ):
                has_drawing = _has_drawing(para_obj)
                text = para_obj.text.strip()

                # 纯图片段落（无文字）——跳过 ONNX 分类
                if not text and has_drawing:
                    tag, score = "image", 1.0
                else:
                    tag = pred["label"]
                    score = pred["score"]

                response = {
                    "category": tag,
                    "paragraph": para_obj.text,
                    "meta": {
                        "score": score,
                        "comment": f"置信度：{score:.4f}",
                        "original_text": full_text,
                        "index": i + _j,
                        "fingerprint": get_paragraph_xml_fingerprint(para_obj),
                    },
                }

                # 声明式图片数据（blob sha256 + 尺寸），不再存 raw XML
                if has_drawing:
                    info = _extract_image_info(para_obj)
                    if info.get("sha256"):
                        response["meta"]["image_source"] = info.get("source", "")
                        response["meta"]["image_sha256"] = info.get("sha256", "")
                        response["meta"]["width_emu"] = info.get("width_emu", 0)
                        response["meta"]["height_emu"] = info.get("height_emu", 0)
                        response["meta"]["alignment"] = info.get("alignment", "center")
                        # 保存图片文件到 image_dir，af 时可直接引用
                        if image_dir and info.get("blob"):
                            import os as _os

                            _os.makedirs(image_dir, exist_ok=True)
                            sha = info["sha256"]
                            ext = (
                                _os.path.splitext(info.get("source", ".png"))[-1]
                                or ".png"
                            )
                            fname = f"{sha[:16]}{ext}"
                            fpath = _os.path.join(image_dir, fname)
                            if not _os.path.exists(fpath):
                                with open(fpath, "wb") as _f:
                                    _f.write(info["blob"])
                            response["meta"]["image_path"] = fpath

                # 置信度过低处理
                if score < 0.6:
                    response["category"] = "body_text"
                    response["meta"]["comment"] = (
                        f"原预测标签为 '{tag}'，置信度 {score:.4f} < 0.6，已强制设为 'body_text'"
                    )
                result.append(response)

        return result
