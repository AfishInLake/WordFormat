"""docx 文档级操作。"""

import os

from docx.oxml.ns import qn
from loguru import logger


def remove_all_numbering(doc):
    """
    强制解除样式与列表的绑定
    Args:
        doc:
    Returns:

    """
    title_style_names = ["Heading 1", "Heading 2", "Heading 3"]

    for style_name in title_style_names:
        if style_name in doc.styles:
            style = doc.styles[style_name]
            style_element = style._element

            # 删除 <w:pPr> 中的 numPr（样式级别的编号）
            pPr = style_element.find(qn("w:pPr"))
            if pPr is not None:
                numPr = pPr.find(qn("w:numPr"))
                if numPr is not None:
                    pPr.remove(numPr)

                # 可选：也删除 outlineLvl（大纲级别，有时触发编号）
                outlineLvl = pPr.find(qn("w:outlineLvl"))
                if outlineLvl is not None:
                    pPr.remove(outlineLvl)

            logger.debug(f"已解除样式 '{style_name}' 的编号绑定")


def ensure_directory_exists(path):
    """
    检查路径是否存在，如果不存在则创建对应的文件夹。

    参数:
        path (str): 需要检查或创建的文件夹路径

    说明:
        - 如果路径已存在且是文件夹，则不做任何操作
        - 如果路径不存在，则递归创建所有必需的父目录
        - 如果路径存在但是是文件，则抛出 ValueError
    """
    if os.path.exists(path):
        if not os.path.isdir(path):
            raise ValueError(f"路径已存在但不是文件夹：'{path}'")
    else:
        os.makedirs(path, exist_ok=True)
        logger.info(f"已创建文件夹：'{path}'")


def para_contains_image(para) -> bool:
    """检查段落是否包含内联图片（w:drawing）。"""
    from docx.oxml.ns import qn

    for r in para._element.findall(qn("w:r")):
        if r.find(qn("w:drawing")) is not None:
            return True
    return False
