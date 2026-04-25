"""
这是测试库的接口测试文件
"""
import os
import json
from wordformat.set_tag import set_tag_main
from loguru import logger

docx_path = r'word/毕业设计说明书.docx'
json_path = r'output/毕业设计说明书.json'
config_path = r'example/undergrad_thesis.yaml'

if __name__ == '__main__':
    """
    生成段落标记的入口函数
    """
    json_data = set_tag_main(
        docx_path=docx_path, configpath=config_path
    )
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)
    logger.info(f"保存成功：{json_data}")

    """
    
    """
