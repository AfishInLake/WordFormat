#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py
import json
from pathlib import Path
from typing import Dict, Any

import yaml
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

from src.agent.api import OpenAIAgent
from src.agent.message import MessageManager


class Style:
    """样式"""

    def __init__(self, path: str | Path):
        """
        Args:
            path: yaml配置加载路径
        """
        self.style = self.load_yaml_config(path)

    def load_yaml_config(self, file_path: str | Path) -> Dict[str, Any]:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"{file_path} not found")
        with open(path, 'r', encoding='utf-8') as f:
            try:
                config = yaml.safe_load(f)
                if config is None:
                    config = {}
                return config
            except yaml.YAMLError as e:
                raise ValueError(f"格式错误: {file_path}") from e


class DocxBase:
    def __init__(self, docx_file, system_prompt):
        from docx import Document
        self.docx_file = docx_file
        self.document = Document(docx_file)

        self.base_agent = OpenAIAgent(
            system_prompt=system_prompt,
            messageManager=MessageManager(),
            model="qwen3-4b-no-think",
            baseurl="http://localhost:11434/v1",
        )

    async def parse(self, rules: Style) -> list[dict]:
        result = []
        for paragraph in self.document.paragraphs:
            # 跳过空段落
            if not paragraph.text:
                continue
            try:
                response = await self.parse_paragraph(paragraph.text)
                if response['category'] == 'heading_fulu':
                    break
            except Exception as e:
                response = {
                    'category': 'body_text',
                    'comment': str(e),
                    'paragraph': paragraph.text
                }
            result.append(response)
        return result

    @retry(
        retry=retry_if_exception_type(Exception),
        stop=stop_after_attempt(3),
        wait=wait_fixed(5)
    )
    async def parse_paragraph(self, paragraph: str):
        # 创建基础agent
        response = await self.base_agent.response(
            f"""当前段落：{paragraph}\n当前段落是正文还是标题？如果是标题是几级标题？""",
            stream=False,
            response_format='json'
        )
        jsondata = json.loads(response)
        for key in ['category', 'comment']:
            if key not in jsondata:
                raise ValueError(f"{key} not found in response")
        jsondata['paragraph'] = paragraph
        self.base_agent.message_manager.clear()  # 清空消息
        print(jsondata)
        return jsondata


if __name__ == '__main__':
    from docx import Document

    # 创建一个新文档
    # doc = Document()
    # p = doc.add_paragraph("这是一段有注释的文字。")
    # run = p.runs[0]  # 获取第一个 run
    # doc.add_comment(run, text="这是一个注释！", author="作者", initials="ZS")
    # doc.save("output.docx")
