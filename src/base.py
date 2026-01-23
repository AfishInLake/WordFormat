#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py
import json

from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

from src.agent.api import OpenAIAgent
from src.agent.message import MessageManager
from src.settings import API_KEY, MODEL, MODEL_URL
from src.utils import get_paragraph_fingerprint


class DocxBase:
    def __init__(self, docx_file, system_prompt):
        from docx import Document
        self.docx_file = docx_file
        self.document = Document(docx_file)

        self.base_agent = OpenAIAgent(
            system_prompt=system_prompt,
            messageManager=MessageManager(),
            model=MODEL,
            baseurl=MODEL_URL,
            api_key=API_KEY
        )

    async def parse(self) -> list[dict]:
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
                    'paragraph': paragraph.text,

                }

            response['fingerprint'] = get_paragraph_fingerprint(paragraph)
            print(response)
            result.append(response)
        return result

    @retry(
        retry=retry_if_exception_type(Exception),
        stop=stop_after_attempt(5),
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
        result_dict = {}
        for key in ['category', 'comment']:
            if key not in jsondata:
                raise Exception(f"{key} not found in response")
            result_dict[key] = jsondata.get(key)
        jsondata['paragraph'] = paragraph
        self.base_agent.message_manager.clear()  # 清空消息
        return jsondata


if __name__ == '__main__':
    from docx import Document

    # 创建一个新文档
    # doc = Document()
    # p = doc.add_paragraph("这是一段有注释的文字。")
    # run = p.runs[0]  # 获取第一个 run
    # doc.add_comment(run, text="这是一个注释！", author="作者", initials="ZS")
    # doc.save("output.docx")
