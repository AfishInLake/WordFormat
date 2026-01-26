#! /usr/bin/env python
# @Time    : 2025/12/22 21:47
# @Author  : afish
# @File    : DocxBase.py
import json

from loguru import logger
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_fixed

from src.agent.api import OpenAIAgent
from src.agent.message import MessageManager
from src.settings import API_KEY, MODEL, MODEL_URL
from src.utils import get_paragraph_xml_fingerprint


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
            api_key=API_KEY,
        )

    async def parse(self) -> list[dict]:
        # 正则提取
        # bert打置信度
        # 置信度低和正则无法提取的交由llm处理
        result = []
        for paragraph in self.document.paragraphs:
            # 跳过空段落
            if not paragraph.text:
                continue
            try:
                response = await self.parse_paragraph(paragraph.text)
            except Exception as e:
                response = {
                    "category": "body_text",
                    "comment": str(e),
                    "paragraph": paragraph.text,
                }

            response["fingerprint"] = get_paragraph_xml_fingerprint(paragraph)
            logger.info(response)
            if response["category"] == "body_text":
                break
            result.append(response)
        self.base_agent.print_token_usage()
        return result

    @retry(
        retry=retry_if_exception_type(Exception),
        stop=stop_after_attempt(5),
        wait=wait_fixed(5),
        reraise=True,
    )
    async def parse_paragraph(self, paragraph: str):
        self.base_agent.message_manager.clear()  # 清空消息
        # 创建基础agent
        response = await self.base_agent.response(
            f"""当前段落：{paragraph}\n当前段落是正文还是标题？如果是标题是几级标题？""",
            stream=False,
            response_format="json",
        )
        jsondata = json.loads(response)
        result_dict = {}
        for key in ["category"]:
            if key not in jsondata:
                raise Exception(f"{key} not found in response")
            result_dict[key] = jsondata.get(key)
        jsondata["paragraph"] = paragraph

        return jsondata
