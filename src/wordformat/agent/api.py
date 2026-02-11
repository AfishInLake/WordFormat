#! /usr/bin/env python
# @Time    : 2026/1/11 17:38
# @Author  : afish
# @File    : api.py
import json
from collections.abc import AsyncGenerator

from loguru import logger
from openai import AsyncOpenAI
from openai.types.chat import ChatCompletion, ChatCompletionChunk, ChatCompletionMessage
from openai.types.chat.chat_completion_message_tool_call import (
    ChatCompletionMessageToolCall,
)


class OpenAIAgent:
    def __init__(
        self,
        system_prompt: str,
        messageManager,
        model: str,
        baseurl,
        api_key: str = "",
        *args,
        **kwargs,
    ):
        self.total_tokens = 0
        self.total_completion_tokens = 0
        self.total_prompt_tokens = 0
        self.completion = None
        self.model = model
        self.client: AsyncOpenAI | None = None
        self.message_manager = messageManager
        self.client = AsyncOpenAI(api_key=api_key, base_url=baseurl, **kwargs)
        self.message_manager.add_system_message(system_prompt)

    async def get_response(
        self, stream: bool = False, response_format: str = "json", *args, **kwargs
    ) -> ChatCompletion | AsyncGenerator[ChatCompletionChunk, None]:
        """获取响应，支持流式"""
        api_params = {
            "model": self.model,
            "messages": self.message_manager.messages,
            "stream": stream,
        }
        extra_body = {
            "options": {
                "temperature": 0.0,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
                "num_predict": 96,
            }
        }

        match response_format:
            case "json":
                api_params["response_format"] = {"type": "json_object"}
                api_params["parallel_tool_calls"] = False
                api_params.pop("tools", None)

        response = await self.client.chat.completions.create(
            **api_params, extra_body=extra_body
        )
        # if not stream:
        #     logger.debug(
        #         f"API调用成功，token使用: {response.usage if hasattr(response, 'usage') else '未知'}"
        #     )
        return response

    async def get_function(self) -> list[ChatCompletionMessageToolCall] | None:
        """获取工具调用（非流式）"""
        if self.completion is None:
            self.completion = await self.get_response(stream=False)
        msg = self.completion.choices[0].message
        return msg.tool_calls if hasattr(msg, "tool_calls") and msg.tool_calls else None

    def _get_function_name(self, tool_call: ChatCompletionMessageToolCall) -> str:
        return tool_call.function.name

    def _get_arguments(self, tool_call: ChatCompletionMessageToolCall) -> dict:
        return json.loads(tool_call.function.arguments)

    async def get_message(self) -> ChatCompletionMessage:
        """获取完整消息（非流式）"""
        completion = await self.get_response(stream=False)
        return completion.choices[0].message

    async def response(
        self, user_prompt: str, stream: bool = False, *args, **kwargs
    ) -> AsyncGenerator | str:
        """
        处理用户输入。
        - 如果 stream=False：返回完整响应（兼容旧逻辑）
        - 如果 stream=True：返回流式生成器（供前端消费）
        """
        self.message_manager.add_user_message(user_prompt)

        if stream:
            return self._response_stream(*args, **kwargs)
        else:
            return await self._response_non_stream(*args, **kwargs)

    async def _response_non_stream(self, *args, **kwargs) -> str:
        """非流式完整响应（带工具调用循环）"""
        while True:
            # 获取API响应
            response = await self.get_response(stream=False, *args, **kwargs)  # noqa B026
            msg = response.choices[0].message
            # 更新 token 累计统计
            self._update_token_usage(getattr(response, "usage", None))
            # 最终响应
            if msg.content:
                self.message_manager.add_assistant_message(msg)
                return msg.content

            logger.warning("收到空响应")
            return ""

    async def _response_stream(self, *args, **kwargs) -> AsyncGenerator[str, None]:
        """流式响应"""
        # 获取初始响应（非流式）以检查工具调用
        response = await self.get_response(stream=False, *args, **kwargs)  # noqa B026
        msg = response.choices[0].message

        stream_response = await self.get_response(stream=True)

        # 收集并流式输出内容
        full_content = ""
        async for chunk in stream_response:
            delta = chunk.choices[0].delta
            if delta.content:
                full_content += delta.content
                yield delta.content

        # 添加最终助手消息
        self.message_manager.add_assistant_message(msg)

    def _update_token_usage(self, usage) -> None:
        """根据 usage 对象更新累计 token 计数"""
        if usage is None:
            return
        self.total_prompt_tokens += getattr(usage, "prompt_tokens", 0)
        self.total_completion_tokens += getattr(usage, "completion_tokens", 0)
        self.total_tokens += getattr(usage, "total_tokens", 0)

    def print_token_usage(self) -> None:
        """打印 token 用量到日志"""
        logger.debug(
            f"累计使用 {self.total_tokens} tokens，其中 {self.total_prompt_tokens} tokens 用于提示，"
            f"{self.total_completion_tokens} tokens 用于生成"
        )
