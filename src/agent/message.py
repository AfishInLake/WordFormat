#! /usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2026/1/11 17:39
# @Author  : afish
# @File    : message.py
import threading
from typing import Union, Dict


class MessageManager:

    def __init__(self):
        self._messages = []
        self._instance_lock = threading.Lock()  # 实例级操作锁

    def add_message(self, role: str, content: str):
        """添加角色信息"""
        with self._instance_lock:
            self._messages.append({
                "role": role,
                "content": content
            })

    def add_dict_message(self, content):
        """保存对话历史或上下文信息"""
        with self._instance_lock:
            self._messages.append(content)

    def add_user_message(self, content: str):
        """添加角色信息"""
        self.add_dict_message(
            {
                'role': 'user',
                'content': content
            }
        )

    def add_assistant_message(self, msg: Union[str, Dict]):
        """安全添加AI消息"""
        if isinstance(msg, dict):
            self.add_dict_message(msg)
            return
        content = msg.content if hasattr(msg, 'content') else msg
        tool_calls = getattr(msg, 'tool_calls', None)
        message = {"role": "assistant", "content": content}
        if tool_calls:
            message["tool_calls"] = [
                {
                    "id": call.id,
                    "type": call.type,
                    "function": {
                        "name": call.function.name,
                        "arguments": call.function.arguments
                    }
                }
                for call in tool_calls
            ]

        self.add_dict_message(message)

    def add_tool_message(self, content: str, tool_call_id):
        """添加工具信息"""
        self.add_dict_message(
            {
                'role': 'tool',
                'content': content,
                'tool_call_id': tool_call_id
            }
        )

    def add_system_message(self, content: str):
        """添加系统信息"""
        self.add_dict_message(
            {
                'role': 'system',
                'content': content
            }
        )

    def get_messages(self):
        """获取信息"""
        with self._instance_lock:
            return self._messages.copy()

    def reset_messages(self):
        """清空信息"""
        with self._instance_lock:
            self._messages.clear()

    @property
    def messages(self):
        """返回信息列表"""
        with self._instance_lock:
            return self._messages.copy()

    def clear(self):
        self._messages = list(filter(lambda msg: msg.get('role') == 'system', self._messages))
