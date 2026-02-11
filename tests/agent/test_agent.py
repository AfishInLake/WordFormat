#! /usr/bin/env python
# -*- coding: utf-8 -*-
import pytest
from unittest import mock
import numpy as np

from wordformat.agent.message import MessageManager
from wordformat.agent.api import OpenAIAgent
from wordformat.agent.onnx_infer import (
    _get_model_paths,
    _get_best_onnx_providers,
    _load_model,
    onnx_single_infer,
    onnx_batch_infer,
    safe_batch_infer
)


class TestMessageManager:
    """测试MessageManager类"""
    
    def test_init(self):
        """测试初始化"""
        manager = MessageManager()
        assert manager.messages == []
    
    def test_add_message(self):
        """测试添加消息"""
        manager = MessageManager()
        manager.add_message("user", "测试消息")
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "user"
        assert manager.messages[0]["content"] == "测试消息"
    
    def test_add_dict_message(self):
        """测试添加字典消息"""
        manager = MessageManager()
        manager.add_dict_message({"role": "user", "content": "测试消息"})
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "user"
        assert manager.messages[0]["content"] == "测试消息"
    
    def test_add_user_message(self):
        """测试添加用户消息"""
        manager = MessageManager()
        manager.add_user_message("测试消息")
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "user"
        assert manager.messages[0]["content"] == "测试消息"
    
    def test_add_system_message(self):
        """测试添加系统消息"""
        manager = MessageManager()
        manager.add_system_message("测试系统消息")
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "system"
        assert manager.messages[0]["content"] == "测试系统消息"
    
    def test_add_tool_message(self):
        """测试添加工具消息"""
        manager = MessageManager()
        manager.add_tool_message("测试工具消息", "tool_call_id_123")
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "tool"
        assert manager.messages[0]["content"] == "测试工具消息"
        assert manager.messages[0]["tool_call_id"] == "tool_call_id_123"
    
    def test_add_assistant_message(self):
        """测试添加助手消息"""
        manager = MessageManager()
        # 测试字符串消息
        manager.add_assistant_message("测试助手消息")
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "assistant"
        assert manager.messages[0]["content"] == "测试助手消息"
        
        # 测试字典消息
        manager.add_assistant_message({"role": "assistant", "content": "测试助手消息2"})
        assert len(manager.messages) == 2
        assert manager.messages[1]["role"] == "assistant"
        assert manager.messages[1]["content"] == "测试助手消息2"
    
    def test_reset_messages(self):
        """测试重置消息"""
        manager = MessageManager()
        manager.add_user_message("测试消息")
        assert len(manager.messages) == 1
        manager.reset_messages()
        assert len(manager.messages) == 0
    
    def test_clear(self):
        """测试清空消息（保留系统消息）"""
        manager = MessageManager()
        manager.add_system_message("测试系统消息")
        manager.add_user_message("测试用户消息")
        assert len(manager.messages) == 2
        manager.clear()
        assert len(manager.messages) == 1
        assert manager.messages[0]["role"] == "system"


class TestOpenAIAgent:
    """测试OpenAIAgent类"""
    
    def test_init(self):
        """测试初始化"""
        message_manager = MessageManager()
        system_prompt = "测试系统提示"
        agent = OpenAIAgent(
            system_prompt=system_prompt,
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        assert agent.model == "gpt-3.5-turbo"
        assert agent.message_manager == message_manager
        assert len(agent.message_manager.messages) == 1
        assert agent.message_manager.messages[0]["role"] == "system"
        assert agent.message_manager.messages[0]["content"] == system_prompt
    
    @pytest.mark.asyncio
    @mock.patch('wordformat.agent.api.AsyncOpenAI')
    async def test_get_response(self, mock_async_openai):
        """测试获取响应"""
        # 模拟AsyncOpenAI客户端
        mock_client = mock_async_openai.return_value
        mock_response = mock.MagicMock()
        # 使用asyncio.Future来模拟异步返回值
        import asyncio
        future = asyncio.Future()
        future.set_result(mock_response)
        mock_client.chat.completions.create.return_value = future
        
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试非流式响应
        response = await agent.get_response(stream=False, response_format="json")
        assert response == mock_response
        mock_client.chat.completions.create.assert_called_once()
    
    @pytest.mark.asyncio
    @mock.patch('wordformat.agent.api.AsyncOpenAI')
    async def test_response_non_stream(self, mock_async_openai):
        """测试非流式响应"""
        # 模拟AsyncOpenAI客户端
        mock_client = mock_async_openai.return_value
        mock_response = mock.MagicMock()
        mock_message = mock.MagicMock()
        mock_message.content = "测试响应内容"
        mock_response.choices = [mock.MagicMock(message=mock_message)]
        # 使用asyncio.Future来模拟异步返回值
        import asyncio
        future = asyncio.Future()
        future.set_result(mock_response)
        mock_client.chat.completions.create.return_value = future
        
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试非流式响应
        response = await agent.response("测试用户输入", stream=False)
        assert response == "测试响应内容"
        assert len(agent.message_manager.messages) == 3  # 系统消息 + 用户消息 + 助手消息
    
    def test_update_token_usage(self):
        """测试更新token使用情况"""
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试更新token使用情况
        mock_usage = mock.MagicMock()
        mock_usage.prompt_tokens = 10
        mock_usage.completion_tokens = 5
        mock_usage.total_tokens = 15
        agent._update_token_usage(mock_usage)
        assert agent.total_prompt_tokens == 10
        assert agent.total_completion_tokens == 5
        assert agent.total_tokens == 15
    
    def test_print_token_usage(self):
        """测试打印token使用情况"""
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试打印token使用情况（不应抛出异常）
        agent.print_token_usage()
    
    def test_get_function_name(self):
        """测试获取函数名称"""
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 模拟tool_call对象
        mock_tool_call = mock.MagicMock()
        mock_tool_call.function.name = "test_function"
        
        # 测试获取函数名称
        function_name = agent._get_function_name(mock_tool_call)
        assert function_name == "test_function"
    
    def test_get_arguments(self):
        """测试获取函数参数"""
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 模拟tool_call对象
        mock_tool_call = mock.MagicMock()
        mock_tool_call.function.arguments = '{"param1": "value1", "param2": "value2"}'
        
        # 测试获取函数参数
        arguments = agent._get_arguments(mock_tool_call)
        assert arguments == {"param1": "value1", "param2": "value2"}
    
    @pytest.mark.asyncio
    @mock.patch('wordformat.agent.api.AsyncOpenAI')
    async def test_get_function(self, mock_async_openai):
        """测试获取工具调用"""
        # 模拟AsyncOpenAI客户端
        mock_client = mock_async_openai.return_value
        mock_response = mock.MagicMock()
        mock_message = mock.MagicMock()
        mock_message.tool_calls = [mock.MagicMock()]
        mock_response.choices = [mock.MagicMock(message=mock_message)]
        # 使用asyncio.Future来模拟异步返回值
        import asyncio
        future = asyncio.Future()
        future.set_result(mock_response)
        mock_client.chat.completions.create.return_value = future
        
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试获取工具调用
        tool_calls = await agent.get_function()
        assert tool_calls is not None
        assert len(tool_calls) == 1
    
    @pytest.mark.asyncio
    @mock.patch('wordformat.agent.api.AsyncOpenAI')
    async def test_get_message(self, mock_async_openai):
        """测试获取完整消息"""
        # 模拟AsyncOpenAI客户端
        mock_client = mock_async_openai.return_value
        mock_response = mock.MagicMock()
        mock_message = mock.MagicMock()
        mock_response.choices = [mock.MagicMock(message=mock_message)]
        # 使用asyncio.Future来模拟异步返回值
        import asyncio
        future = asyncio.Future()
        future.set_result(mock_response)
        mock_client.chat.completions.create.return_value = future
        
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试获取完整消息
        message = await agent.get_message()
        assert message == mock_message
    
    @pytest.mark.asyncio
    @mock.patch('wordformat.agent.api.AsyncOpenAI')
    async def test_response_empty_content(self, mock_async_openai):
        """测试空响应内容"""
        # 模拟AsyncOpenAI客户端
        mock_client = mock_async_openai.return_value
        mock_response = mock.MagicMock()
        mock_message = mock.MagicMock()
        mock_message.content = None  # 空响应
        mock_response.choices = [mock.MagicMock(message=mock_message)]
        # 使用asyncio.Future来模拟异步返回值
        import asyncio
        future = asyncio.Future()
        future.set_result(mock_response)
        mock_client.chat.completions.create.return_value = future
        
        message_manager = MessageManager()
        agent = OpenAIAgent(
            system_prompt="测试系统提示",
            messageManager=message_manager,
            model="gpt-3.5-turbo",
            baseurl="http://localhost:8000",
            api_key="test_api_key"
        )
        
        # 测试空响应
        response = await agent.response("测试用户输入", stream=False)
        assert response == ""


class TestONNXInfer:
    """测试ONNX推理模块"""
    
    def test_get_model_paths(self):
        """测试获取模型路径"""
        paths = _get_model_paths()
        assert isinstance(paths, dict)
        assert "onnx" in paths
        assert "tokenizer" in paths
        assert "id2label" in paths
    
    def test_get_best_onnx_providers(self):
        """测试获取最佳ONNX提供者"""
        # 测试CUDA提供者
        with mock.patch('onnxruntime.get_available_providers', return_value=['CUDAExecutionProvider', 'CPUExecutionProvider']):
            providers = _get_best_onnx_providers()
            assert providers == ['CUDAExecutionProvider']
        
        # 测试DirectML提供者
        with mock.patch('onnxruntime.get_available_providers', return_value=['DmlExecutionProvider', 'CPUExecutionProvider']):
            providers = _get_best_onnx_providers()
            assert providers == ['DmlExecutionProvider']
        
        # 测试CPU提供者
        with mock.patch('onnxruntime.get_available_providers', return_value=['CPUExecutionProvider']):
            providers = _get_best_onnx_providers()
            assert providers == ['CPUExecutionProvider']
        
        # 测试异常情况
        with mock.patch('onnxruntime.get_available_providers', side_effect=Exception('Test error')):
            providers = _get_best_onnx_providers()
            assert providers == ['CPUExecutionProvider']
    
    @mock.patch('wordformat.agent.onnx_infer._tokenizer', None)
    @mock.patch('wordformat.agent.onnx_infer._ort_sess', None)
    @mock.patch('wordformat.agent.onnx_infer._id2label', None)
    @mock.patch('wordformat.agent.onnx_infer._get_model_paths')
    @mock.patch('wordformat.agent.onnx_infer._get_best_onnx_providers')
    @mock.patch('onnxruntime.InferenceSession')
    @mock.patch('tokenizers.Tokenizer.from_file')
    def test_load_model(self, mock_tokenizer_from_file, mock_inference_session, mock_get_best_providers, mock_get_model_paths):
        """测试加载模型"""
        # 模拟返回值
        mock_paths = {
            'onnx': 'test_model.onnx',
            'tokenizer': 'test_tokenizer.json',
            'id2label': 'test_id2label.json'
        }
        mock_get_model_paths.return_value = mock_paths
        mock_get_best_providers.return_value = ['CPUExecutionProvider']
        
        # 模拟文件打开
        mock_file = mock.MagicMock()
        mock_file.__enter__.return_value.read.return_value = '{"0": "label0", "1": "label1"}'
        
        with mock.patch('builtins.open', return_value=mock_file):
            _load_model()
            # 验证调用
            mock_tokenizer_from_file.assert_called_once_with(mock_paths['tokenizer'])
            mock_inference_session.assert_called_once()
    
    @mock.patch('wordformat.agent.onnx_infer._tokenizer')
    @mock.patch('wordformat.agent.onnx_infer._ort_sess')
    @mock.patch('wordformat.agent.onnx_infer._id2label')
    @mock.patch('wordformat.agent.onnx_infer._load_model')
    def test_onnx_single_infer(self, mock_load_model, mock_id2label, mock_ort_sess, mock_tokenizer):
        """测试单条推理"""
        # 模拟tokenizer
        mock_encoded = mock.MagicMock()
        mock_encoded.ids = [1, 2, 3]
        mock_encoded.attention_mask = [1, 1, 1]
        mock_encoded.type_ids = [0, 0, 0]
        mock_tokenizer.encode.return_value = mock_encoded
        
        # 模拟ort_sess
        mock_logits = np.array([[0.5, -0.5]])
        mock_ort_sess.run.return_value = [mock_logits]
        
        # 模拟id2label
        mock_id2label.__getitem__.return_value = "label0"
        
        # 测试单条推理
        text = "测试文本"
        result = onnx_single_infer(text)
        assert isinstance(result, dict)
        assert "label" in result
        assert "score" in result
    
    @mock.patch('wordformat.agent.onnx_infer._tokenizer')
    @mock.patch('wordformat.agent.onnx_infer._ort_sess')
    @mock.patch('wordformat.agent.onnx_infer._id2label')
    @mock.patch('wordformat.agent.onnx_infer._load_model')
    def test_onnx_single_infer_error(self, mock_load_model, mock_id2label, mock_ort_sess, mock_tokenizer):
        """测试单条推理错误处理"""
        # 模拟tokenizer
        mock_encoded = mock.MagicMock()
        mock_encoded.ids = [1, 2, 3]
        mock_encoded.attention_mask = [1, 1, 1]
        mock_encoded.type_ids = [0, 0, 0]
        mock_tokenizer.encode.return_value = mock_encoded
        
        # 模拟ort_sess.run抛出异常
        mock_ort_sess.run.side_effect = Exception('Test error')
        
        # 测试错误处理
        text = "测试文本"
        result = onnx_single_infer(text)
        assert isinstance(result, dict)
        assert result["label"] == ""
        assert result["score"] == 0.0
    
    @mock.patch('wordformat.agent.onnx_infer._tokenizer')
    @mock.patch('wordformat.agent.onnx_infer._ort_sess')
    @mock.patch('wordformat.agent.onnx_infer._id2label')
    @mock.patch('wordformat.agent.onnx_infer._load_model')
    def test_onnx_batch_infer(self, mock_load_model, mock_id2label, mock_ort_sess, mock_tokenizer):
        """测试批量推理"""
        # 模拟tokenizer
        mock_encoded = mock.MagicMock()
        mock_encoded.ids = [1, 2, 3]
        mock_encoded.attention_mask = [1, 1, 1]
        mock_encoded.type_ids = [0, 0, 0]
        mock_tokenizer.encode.return_value = mock_encoded
        
        # 模拟ort_sess
        mock_logits = np.array([[0.5, -0.5], [0.3, -0.3]])
        mock_ort_sess.run.return_value = [mock_logits]
        
        # 模拟id2label
        mock_id2label.get.side_effect = lambda x, default: f"label{x}"
        
        # 测试批量推理
        texts = ["测试文本1", "测试文本2"]
        results = onnx_batch_infer(texts)
        assert isinstance(results, list)
        assert len(results) == 2
    
    @mock.patch('wordformat.agent.onnx_infer._tokenizer')
    @mock.patch('wordformat.agent.onnx_infer._ort_sess')
    @mock.patch('wordformat.agent.onnx_infer._id2label')
    @mock.patch('wordformat.agent.onnx_infer._load_model')
    def test_onnx_batch_infer_error(self, mock_load_model, mock_id2label, mock_ort_sess, mock_tokenizer):
        """测试批量推理错误处理"""
        # 模拟tokenizer
        mock_encoded = mock.MagicMock()
        mock_encoded.ids = [1, 2, 3]
        mock_encoded.attention_mask = [1, 1, 1]
        mock_encoded.type_ids = [0, 0, 0]
        mock_tokenizer.encode.return_value = mock_encoded
        
        # 模拟ort_sess.run抛出异常
        mock_ort_sess.run.side_effect = Exception('Test error')
        
        # 测试错误处理
        texts = ["测试文本1", "测试文本2"]
        results = onnx_batch_infer(texts)
        assert isinstance(results, list)
        assert len(results) == 2
        assert all(["原始文本" in result for result in results])
    
    @mock.patch('wordformat.agent.onnx_infer.onnx_batch_infer')
    def test_safe_batch_infer(self, mock_onnx_batch_infer):
        """测试安全批量推理"""
        # 模拟onnx_batch_infer，根据输入长度返回对应结果
        def mock_batch_infer(texts):
            return [{"label": f"test{i}"} for i, _ in enumerate(texts)]
        
        mock_onnx_batch_infer.side_effect = mock_batch_infer
        
        # 测试安全批量推理
        texts = ["测试文本1", "测试文本2"]
        results = safe_batch_infer(texts, max_batch_size=1)
        assert isinstance(results, list)
        assert len(results) == 2
    
    def test_onnx_batch_infer_empty(self):
        """测试批量推理（空输入）"""
        results = onnx_batch_infer([])
        assert isinstance(results, list)
        assert len(results) == 0
    
    def test_safe_batch_infer_empty(self):
        """测试安全批量推理（空输入）"""
        results = safe_batch_infer([])
        assert isinstance(results, list)
        assert len(results) == 0

