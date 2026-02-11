#! /usr/bin/env python
# -*- coding: utf-8 -*-
import pytest
from unittest import mock
import numpy as np

from wordformat.agent.message import MessageManager
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

