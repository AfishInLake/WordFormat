#! /usr/bin/env python
# -*- coding: utf-8 -*-
import pytest
from unittest import mock
from pathlib import Path

from wordformat.rules.node import TreeNode, FormatNode
from wordformat.config.datamodel import NodeConfigRoot


class TestTreeNode:
    """测试TreeNode类"""
    
    def test_init(self):
        """测试初始化"""
        node = TreeNode("test_value")
        assert node.value == "test_value"
        assert node.children == []
    
    def test_config_property(self):
        """测试config属性"""
        node = TreeNode("test_value")
        assert node.config == {}
    
    def test_load_config_valid(self):
        """测试加载有效配置"""
        class TestNode(TreeNode):
            NODE_TYPE = "test.node"
        
        node = TestNode("test_value")
        full_config = {
            "test": {
                "node": {
                    "key1": "value1",
                    "key2": "value2"
                }
            }
        }
        node.load_config(full_config)
        assert node.config == {"key1": "value1", "key2": "value2"}
    
    def test_load_config_invalid(self):
        """测试加载无效配置"""
        class TestNode(TreeNode):
            NODE_TYPE = "test.node"
        
        node = TestNode("test_value")
        # 测试路径不存在
        full_config = {"other": {"node": {}}}
        node.load_config(full_config)
        assert node.config == {}
        
        # 测试非字典配置
        full_config = "not a dict"
        node.load_config(full_config)
        assert node.config == {}
    
    def test_set_fingerprint_with_category_top(self):
        """测试设置指纹（category为top）"""
        node = TreeNode({"category": "top"})
        # 不应设置fingerprint属性
        assert not hasattr(node, "fingerprint")
    
    def test_set_fingerprint_with_fingerprint(self):
        """测试设置指纹（有fingerprint键）"""
        node = TreeNode({"fingerprint": "test_fingerprint"})
        assert node.fingerprint == "test_fingerprint"
    
    def test_set_fingerprint_without_fingerprint(self):
        """测试设置指纹（无fingerprint键）"""
        with pytest.raises(ValueError):
            TreeNode({"key": "value"})
    
    def test_add_child(self):
        """测试添加子节点"""
        node = TreeNode("parent")
        child = node.add_child("child")
        assert len(node.children) == 1
        assert child.value == "child"
    
    def test_add_child_node(self):
        """测试直接添加子节点"""
        node = TreeNode("parent")
        child_node = TreeNode("child")
        node.add_child_node(child_node)
        assert len(node.children) == 1
        assert node.children[0] == child_node
    
    def test_repr(self):
        """测试字符串表示"""
        node = TreeNode("test_value")
        assert repr(node) == "TreeNode(test_value)"


class TestFormatNode:
    """测试FormatNode类"""
    
    def setup_class(self):
        """设置测试类"""
        class TestConfig:
            pass
        
        class TestFormatNode(FormatNode):
            NODE_TYPE = "test.format"
            CONFIG_MODEL = TestConfig
            
            def _base(self, doc, p, r):
                pass
        
        self.TestFormatNode = TestFormatNode
    
    def test_init(self):
        """测试初始化"""
        paragraph = mock.MagicMock()
        node = self.TestFormatNode("test_value", 1, paragraph)
        assert node.value == "test_value"
        assert node.level == 1
        assert node.paragraph == paragraph
        assert node.expected_rule is None
    
    def test_pydantic_config_not_loaded(self):
        """测试获取未加载的Pydantic配置"""
        node = self.TestFormatNode("test_value", 1)
        with pytest.raises(ValueError):
            node.pydantic_config
    
    @mock.patch('builtins.open', new_callable=mock.mock_open, read_data='{}')
    def test_load_yaml_config(self, mock_open):
        """测试加载YAML配置"""
        # 由于NodeConfigRoot需要完整的配置结构，这里会失败，所以我们只测试异常情况
        with pytest.raises(ValueError):
            self.TestFormatNode.load_yaml_config("test_config.yaml")
    
    @mock.patch('builtins.open', new_callable=mock.mock_open, read_data='{}')
    def test_load_yaml_config_file_not_found(self, mock_open):
        """测试加载不存在的YAML配置文件"""
        mock_open.side_effect = FileNotFoundError
        with pytest.raises(FileNotFoundError):
            self.TestFormatNode.load_yaml_config("non_existent.yaml")
    
    def test_update_paragraph(self):
        """测试更新段落"""
        node = self.TestFormatNode("test_value", 1)
        new_paragraph = mock.MagicMock()
        node.update_paragraph(new_paragraph)
        assert node.paragraph == new_paragraph
    
    def test_check_format(self):
        """测试检查格式"""
        node = self.TestFormatNode("test_value", 1)
        doc = mock.MagicMock()
        # 不应抛出异常
        node.check_format(doc)
    
    def test_apply_format(self):
        """测试应用格式"""
        node = self.TestFormatNode("test_value", 1)
        doc = mock.MagicMock()
        # 不应抛出异常
        node.apply_format(doc)
    
    def test_add_comment(self):
        """测试添加注释"""
        node = self.TestFormatNode("test_value", 1)
        doc = mock.MagicMock()
        runs = mock.MagicMock()
        node.add_comment(doc, runs, "测试注释")
        doc.add_comment.assert_called_once_with(runs=runs, text="测试注释", author="论文解析器", initials="afish")
    
    def test_add_comment_empty(self):
        """测试添加空注释"""
        node = self.TestFormatNode("test_value", 1)
        doc = mock.MagicMock()
        runs = mock.MagicMock()
        node.add_comment(doc, runs, "   ")
        doc.add_comment.assert_not_called()

    def test_load_yaml_config_validation_error(self):
        """测试加载YAML配置验证错误"""
        with mock.patch('builtins.open', new_callable=mock.mock_open, read_data='{}'):
            with pytest.raises(ValueError):
                self.TestFormatNode.load_yaml_config("test_config.yaml")

    def test_load_config_unknown_config_type(self):
        """测试加载未知类型的配置"""
        # 创建一个使用未知配置类型的FormatNode子类
        class UnknownConfig:
            pass
        
        class TestFormatNodeUnknown(FormatNode):
            NODE_TYPE = "test.unknown"
            CONFIG_MODEL = UnknownConfig
            
            def _base(self, doc, p, r):
                pass
        
        node = TestFormatNodeUnknown("test_value", 1)
        mock_config = mock.MagicMock()
        
        with pytest.raises(ValueError):
            node.load_config(mock_config)

    def test_format_node_base_not_implemented(self):
        """测试FormatNode的_base方法未实现"""
        # 创建一个未实现_base方法的FormatNode子类
        class TestFormatNodeNoBase(FormatNode):
            NODE_TYPE = "test.nobase"
            CONFIG_MODEL = self.TestFormatNode.CONFIG_MODEL
        
        node = TestFormatNodeNoBase("test_value", 1)
        doc = mock.MagicMock()
        
        with pytest.raises(NotImplementedError):
            node._base(doc, True, True)
