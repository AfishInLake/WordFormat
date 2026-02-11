#! /usr/bin/env python
# -*- coding: utf-8 -*-
import pytest
from unittest import mock

from wordformat.word_structure.node_factory import create_node
from wordformat.word_structure.tree_builder import DocumentTreeBuilder
from wordformat.word_structure.utils import find_and_modify_first, promote_bodytext_in_subtrees_of_type
from wordformat.rules.node import FormatNode
from wordformat.rules.body import BodyText


class TestNodeFactory:
    """测试node_factory.py"""
    
    def test_create_node_with_valid_category(self):
        """测试创建有效类别的节点"""
        # 模拟CATEGORY_TO_CLASS中的类
        mock_class = mock.MagicMock(spec=FormatNode)
        mock_instance = mock.MagicMock()
        mock_class.return_value = mock_instance
        mock_instance.load_config = mock.MagicMock()
        
        # 模拟node_factory模块中的CATEGORY_TO_CLASS
        with mock.patch('wordformat.word_structure.node_factory.CATEGORY_TO_CLASS', {'test_category': mock_class}):
            item = {'category': 'test_category', 'content': 'test', 'fingerprint': 'test_fingerprint'}
            result = create_node(item=item, level=1, config={})
            assert result == mock_instance
            mock_class.assert_called_once_with(value=item, expected_rule={}, level=1)
            mock_instance.load_config.assert_called_once_with({})
    
    def test_create_node_with_invalid_category(self):
        """测试创建无效类别的节点"""
        # 模拟CATEGORY_TO_CLASS为空
        with mock.patch('wordformat.word_structure.settings.CATEGORY_TO_CLASS', {}):
            item = {'category': 'invalid_category', 'content': 'test'}
            result = create_node(item=item, level=1, config={})
            assert result is None
    
    def test_create_node_missing_category(self):
        """测试创建缺少类别的节点"""
        item = {'content': 'test'}  # 缺少category
        with pytest.raises(ValueError):
            create_node(item=item, level=1, config={})


class TestDocumentTreeBuilder:
    """测试tree_builder.py"""
    
    def test_init(self):
        """测试初始化"""
        builder = DocumentTreeBuilder()
        assert hasattr(builder, 'stack')
    
    def test_create_root_node(self):
        """测试创建根节点"""
        builder = DocumentTreeBuilder()
        root = builder._create_root_node()
        assert isinstance(root, FormatNode)
        assert root.value['category'] == 'top'
        assert root.level == 0
    
    def test_determine_level(self):
        """测试确定层级"""
        builder = DocumentTreeBuilder()
        # 模拟LEVEL_MAP
        with mock.patch('wordformat.word_structure.tree_builder.LEVEL_MAP', {'test_category': 1}):
            level = builder._determine_level('test_category')
            assert level == 1
            # 测试未找到的情况
            level = builder._determine_level('unknown_category')
            assert level == 999
    
    def test_is_heading_category(self):
        """测试是否为标题类别"""
        builder = DocumentTreeBuilder()
        # 模拟HEADING_CATEGORIES
        with mock.patch.object(builder, 'HEADING_CATEGORIES', {'test_category': mock.MagicMock()}):
            assert builder._is_heading_category('test_category') is True
            assert builder._is_heading_category('unknown_category') is False
    
    @mock.patch('wordformat.word_structure.node_factory.create_node')
    def test_create_node_from_item(self, mock_create_node):
        """测试从项目创建节点"""
        builder = DocumentTreeBuilder()
        mock_node = mock.MagicMock()
        mock_create_node.return_value = mock_node
        
        item = {'category': 'test_category'}
        result = builder._create_node_from_item(item)
        assert result == mock_node
        mock_create_node.assert_called_once()
    
    def test_attach_heading_node(self):
        """测试附加标题节点"""
        builder = DocumentTreeBuilder()
        # 创建根节点并压入栈
        root = builder._create_root_node()
        builder.stack.push(root)
        
        # 创建测试节点
        test_node = FormatNode(value={'category': 'test', 'paragraph': 'test', 'fingerprint': 'test_fingerprint'}, expected_rule={}, level=1)
        
        # 附加节点
        builder._attach_heading_node(test_node)
        assert len(root.children) == 1
        assert root.children[0] == test_node
    
    def test_attach_body_node(self):
        """测试附加正文节点"""
        builder = DocumentTreeBuilder()
        # 创建根节点并压入栈
        root = builder._create_root_node()
        builder.stack.push(root)
        
        # 创建测试节点
        test_node = FormatNode(value={'category': 'test', 'paragraph': 'test', 'fingerprint': 'test_fingerprint'}, expected_rule={}, level=2)
        
        # 附加节点
        builder._attach_body_node(test_node)
        assert len(root.children) == 1
        assert root.children[0] == test_node
    
    @mock.patch('wordformat.word_structure.node_factory.create_node')
    def test_build_tree(self, mock_create_node):
        """测试构建树"""
        builder = DocumentTreeBuilder()
        mock_node = mock.MagicMock()
        mock_node.level = 1
        mock_create_node.return_value = mock_node
        
        items = [{'category': 'test_category'}]
        root = builder.build_tree(items)
        assert isinstance(root, FormatNode)
        assert len(root.children) == 1


class TestUtils:
    """测试utils.py"""
    
    def test_find_and_modify_first(self):
        """测试查找并修改第一个满足条件的节点"""
        # 创建测试树
        root = FormatNode(value={'category': 'top'}, expected_rule={}, level=0)
        child1 = FormatNode(value={'category': 'child1', 'fingerprint': 'child1_fingerprint'}, expected_rule={}, level=1)
        child2 = FormatNode(value={'category': 'child2', 'fingerprint': 'child2_fingerprint'}, expected_rule={}, level=1)
        root.add_child_node(child1)
        root.add_child_node(child2)
        
        # 测试查找条件
        def condition(node):
            return node.value.get('category') == 'child1'
        
        result = find_and_modify_first(root, condition)
        assert result == child1
    
    def test_promote_bodytext_in_subtrees_of_type(self):
        """测试升级子树中的BodyText节点"""
        # 创建测试树
        root = FormatNode(value={'category': 'top'}, expected_rule={}, level=0)
        parent_node = FormatNode(value={'category': 'parent', 'fingerprint': 'parent_fingerprint'}, expected_rule={}, level=1)
        body_node = BodyText(value={'category': 'body', 'paragraph': 'test', 'fingerprint': 'body_fingerprint'}, expected_rule={}, level=2)
        
        root.add_child_node(parent_node)
        parent_node.add_child_node(body_node)
        
        # 创建目标类型
        class TargetType(BodyText):
            pass
        
        # 执行升级
        promote_bodytext_in_subtrees_of_type(root, FormatNode, TargetType)
        
        # 验证升级结果
        assert isinstance(parent_node.children[0], TargetType)
