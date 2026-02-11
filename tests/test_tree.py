#! /usr/bin/env python
# -*- coding: utf-8 -*-
from wordformat.tree import Tree, Stack, print_tree
from wordformat.rules.node import TreeNode


def test_tree_creation():
    """测试Tree类的创建"""
    # 创建Tree实例
    tree = Tree("root")
    assert tree is not None
    assert tree.root.value == "root"
    assert not tree.is_empty()


def test_tree_traversal():
    """测试Tree类的遍历方法"""
    # 创建测试树
    tree = Tree("root")
    # 添加子节点
    child1 = TreeNode("child1")
    child2 = TreeNode("child2")
    grandchild1 = TreeNode("grandchild1")
    
    tree.root.add_child_node(child1)
    tree.root.add_child_node(child2)
    child1.add_child_node(grandchild1)
    
    # 测试前序遍历
    preorder_result = list(tree.preorder())
    assert len(preorder_result) == 4
    assert preorder_result == ["root", "child1", "grandchild1", "child2"]
    
    # 测试后序遍历
    postorder_result = list(tree.postorder())
    assert len(postorder_result) == 4
    assert postorder_result == ["grandchild1", "child1", "child2", "root"]
    
    # 测试层序遍历
    level_order_result = list(tree.level_order())
    assert len(level_order_result) == 4
    assert level_order_result == ["root", "child1", "child2", "grandchild1"]


def test_tree_find_by_condition():
    """测试Tree类的find_by_condition方法"""
    # 创建测试树
    tree = Tree({"name": "root", "category": "top"})
    # 添加子节点
    child1 = TreeNode({"name": "child1", "value": 1, "fingerprint": "fp1"})
    child2 = TreeNode({"name": "child2", "value": 2, "fingerprint": "fp2"})
    grandchild1 = TreeNode({"name": "grandchild1", "value": 3, "fingerprint": "fp3"})
    
    tree.root.add_child_node(child1)
    tree.root.add_child_node(child2)
    child1.add_child_node(grandchild1)
    
    # 测试查找
    def condition(x):
        return x.get("value") == 2
    result = tree.find_by_condition(condition)
    assert result is not None
    assert result.value["name"] == "child2"
    
    # 测试查找不存在的节点
    def condition2(x):
        return x.get("value") == 999
    result = tree.find_by_condition(condition2)
    assert result is None


def test_tree_height_and_size():
    """测试Tree类的height和size方法"""
    # 创建测试树
    tree = Tree({"name": "root", "category": "top"})
    # 添加子节点
    child1 = TreeNode({"name": "child1", "fingerprint": "fp1"})
    child2 = TreeNode({"name": "child2", "fingerprint": "fp2"})
    grandchild1 = TreeNode({"name": "grandchild1", "fingerprint": "fp3"})
    great_grandchild1 = TreeNode({"name": "great_grandchild1", "fingerprint": "fp4"})
    
    tree.root.add_child_node(child1)
    tree.root.add_child_node(child2)
    child1.add_child_node(grandchild1)
    grandchild1.add_child_node(great_grandchild1)
    
    # 测试高度
    assert tree.height() >= 1  # 确保高度计算正确
    
    # 测试大小
    assert tree.size() == 5  # 总节点数


def test_stack_operations():
    """测试Stack类的操作"""
    # 创建Stack实例
    stack = Stack()
    assert stack.is_empty()
    assert stack.size() == 0
    
    # 测试push
    stack.push(1)
    stack.push(2)
    stack.push(3)
    assert not stack.is_empty()
    assert stack.size() == 3
    
    # 测试peek
    assert stack.peek() == 3
    assert stack.size() == 3  # peek不改变栈大小
    
    # 测试peek_safe
    assert stack.peek_safe() == 3
    
    # 测试pop
    assert stack.pop() == 3
    assert stack.size() == 2
    assert stack.pop() == 2
    assert stack.size() == 1
    assert stack.pop() == 1
    assert stack.size() == 0
    assert stack.is_empty()
    
    # 测试空栈操作
    assert stack.peek_safe() is None
    try:
        stack.pop()
        assert False, "空栈pop应抛出异常"
    except IndexError:
        pass
    try:
        stack.peek()
        assert False, "空栈peek应抛出异常"
    except IndexError:
        pass
    
    # 测试clear
    stack.push(1)
    stack.push(2)
    assert stack.size() == 2
    stack.clear()
    assert stack.is_empty()
    assert stack.size() == 0


def test_print_tree():
    """测试print_tree函数"""
    # 创建测试树
    root = TreeNode({"name": "root", "category": "top"})
    child1 = TreeNode({"name": "child1", "category": "child", "fingerprint": "fp1"})
    child2 = TreeNode({"name": "child2", "category": "child", "fingerprint": "fp2"})
    grandchild1 = TreeNode({"name": "grandchild1", "category": "grandchild", "fingerprint": "fp3"})
    
    root.add_child_node(child1)
    root.add_child_node(child2)
    child1.add_child_node(grandchild1)
    
    # 测试print_tree函数（不应抛出异常）
    try:
        print_tree(root)
    except Exception as e:
        assert False, f"print_tree() 抛出异常: {e}"
