#! /usr/bin/env python
# -*- coding: utf-8 -*-
import tempfile
import os

from wordformat.config.config import init_config, get_config


def test_init_config():
    """测试初始化配置的功能"""
    # 使用示例配置文件
    config_path = "example/undergrad_thesis.yaml"
    
    # 初始化配置
    init_config(config_path)
    
    # 验证配置是否成功加载
    config = get_config()
    assert config is not None


def test_get_config():
    """测试获取配置的功能"""
    # 确保配置已初始化
    config_path = "example/undergrad_thesis.yaml"
    init_config(config_path)
    
    # 获取配置
    config = get_config()
    assert config is not None
    
    # 再次获取配置（应该返回缓存的配置）
    config2 = get_config()
    assert config2 is not None
    assert config2 == config
