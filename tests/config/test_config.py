#! /usr/bin/env python
# -*- coding: utf-8 -*-
import tempfile
import os
import pytest

from wordformat.config.config import init_config, get_config, clear_config, LazyConfig, ConfigNotLoadedError


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


def test_clear_config():
    """测试清除配置的功能"""
    # 初始化并加载配置
    config_path = "example/undergrad_thesis.yaml"
    init_config(config_path)
    get_config()
    
    # 清除配置
    clear_config()
    
    # 验证配置已清除
    lazy_config = LazyConfig()
    assert not lazy_config._loaded
    assert lazy_config._config is None


def test_config_not_initialized():
    """测试配置未初始化时的错误情况"""
    # 清除配置
    clear_config()
    
    # 尝试获取配置，应该抛出异常
    with pytest.raises(ConfigNotLoadedError):
        get_config()


def test_lazy_config_singleton():
    """测试LazyConfig的单例模式"""
    # 创建两个实例
    instance1 = LazyConfig()
    instance2 = LazyConfig()
    
    # 验证它们是同一个实例
    assert instance1 is instance2


def test_lazy_config_config_path_property():
    """测试LazyConfig的config_path属性"""
    lazy_config = LazyConfig()
    config_path = "example/undergrad_thesis.yaml"
    
    # 初始化配置路径
    lazy_config.init(config_path)
    
    # 验证config_path属性
    assert lazy_config.config_path == config_path


def test_lazy_config_clear():
    """测试LazyConfig的clear方法"""
    lazy_config = LazyConfig()
    config_path = "example/undergrad_thesis.yaml"
    
    # 初始化并加载配置
    lazy_config.init(config_path)
    lazy_config.load()
    
    # 清除配置
    lazy_config.clear()
    
    # 验证配置已清除
    assert not lazy_config._loaded
    assert lazy_config._config is None
    assert lazy_config._config_path is None


def test_lazy_config_load_failure():
    """测试配置加载失败的情况"""
    lazy_config = LazyConfig()
    
    # 初始化一个不存在的配置文件路径
    non_existent_path = "non_existent_config.yaml"
    lazy_config.init(non_existent_path)
    
    # 尝试加载配置，应该抛出异常
    with pytest.raises(Exception):
        lazy_config.load()
