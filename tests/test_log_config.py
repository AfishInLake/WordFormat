#! /usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
import logging
from unittest import mock

from wordformat.log_config import fix_std_streams, InterceptHandler, setup_uvicorn_loguru


def test_fix_std_streams():
    """测试fix_std_streams函数"""
    # 保存原始的stdout和stderr
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    
    try:
        # 测试stdout为空的情况
        sys.stdout = None
        fix_std_streams()
        assert sys.stdout is not None
        
        # 测试stderr为空的情况
        sys.stderr = None
        fix_std_streams()
        assert sys.stderr is not None
        
        # 测试正常情况（stdout和stderr不为空）
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        fix_std_streams()
        assert sys.stdout == original_stdout
        assert sys.stderr == original_stderr
        
    finally:
        # 恢复原始的stdout和stderr
        sys.stdout = original_stdout
        sys.stderr = original_stderr


def test_intercept_handler():
    """测试InterceptHandler类"""
    # 创建InterceptHandler实例
    handler = InterceptHandler()
    assert handler is not None
    
    # 测试emit方法（不应抛出异常）
    try:
        # 创建一个日志记录
        record = logging.makeLogRecord({
            'levelname': 'INFO',
            'msg': '测试日志消息',
            'exc_info': None
        })
        # 调用emit方法
        handler.emit(record)
        # 测试不同级别的日志
        record = logging.makeLogRecord({
            'levelname': 'ERROR',
            'msg': '测试错误日志',
            'exc_info': None
        })
        handler.emit(record)
    except Exception as e:
        assert False, f"InterceptHandler.emit() 抛出异常: {e}"


def test_setup_uvicorn_loguru():
    """测试setup_uvicorn_loguru函数"""
    # 测试setup_uvicorn_loguru函数（不应抛出异常）
    try:
        setup_uvicorn_loguru()
        # 验证环境变量是否被设置
        assert os.environ.get("UVICORN_NO_COLOR") == "1"
    except Exception as e:
        assert False, f"setup_uvicorn_loguru() 抛出异常: {e}"
