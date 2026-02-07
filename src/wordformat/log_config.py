#! /usr/bin/env python
# @Time    : 2026/2/6 17:17
# @Author  : afish
# @File    : log_config.py
# log_config.py
import logging  # 新增：导入原生 logging
import os
import sys

from loguru import logger


# 修复 PyInstaller -w 模式下 sys.stdout/sys.stderr 为空的问题
def fix_std_streams():
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")


# 修正：继承 logging.Handler，补全 level 属性
class InterceptHandler(logging.Handler):
    def __init__(self, level=logging.NOTSET):  # 显式初始化 level
        super().__init__(level)

    def emit(self, record):
        # 获取 Loguru 对应的日志级别
        try:
            level = logger.level(record.levelname).name
        except ValueError:
            level = record.levelno

        # 找到调用者信息
        frame, depth = sys._getframe(6), 6
        while frame and frame.f_code.co_filename == logging.__file__:
            frame = frame.f_back
            depth += 1

        # 输出到 Loguru
        logger.opt(depth=depth, exception=record.exc_info).log(
            level, record.getMessage()
        )


# 彻底替换 Uvicorn 的日志配置
def setup_uvicorn_loguru():
    # 修复标准输出
    fix_std_streams()

    # 移除 Uvicorn 所有原生 handler
    for name in ["uvicorn", "uvicorn.error", "uvicorn.access"]:
        uvicorn_logger = logging.getLogger(name)
        uvicorn_logger.handlers = []
        uvicorn_logger.propagate = False
        uvicorn_logger.setLevel(logging.INFO)  # 显式设置级别

    # 给 Uvicorn 绑定 Loguru 的 InterceptHandler
    intercept_handler = InterceptHandler()
    logging.getLogger("uvicorn").addHandler(intercept_handler)
    logging.getLogger("uvicorn.error").addHandler(intercept_handler)
    logging.getLogger("uvicorn.access").addHandler(intercept_handler)

    # 禁用 Uvicorn 颜色输出
    os.environ["UVICORN_NO_COLOR"] = "1"
