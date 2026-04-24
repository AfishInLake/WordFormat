#! /usr/bin/env python
# @Time    : 2026/4/8
# @Author  : afish
# @File    : log_config.py
import sys
from pathlib import Path
from loguru import logger


def setup_logger():
    """配置日志"""
    # 确定基础目录
    if getattr(sys, "frozen", False):
        BASE_DIR = Path(sys.executable).parent
    else:
        BASE_DIR = Path(__file__).parent.parent.parent
    
    LOG_FILE = BASE_DIR / "api.log"
    
    # 移除默认输出
    logger.remove()
    
    # 添加文件输出
    # 注意：不使用 enqueue=True，因为在沙箱环境中 multiprocessing.SimpleQueue()
    # 会因缺少 /dev/shm 而失败（FileNotFoundError）
    logger.add(
        LOG_FILE,  # 日志文件
        rotation="500 MB",  # 按大小分割
        retention="7 days",  # 保留7天
        encoding="utf-8",
        backtrace=True,  # 显示完整堆栈
        diagnose=True,  # 显示变量信息
    )
    
    # 添加控制台输出（非 -w 模式）
    if not sys.stdout.closed:
        logger.add(
            sys.stdout,
            colorize=True,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
        )


def setup_uvicorn_loguru():
    """修复 Uvicorn 日志，使其使用 Loguru"""
    import logging
    import uvicorn
    
    # 禁用 Uvicorn 的默认日志
    for logger_name in ["uvicorn", "uvicorn.access", "uvicorn.error"]:
        uvicorn_logger = logging.getLogger(logger_name)
        uvicorn_logger.disabled = True
    
    # 配置 Uvicorn 使用 Loguru
    uvicorn.config.LOGGING_CONFIG = {
        "version": 1,
        "disable_existing_loggers": True,
        "formatters": {
            "default": {
                "()": "uvicorn.logging.DefaultFormatter",
                "fmt": "%(levelprefix)s %(message)s",
                "use_colors": None,
            },
            "access": {
                "()": "uvicorn.logging.AccessFormatter",
                "fmt": '%(levelprefix)s %(client_addr)s - "%(request_line)s" %(status_code)s',
            },
        },
        "handlers": {
            "default": {
                "formatter": "default",
                "class": "logging.NullHandler",
            },
            "access": {
                "formatter": "access",
                "class": "logging.NullHandler",
            },
        },
        "loggers": {
            "uvicorn": {"handlers": ["default"], "level": "INFO"},
            "uvicorn.error": {"handlers": ["default"], "level": "INFO"},
            "uvicorn.access": {"handlers": ["access"], "level": "INFO"},
        },
    }
