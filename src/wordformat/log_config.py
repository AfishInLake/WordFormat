#! /usr/bin/env python
# @Time    : 2026/4/8
# @Author  : afish
# @File    : log_config.py
import sys
from pathlib import Path

from loguru import logger


def setup_logger(log_dir: str | None = None):
    """配置日志。

    Args:
        log_dir: 日志文件目录。为 None 时仅输出到控制台（测试/开发模式）。
    """
    logger.remove()

    # 文件输出（仅指定目录时启用）
    if log_dir:
        log_path = Path(log_dir) / "api.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        logger.add(
            str(log_path),
            rotation="500 MB",
            retention="7 days",
            encoding="utf-8",
            level="DEBUG",
            backtrace=True,
            diagnose=True,
        )

    # 控制台输出：仅 INFO 及以上（DEBUG 不刷屏）
    if not sys.stdout.closed:
        logger.add(
            sys.stdout,
            colorize=True,
            level="INFO",
            format=(
                "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
                "<level>{level: <8}</level> | "
                "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - "
                "<level>{message}</level>"
            ),
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
