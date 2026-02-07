#! /usr/bin/env python
# @Time    : 2026/2/5 21:42
# @Author  : afish
# @File    : start_api.py
# start_api.py
import sys

from loguru import logger

from wordformat.api import app
from wordformat.log_config import setup_uvicorn_loguru
from wordformat.settings import BASE_DIR, HOST, PORT

# ========== 第一步：初始化 Loguru + 修复 Uvicorn 日志 ==========
setup_uvicorn_loguru()
LOG_FILE = BASE_DIR / "api.log"
# 自定义 Loguru 输出（可选，根据需求配置）
logger.remove()  # 移除默认输出
logger.add(
    LOG_FILE,  # 日志文件
    rotation="500 MB",  # 按大小分割
    retention="7 days",  # 保留7天
    encoding="utf-8",
    enqueue=True,  # 异步写入
    backtrace=True,  # 显示完整堆栈
    diagnose=True,  # 显示变量信息
)
# 如果需要控制台输出（非 -w 模式），再添加控制台 handler
if not sys.stdout.closed:
    logger.add(
        sys.stdout,
        colorize=True,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",  # noqa E501
    )


def main():
    import uvicorn

    uvicorn.run(
        app,
        host=HOST,
        port=PORT,
        log_config=None,
        access_log=True,
        reload=False,
        use_colors=False,
    )


if __name__ == "__main__":
    main()
