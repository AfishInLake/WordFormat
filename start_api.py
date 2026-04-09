#! /usr/bin/env python
# @Time    : 2026/2/5 21:42
# @Author  : afish
# @File    : start_api.py
# start_api.py

# 在所有导入之前禁用 multiprocessing 的 resource_tracker
import os
os.environ['MULTIPROCESSING_RESOURCE_TRACKER'] = '0'

import sys

# 检查是否是 multiprocessing 的子进程（resource_tracker 或 forkserver）
# 如果是，跳过主程序代码，让 PyInstaller 的 runtime hook 处理
_IS_MULTIPROCESSING_CHILD = False
if len(sys.argv) >= 2 and sys.argv[-2] == '-c' and sys.argv[-1].startswith(
    ('from multiprocessing.resource_tracker import main', 'from multiprocessing.forkserver import main')
):
    _IS_MULTIPROCESSING_CHILD = True

if not _IS_MULTIPROCESSING_CHILD:
    from loguru import logger

    from wordformat.api import app
    from wordformat.log_config import setup_logger, setup_uvicorn_loguru
    from wordformat.settings import HOST, PORT

    # ========== 第一步：初始化 Loguru + 修复 Uvicorn 日志 ==========
    setup_logger()
    setup_uvicorn_loguru()


    def main():
        import socket
        import uvicorn
        import signal

        # 检查端口是否可用
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind((HOST, PORT))
            sock.close()
        except OSError as e:
            logger.error(f"❌ 端口 {PORT} 已被占用，无法启动服务！错误：{e}")
            logger.error(f"💡 请检查是否有其他进程占用端口：lsof -i :{PORT}")
            # 杀掉所有子进程（包括 resource_tracker）
            import subprocess
            try:
                subprocess.run(['pkill', '-9', '-P', str(os.getpid())], check=False)
            except Exception:
                pass
            # 使用 os._exit() 强制退出，避免 PyInstaller 自动重启
            os._exit(1)

        try:
            uvicorn.run(
                app,
                host=HOST,
                port=PORT,
                log_config=None,
                access_log=True,
                reload=False,
                use_colors=False,
            )
        except Exception as e:
            logger.error(f"❌ 服务启动失败：{e}")
            # 杀掉所有子进程（包括 resource_tracker）
            try:
                subprocess.run(['pkill', '-9', '-P', str(os.getpid())], check=False)
            except Exception:
                pass
            os._exit(1)


    if __name__ == "__main__":
        main()

    # PyInstaller 打包后的入口点
    if getattr(sys, 'frozen', False) and __name__ != "__main__":
        main()
