#! /usr/bin/env python
# @Time    : 2026/2/5 21:42
# @Author  : afish
# @File    : start_api.py
from src.settings import HOST, PORT

# ---------------------- 启动服务 ----------------------
if __name__ == "__main__":
    import uvicorn

    # 启动UVICORN服务器，监听所有IP，端口8000，自动重载
    uvicorn.run(app="src.api:app", host=HOST, port=PORT, reload=True, log_level="info")
