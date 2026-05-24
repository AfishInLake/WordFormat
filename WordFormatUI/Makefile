.PHONY: help dev build preview install clean

help:
	@echo "WordFormatUI Commands"
	@echo "   make dev       # 启动 Vite 开发服务器"
	@echo "   make build     # 构建生产静态文件"
	@echo "   make preview   # 预览构建产物"
	@echo "   make install   # 安装依赖"
	@echo "   make clean     # 清理构建产物"

dev:
	npm run dev

build:
	npm run build

preview:
	npm run preview

install:
	npm install

clean:
	@echo "Cleaning..."
	rm -rf dist node_modules/.vite
	@echo "Clean complete."
