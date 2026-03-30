#!/bin/bash

set -euo pipefail  # 遇错即停，未定义变量报错

# ========== 获取项目根目录 ==========
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(realpath "$SCRIPT_DIR/..")"

echo "========================"
echo "一键打包脚本（Linux Web API 版）"
echo "========================"
echo "📌 项目根目录: $PROJECT_ROOT"
echo

# ========== 路径定义 ==========
VENV_DIR="$PROJECT_ROOT/.venv"
DIST_EXE="$PROJECT_ROOT/dist/wordformat/wordformat"

# ========== 检查虚拟环境 ==========
if [ ! -f "$VENV_DIR/bin/activate" ]; then
    echo "❌ 虚拟环境不存在！请先运行：uv venv"
    exit 1
fi

echo "🔄 激活虚拟环境..."
source "$VENV_DIR/bin/activate"
echo "✅ 虚拟环境已激活！"
echo

# ========== 清理旧产物 ==========
echo "🔄 清理旧产物..."
rm -rf "$PROJECT_ROOT/dist" "$PROJECT_ROOT/build"
rm -f "$PROJECT_ROOT/wordformat.spec"
echo "✅ 清理完成！"
echo

# ========== 安装依赖 ==========
echo "🔄 同步依赖..."
uv sync || { echo "❌ 依赖安装失败！"; exit 1; }
echo "✅ 依赖同步完成！"
echo

# ========== 打包 ==========
echo "🔄 开始打包..."
cd "$PROJECT_ROOT"

pyinstaller -D --noconfirm -n wordformat \
  --paths "src" \
  --add-data "src/wordformat/data:wordformat/data" \
  --hidden-import=wordformat.api \
  --hidden-import=wordformat.data \
  --hidden-import=fastapi \
  --hidden-import=uvicorn \
  --hidden-import=pydantic \
  --hidden-import=zoneinfo \
  --hidden-import=tokenizers \
  start_api.py

if [ $? -ne 0 ]; then
    echo "❌ 打包失败！"
    exit 1
fi
echo "✅ 打包完成！"
echo

# ========== （可选）验证是否包含 wordformat.data ==========
echo "🔍 验证打包结果是否包含 wordformat.data..."
if [ ! -d "$PROJECT_ROOT/dist/wordformat/_internal/wordformat/data" ]; then
    echo "⚠️  警告：wordformat/data 未被打包进 _internal！"
    echo "   请检查 src/wordformat/data 目录是否存在 __init__.py"
else
    echo "✅ wordformat.data 已正确包含。"
fi
echo

# ========== 完成 ==========
echo "========================"
echo "✅ 打包成功！"
echo "🚀 运行命令: $DIST_EXE"
echo "========================"
