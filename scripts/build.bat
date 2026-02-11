@echo off
chcp 65001 > nul
setlocal disabledelayedexpansion

:: ========== 从参数获取项目根目录 ==========
if "%~1"=="" (
    echo ❌ 错误：未指定项目根目录。
    echo 请通过 'make build' 调用此脚本。
    pause & exit /b 1
)
set "PROJECT_ROOT=%~1"

echo ========================
echo 一键打包脚本（Web API 版）
echo ========================
echo 📌 项目根目录: %PROJECT_ROOT%
echo.

:: ========== 路径定义 ==========
set "VENV_DIR=%PROJECT_ROOT%/.venv"
set "EXE_PATH=%PROJECT_ROOT%/dist/wordformat/wordformat.exe"

:: ========== 激活虚拟环境 ==========
echo 🔄 激活虚拟环境...
set "ACTIVATE_SCRIPT=%VENV_DIR%/Scripts/activate.bat"
if not exist "%ACTIVATE_SCRIPT%" (
    echo ❌ 虚拟环境不存在！请先运行 uv venv
    pause & exit /b 1
)
call "%ACTIVATE_SCRIPT%"
echo ✅ 虚拟环境激活完成！
echo.

:: ========== 清理 ==========
echo 🔄 清理旧产物...
rd /s/q "%PROJECT_ROOT%/dist" 2>nul
rd /s/q "%PROJECT_ROOT%/build" 2>nul
del "%PROJECT_ROOT%/wordformat.spec" 2>nul
echo ✅ 清理完成！
echo.

:: ========== 安装依赖 ==========
echo 🔄 同步依赖...
uv sync || (echo ❌ 依赖安装失败！ & pause & exit /b 1)
echo ✅ 依赖同步完成！
echo.

:: ========== 打包 ==========
echo 🔄 开始打包...
cd /d "%PROJECT_ROOT%"

pyinstaller -D --noconfirm -n wordformat ^
  --paths "src" ^
  --add-data "src/wordformat/data;wordformat/data" ^
  --hidden-import=wordformat.api ^
  --hidden-import=fastapi ^
  --hidden-import=uvicorn ^
  --hidden-import=pydantic ^
  --hidden-import=zoneinfo ^
  --hidden-import=tokenizers ^
  --exclude-module onnxruntime ^
  start_api.py

if errorlevel 1 (
    echo ❌ 打包失败！
    pause & exit /b 1
)
echo ✅ 打包完成
echo.

:: ========== 复制 ONNX Runtime DLL ==========
echo 🔄 复制 ONNX Runtime DLL...
set "ORT_CAPI_SRC=%VENV_DIR%/Lib/site-packages/onnxruntime"
set "ORT_CAPI_DST=%PROJECT_ROOT%/dist/wordformat/_internal/onnxruntime"

if not exist "%ORT_CAPI_SRC%" (
    echo ❌ ONNX Runtime  目录不存在！请确认已安装 onnxruntime
    pause & exit /b 1
)

xcopy "%ORT_CAPI_SRC%" "%ORT_CAPI_DST%/" /E /I /Y >nul
echo ✅ ONNX Runtime DLL 已复制到: %ORT_CAPI_DST%
echo.

:: ========== 完成 ==========
echo ========================
echo ✅ 打包成功
echo 🚀 运行命令: %EXE_PATH%
echo ========================
pause
