@echo off
chcp 65001 > nul
setlocal disabledelayedexpansion
:: ========== 核心：获取脚本所在目录（项目根目录），适配任意路径 ==========
set "PROJECT_ROOT=%~dp0"
:: 去除路径末尾的反斜杠（避免拼接错误）
if "%PROJECT_ROOT:~-1%"=="\" set "PROJECT_ROOT=%PROJECT_ROOT:~0,-1%"

echo ========================
echo 一键打包脚本（通用版-显示控制台）
echo ========================
echo 📌 项目根目录: %PROJECT_ROOT%
echo.

:: ========== 1. 动态定义所有路径（同步重命名为wordformat） ==========
set "VENV_DIR=%PROJECT_ROOT%\.venv"
set "ORIGIN_ORT_DIR=%VENV_DIR%\Lib\site-packages\onnxruntime"
set "TARGET_ORT_DIR=%PROJECT_ROOT%\dist\wordformat\_internal\onnxruntime"  # 改为wordformat
set "ORIGIN_MODEL_DIR=%PROJECT_ROOT%\model"
set "TARGET_MODEL_DIR=%PROJECT_ROOT%\dist\wordformat\model"  # 改为wordformat
set "EXE_PATH=%PROJECT_ROOT%\dist\wordformat\wordformat.exe"  # 最终exe名称

:: ========== 2. 激活虚拟环境 ==========
echo 🔄 激活虚拟环境...
set "ACTIVATE_SCRIPT=%VENV_DIR%\Scripts\activate.bat"
if not exist "%ACTIVATE_SCRIPT%" (
    echo ❌ 虚拟环境激活脚本不存在！路径：%ACTIVATE_SCRIPT%
    pause
    exit /b 1
)
call "%ACTIVATE_SCRIPT%"
if errorlevel 1 (
    echo ❌ 虚拟环境激活失败！
    pause
    exit /b 1
)
echo ✅ 虚拟环境激活完成！
echo.

:: ========== 3. 清理旧产物 ==========
echo 🔄 清理旧打包文件...
if exist "%PROJECT_ROOT%\dist" rmdir /s/q "%PROJECT_ROOT%\dist"
if exist "%PROJECT_ROOT%\build" rmdir /s/q "%PROJECT_ROOT%\build"
if exist "%PROJECT_ROOT%\wordformat.spec" del "%PROJECT_ROOT%\wordformat.spec"  # 清理新spec文件
echo ✅ 旧文件清理完成！
echo.

:: ========== 4. 安装/更新依赖 ==========
echo 🔄 安装打包依赖...
uv sync
if errorlevel 1 (
    echo ❌ 依赖安装失败！
    pause
    exit /b 1
)
echo ✅ 依赖安装完成！
echo.

:: ========== 5. 执行PyInstaller打包（取消-w + 重命名为wordformat） ==========
echo 🔄 开始打包程序...
cd /d "%PROJECT_ROOT%"
:: 核心修改：1. 移除-w（显示控制台） 2. -n wordformat（指定产物名称）
pyinstaller -D --noconfirm -n wordformat --add-data "src;src" --hidden-import=src.api --hidden-import=fastapi --hidden-import=uvicorn --hidden-import=pydantic --hidden-import=zoneinfo --exclude-module=onnxruntime start_api.py
if errorlevel 1 (
    echo ❌ 程序打包失败！
    pause
    exit /b 1
)
echo ✅ 程序打包完成！
echo.

:: ========== 6. 复制ONNX到指定目录（同步新名称路径） ==========
echo 🔄 复制onnxruntime到打包目录...
echo   源路径: %ORIGIN_ORT_DIR%
echo   目标路径: %TARGET_ORT_DIR%
:: 检查源ONNX目录是否存在
if not exist "%ORIGIN_ORT_DIR%" (
    echo ❌ 源ONNX目录不存在！路径：%ORIGIN_ORT_DIR%
    pause
    exit /b 1
)
:: 删除旧目录 + 复制
if exist "%TARGET_ORT_DIR%" rmdir /s/q "%TARGET_ORT_DIR%"
xcopy "%ORIGIN_ORT_DIR%" "%TARGET_ORT_DIR%\" /E /I /Y
if errorlevel 1 (
    echo ❌ onnxruntime复制失败！
    pause
    exit /b 1
)
echo ✅ onnxruntime复制完成！
echo.

:: ========== 7. 复制model文件夹到exe同级目录（同步新名称路径） ==========
echo 🔄 复制model文件夹到exe同级目录...
echo   源路径: %ORIGIN_MODEL_DIR%
echo   目标路径: %TARGET_MODEL_DIR%
:: 检查源model目录是否存在
if not exist "%ORIGIN_MODEL_DIR%" (
    echo ❌ 源model目录不存在！路径：%ORIGIN_MODEL_DIR%
    pause
    exit /b 1
)
:: 删除旧目录 + 复制
if exist "%TARGET_MODEL_DIR%" rmdir /s/q "%TARGET_MODEL_DIR%"
xcopy "%ORIGIN_MODEL_DIR%" "%TARGET_MODEL_DIR%\" /E /I /Y
if errorlevel 1 (
    echo ❌ model文件夹复制失败！
    pause
    exit /b 1
)
echo ✅ model文件夹复制完成！
echo.

:: ========== 8. 完成提示（同步新名称） ==========
echo ========================
echo ✅ 全部流程完成
echo 📁 产物路径: %EXE_PATH%
echo 📁 Model路径: %TARGET_MODEL_DIR%
echo ========================
pause