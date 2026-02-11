@echo off
chcp 65001 > nul
REM clean.bat - Clean build artifacts for WordFormat project

echo 🔄 开始清理构建产物

REM 删除文件夹（如果存在），不提示错误
if exist dist (
    rd /s /q dist
    echo   - Removed dist/
)

if exist build (
    rd /s /q build
    echo   - Removed build/
)

if exist output (
    rd /s /q output
    echo   - Removed output/
)

REM 删除 .spec 文件
dir /b *.spec >nul 2>&1
if not errorlevel 1 (
    del /q *.spec >nul
    echo   - Removed *.spec files
)

echo  ✅ 清理完成