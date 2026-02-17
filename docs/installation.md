# 安装指南

本文档详细说明 WordFormat 项目的环境要求和安装步骤。

## 环境要求

- **Python 3.11+**（推荐 3.11 及以上版本）
- **依赖管理工具**：uv（推荐）或 pip

## 安装步骤

### 1. 克隆项目

```bash
git clone https://github.com/AfishInLake/WordFormat.git
cd WordFormat
```

### 2. 安装依赖

使用 uv（推荐）：

```bash
make install
```

或使用 pip：

```bash
pip install -e .
```

### 3. 配置环境变量

创建 `.env` 文件，配置 HOST、PORT 等必要参数：

```env
HOST="127.0.0.1"
# 配置服务端口
PORT="8000"
```

### 4. 启动API服务

```bash
# 在虚拟环境下运行
make server
```

### 5. 构建.exe程序

```bash
# 在虚拟环境下运行
make build
```

## 验证安装

安装完成后，可以通过以下命令验证安装是否成功：

```bash
wordformat --help
```

如果输出了命令行帮助信息，则说明安装成功。

## 故障排查

### 依赖安装失败

- 确保 Python 版本为 3.11+
- 尝试更新 pip 到最新版本：`pip install --upgrade pip`
- 检查网络连接，确保能够访问 PyPI

### 命令行工具不可用

- 检查是否在虚拟环境中安装
- 确保安装时使用了 `-e` 选项（editable mode）
- 检查系统 PATH 是否包含了 Python 的脚本目录

### API 服务启动失败

- 检查 `.env` 文件中的配置是否正确
- 确保端口未被其他服务占用
