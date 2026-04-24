# 安装指南

本文档详细说明 WordFormat 项目的环境要求和安装步骤。

## 环境要求

- **Python 3.10+**（推荐 3.11 及以上版本）
- **依赖管理工具**：uv（推荐）或 pip

## 安装方式

### 方式一：从 PyPI 安装（推荐普通用户）

```bash
# 使用 pip
pip install wordformat

# 使用 uv
uv pip install wordformat
```

安装完成后即可使用 `wf` 和 `wordformat` 命令。

如需 API 服务功能，安装 api 可选依赖：

```bash
pip install wordformat[api]
```

### 方式二：从源码安装（推荐开发者）

1. **克隆项目**

```bash
git clone https://github.com/AfishInLake/WordFormat.git
cd WordFormat
```

2. **安装依赖**

使用 uv（推荐）：

```bash
uv pip install -e .
```

或使用 pip：

```bash
pip install -e .
```

如需完整开发环境（测试、构建、API 等）：

```bash
uv pip install -e ".[dev]"
```

## 验证安装

安装完成后，可以通过以下命令验证安装是否成功：

```bash
wordformat --help
# 或
wf --help
```

如果输出了命令行帮助信息，则说明安装成功。

## 可选依赖说明

| 依赖组 | 安装命令 | 包含内容 |
|--------|----------|----------|
| `api` | `pip install wordformat[api]` | FastAPI、uvicorn（API 服务） |
| `test` | `pip install wordformat[test]` | pytest（测试框架） |
| `dev` | `pip install wordformat[dev]` | 以上全部 + ruff、pre-commit、pyinstaller |

## 故障排查

### 依赖安装失败

- 确保 Python 版本为 3.10+
- 尝试更新 pip 到最新版本：`pip install --upgrade pip`
- 检查网络连接，确保能够访问 PyPI

### 命令行工具不可用

- 检查是否在虚拟环境中安装
- 确保安装时使用了 `-e` 选项（editable mode）
- 检查系统 PATH 是否包含了 Python 的脚本目录

### API 服务启动失败

- 确保已安装 api 可选依赖：`pip install wordformat[api]`
- 检查端口未被其他服务占用
