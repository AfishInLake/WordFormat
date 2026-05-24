# WordFormatUI

![License](https://img.shields.io/github/license/AfishInLake/WordFormatUI?color=blue)
![Electron](https://img.shields.io/badge/electron-33-blue)
![Status](https://img.shields.io/badge/status-开发中-orange)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey)

---

一个基于 **Vue 3 + Electron + Python** 构建的跨平台桌面应用，为 [WordFormat](https://github.com/AfishInLake/WordFormat) 提供图形化操作界面，帮助用户轻松完成学术论文格式自动化检查与修正。

> 💡 **技术栈：Electron + Vue 3**
> 使用 Electron 构建跨平台桌面应用，Python 后端以嵌入式运行时方式随应用分发，用户无需单独安装 Python。

---

## 🔗 原始项目

- 后端核心：[WordFormat](https://github.com/AfishInLake/WordFormat)（Python + bert 驱动的 Word 格式处理引擎）

---

## 📥 下载安装（普通用户）

✅ **无需编译，直接使用！**
前往 [Releases 页面](https://github.com/AfishInLake/WordFormatUI/releases) 下载对应操作系统的安装包：

- **Windows**：`.exe` 安装程序（NSIS）或 `.zip` 便携版
- **macOS**：`.zip` 应用包

---

## 👨‍💻 开发者指南

### 📦 先决条件

请确保已安装以下工具：

| 平台 | 要求 |
|------|------|
| **通用** | Node.js ≥ v18、npm（随 Node.js 自带） |
| **macOS** | Xcode 命令行工具（终端运行 `xcode-select --install`） |
| **Windows** | 无额外要求 |

---

### 🚀 快速启动

**方式一：浏览器模式（推荐，内存占用更低）**

```bash
# 1. 安装前端依赖
npm install

# 2. 构建 Python 后端运行时
make backend

# 3. 浏览器生产模式（构建 + 启动后端 + 打开浏览器）
make browser
```

**方式二：Electron 桌面模式**

```bash
# 1. 安装前端依赖
npm install

# 2. 构建 Python 后端运行时
make backend

# 3. 启动 Electron 开发模式
make dev
```

> ✅ `make backend` 会自动创建 Python 虚拟环境并安装 `wordformat[api]` 到 `electron/python-runtime/`。**此步骤不可跳过！**
>
> 💡 **浏览器模式 vs 桌面模式**：浏览器模式下 Python 后端独立运行、前端在系统浏览器中打开，省去 Electron 的 Chromium 开销（降低约 250MB 内存）。桌面模式下一切内嵌在 Electron 窗口中。两种模式功能完全一致。

---

### 🏗️ 构建生产版本

```bash
npm run electron:build
```

输出路径：
```
release/
```

- **Windows** → `WordFormat Setup x.x.x.exe` + `WordFormat-x.x.x-win.zip`
- **macOS** → `WordFormat-x.x.x-mac.zip`

---

## 🛠️ 常用命令

| 命令 | 说明 |
|------|------|
| `make browser` | 浏览器生产模式（构建 + 启动后端 + 打开浏览器） |
| `make browser-dev` | 浏览器开发模式（Vite 热重载 + 后端） |
| `make dev` | Electron 桌面开发模式 |
| `make backend` | 构建 Python 后端运行时 |
| `make build` | 构建并打包 Electron 分发版 |
| `make install` | 安装 npm 依赖 |
| `make clean` | 清理构建产物 |
| `make help` | 查看所有可用命令 |

---

## 📂 项目结构

```
WordFormatUI/
├── src/                     # Vue 前端源码
│   ├── components/          # Vue 组件
│   ├── composables/         # 组合式函数
│   ├── config-generator/    # 配置生成器
│   ├── electron/            # Electron renderer 端 API 封装
│   │   ├── backend.js       # 后端启动/停止
│   │   ├── dialog.js        # 原生对话框
│   │   ├── download.js      # 文件下载
│   │   └── fs.js            # 文件读写
│   └── utils/               # 工具函数
├── electron/                # Electron 主进程
│   ├── main.cjs             # 主进程入口
│   ├── preload.cjs          # preload 脚本
│   └── python-runtime/      # ← 自动生成，嵌入式 Python 运行时
├── scripts/                 # 构建脚本
│   ├── build-backend.cjs    # Python 后端构建脚本
│   └── start-browser.cjs    # 浏览器模式启动脚本
├── Makefile                 # 统一命令入口
└── package.json             # 前端依赖与打包配置
```

---

## 💡 注意事项

- **双模式运行**：支持浏览器模式和 Electron 桌面模式，功能完全一致。浏览器模式内存占用更低（省去 Chromium ~250MB）。
- **嵌入式运行时**：Python 后端随应用分发，用户无需安装 Python 环境。
- **平台绑定**：Electron 构建产物仅适用于当前操作系统（Windows 构建不能在 macOS 运行）；浏览器模式跨平台通用。
- **后端更新**：若 WordFormat 发布新版本，请更新 `.backend-version` 文件后重新运行 `make backend` 获取最新版本。
- **CI/CD**：项目已配置 GitHub Actions，支持 PR 构建验证与 Tag 触发的多平台自动发布。

---

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE)，欢迎自由使用、修改与分发。

---

> 🌟 **欢迎贡献！**
> 如发现 Bug 或有功能建议，请提交 [Issue](https://github.com/AfishInLake/WordFormatUI/issues) 或 Pull Request。
