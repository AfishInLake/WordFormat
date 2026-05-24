# WordFormatUI

![License](https://img.shields.io/github/license/AfishInLake/WordFormat?color=blue)
![Vue](https://img.shields.io/badge/vue-3-blue)
![Status](https://img.shields.io/badge/status-开发中-orange)

---

为 [WordFormat](https://github.com/AfishInLake/WordFormat) 提供 Web 图形化操作界面的 Vue 3 单页应用，帮助用户轻松完成学术论文格式自动化检查与修正。

> 前端构建产物随 Python 包分发，用户通过 `wf startapi` 启动后端后，在浏览器中直接使用。

---

## 开发者指南

### 先决条件

- Node.js ≥ v18
- npm（随 Node.js 自带）

### 快速启动（开发模式）

```bash
# 1. 安装依赖
npm install

# 2. 启动开发服务器（默认 http://localhost:5173）
npm run dev
```

开发模式下 Vite 提供热重载，API 请求会自动代理到本地 WordFormat 后端（需先启动 `wf startapi`）。

### 构建生产版本

```bash
npm run build
```

构建产物输出到 `dist/`，CI 在打 tag 时会自动构建并随 Python 包分发。

---

## 项目结构

```
WordFormatUI/
├── src/                     # Vue 前端源码
│   ├── components/          # Vue 组件（文档标签检查、文件上传、格式化操作等）
│   ├── composables/         # 组合式函数
│   ├── config-generator/    # YAML 配置可视化生成器
│   └── utils/               # 工具函数（请求封装、设置管理）
├── public/                  # 静态资源
├── index.html               # HTML 入口
├── vite.config.js           # Vite 配置（含 API 代理）
├── Makefile                 # 常用命令入口
└── package.json             # 依赖与脚本
```

---

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动 Vite 开发服务器（热重载） |
| `npm run build` | 构建生产版本到 `dist/` |
| `npm run preview` | 预览生产构建 |
| `make install` | 安装 npm 依赖 |

---

## 后端关联

前端构建产物在 CI 中自动复制到 `src/wordformat/api/static/`，随 `wordformat` Python 包一起发布。用户安装 `wordformat` 后运行 `wf startapi` 即可直接访问前端界面。

---

## 许可证

本项目采用 [MIT 许可证](LICENSE)，欢迎自由使用、修改与分发。

---

> 如发现 Bug 或有功能建议，请提交 [Issue](https://github.com/AfishInLake/WordFormat/issues) 或 Pull Request。
