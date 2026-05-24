# 配置生成器组件包

这是一个基于Vue 3的配置生成器组件包，用于生成和管理YAML配置文件。

## 功能特点

- 支持多个配置部分的管理（警告字段、全局格式、摘要、标题、正文、插图、表格、参考文献、致谢）
- 实时生成YAML配置
- 支持导入/导出YAML配置文件
- 响应式布局，适配不同屏幕尺寸
- 支持全局格式应用到所有局部配置

## 安装依赖

```bash
# 安装js-yaml依赖
npm install js-yaml
```

## 使用方法

### 1. 导入整个组件包

```vue
<template>
  <div>
    <h1>配置管理</h1>
    <ConfigGenerator ref="configGenerator" />
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { ConfigGenerator } from './config-generator'

const configGenerator = ref(null)

// 导出当前配置
const exportConfig = () => {
  const config = configGenerator.value.exportConfig()
  console.log('导出的配置:', config)
}

// 导入配置
const importConfig = (config) => {
  configGenerator.value.importConfig(config)
}

// 重置为默认配置
const resetConfig = () => {
  configGenerator.value.resetToDefault()
}

// 应用全局格式到所有配置
const applyGlobalFormat = () => {
  configGenerator.value.handleApplyGlobalFormat()
}
</script>
```

### 2. 单独使用子组件

```vue
<template>
  <div>
    <h2>全局格式配置</h2>
    <GlobalFormatConfig :config="globalFormat" :show-apply-button="true" @apply-to-all="handleApplyGlobalFormat" />
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { GlobalFormatConfig, defaultGlobalFormat } from './config-generator'

const globalFormat = ref({ ...defaultGlobalFormat })

const handleApplyGlobalFormat = () => {
  // 处理应用全局格式的逻辑
}
</script>
```

## 组件结构

- `ConfigGenerator.vue`: 主组件，包含所有配置部分
- `components/`:
  - `ConfigSection.vue`: 可折叠的配置节容器
  - `WarningFieldsConfig.vue`: 警告字段配置
  - `GlobalFormatConfig.vue`: 全局格式配置
  - `FormatConfig.vue`: 格式配置（用于各个部分的格式设置）
  - `AbstractConfig.vue`: 摘要配置
  - `HeadingsConfig.vue`: 标题配置
  - `FiguresConfig.vue`: 插图配置
  - `TablesConfig.vue`: 表格配置
  - `ReferencesConfig.vue`: 参考文献配置
  - `AcknowledgementsConfig.vue`: 致谢配置
  - `YamlOutput.vue`: YAML输出和导入/导出功能
- `utils.js`: 工具函数和默认配置

## 配置结构

配置对象的结构如下：

```javascript
{
  style_checks_warning: { /* 警告字段配置 */ },
  global_format: { /* 全局格式配置 */ },
  abstract: { /* 摘要配置 */ },
  headings: { /* 标题配置 */ },
  body_text: { /* 正文配置 */ },
  figures: { /* 插图配置 */ },
  tables: { /* 表格配置 */ },
  references: { /* 参考文献配置 */ },
  acknowledgements: { /* 致谢配置 */ }
}
```

## 自定义样式

组件使用了scoped样式，不会影响全局样式。如果需要自定义样式，可以通过CSS变量或覆盖样式来实现。

## 浏览器兼容性

- 支持所有现代浏览器（Chrome, Firefox, Safari, Edge）
- 不支持IE11

## 开发和构建

```bash
# 开发模式
npm run dev

# 构建
npm run build
```
