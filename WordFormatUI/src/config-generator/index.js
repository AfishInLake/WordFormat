// 配置生成器组件包入口

// 导入主要组件
export { default as ConfigGenerator } from './ConfigGenerator.vue'

// 导入子组件
export { default as ConfigSection } from './components/ConfigSection.vue'
export { default as WarningFieldsConfig } from './components/WarningFieldsConfig.vue'
export { default as GlobalFormatConfig } from './components/GlobalFormatConfig.vue'
export { default as FormatConfig } from './components/FormatConfig.vue'
export { default as AbstractConfig } from './components/AbstractConfig.vue'
export { default as HeadingsConfig } from './components/HeadingsConfig.vue'
export { default as FiguresConfig } from './components/FiguresConfig.vue'
export { default as TablesConfig } from './components/TablesConfig.vue'
export { default as ReferencesConfig } from './components/ReferencesConfig.vue'
export { default as AcknowledgementsConfig } from './components/AcknowledgementsConfig.vue'
export { default as NumberingConfig } from './components/NumberingConfig.vue'
export { default as YamlOutput } from './components/YamlOutput.vue'

// 导入工具函数和默认配置
export * from './utils'
