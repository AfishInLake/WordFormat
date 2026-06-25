<template>
  <div class="config-generator">
    <div class="app-content">
      <div class="config-sections">
        <!-- 模板名称 -->
        <ConfigSection title="模板信息">
          <div class="form-item" style="margin-bottom:10px">
            <label>模板名称:</label>
            <input type="text" v-model="userConfig.template_name" style="width:300px">
          </div>
        </ConfigSection>

        <!-- 警告字段配置 -->
        <ConfigSection title="警告字段配置">
          <WarningFieldsConfig :config="userConfig.style_checks_warning" />
        </ConfigSection>

        <!-- 全局基础格式配置 -->
        <ConfigSection title="全局基础格式配置">
          <GlobalFormatConfig :config="userConfig.global_format" :show-apply-button="true"
            @apply-to-all="handleApplyGlobalFormat" />
        </ConfigSection>

        <!-- 摘要配置 -->
        <ConfigSection title="摘要配置">
          <AbstractConfig :config="userConfig.abstract" />
        </ConfigSection>

        <!-- 标题配置 -->
        <ConfigSection title="标题配置">
          <HeadingsConfig :config="userConfig.headings" />
        </ConfigSection>

        <!-- 正文配置 -->
        <ConfigSection title="正文配置">
          <div class="grid grid-cols-2 gap-4 mb-4">
            <div class="form-item">
              <label><input type="checkbox" v-model="userConfig.body_text.rules.punctuation.enabled"> 启用标点符号校验</label>
            </div>
          </div>
          <FormatConfig :config="userConfig.body_text" />
        </ConfigSection>

        <!-- 插图配置 -->
        <ConfigSection title="插图配置">
          <FiguresConfig :config="userConfig.figures" />
        </ConfigSection>

        <!-- 表格配置 -->
        <ConfigSection title="表格配置">
          <TablesConfig :config="userConfig.tables" />
        </ConfigSection>

        <!-- 参考文献配置 -->
        <ConfigSection title="参考文献配置">
          <ReferencesConfig :config="userConfig.references" />
        </ConfigSection>

        <!-- 致谢配置 -->
        <ConfigSection title="致谢配置">
          <AcknowledgementsConfig :config="userConfig.acknowledgements" />
        </ConfigSection>

        <!-- 标题自动编号 -->
        <ConfigSection title="标题自动编号">
          <NumberingConfig :config="userConfig.numbering" />
        </ConfigSection>
      </div>

      <!-- 生成的YAML结果 -->
      <div class="yaml-preview">
        <YamlOutput :yaml-content="generatedYaml" @reset-to-default="resetToDefault" @import-yaml="handleImportYaml" />
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import yaml from 'js-yaml'

// 导入工具函数和默认配置
import { defaultConfig, applyGlobalFormatToAll, mergeWithDefaults } from './utils'

// 导入组件
import ConfigSection from './components/ConfigSection.vue'
import WarningFieldsConfig from './components/WarningFieldsConfig.vue'
import GlobalFormatConfig from './components/GlobalFormatConfig.vue'
import FormatConfig from './components/FormatConfig.vue'
import AbstractConfig from './components/AbstractConfig.vue'
import HeadingsConfig from './components/HeadingsConfig.vue'
import FiguresConfig from './components/FiguresConfig.vue'
import TablesConfig from './components/TablesConfig.vue'
import ReferencesConfig from './components/ReferencesConfig.vue'
import AcknowledgementsConfig from './components/AcknowledgementsConfig.vue'
import NumberingConfig from './components/NumberingConfig.vue'
import YamlOutput from './components/YamlOutput.vue'

// 用户配置
const userConfig = ref(JSON.parse(JSON.stringify(defaultConfig)))

// 计算生成的YAML
const generatedYaml = computed(() => {
  return yaml.dump(userConfig.value, { indent: 2, skipInvalid: true })
})

// 重置为默认配置
const resetToDefault = () => {
  userConfig.value = JSON.parse(JSON.stringify(defaultConfig))
}

// 应用全局格式到所有局部配置
const handleApplyGlobalFormat = () => {
  applyGlobalFormatToAll(userConfig.value)
}

// 处理导入YAML配置
const handleImportYaml = (parsedConfig) => {
  userConfig.value = mergeWithDefaults(parsedConfig, defaultConfig)
}

// 导出当前配置
const exportConfig = () => {
  return JSON.parse(JSON.stringify(userConfig.value))
}

// 导入配置
const importConfig = (config) => {
  userConfig.value = mergeWithDefaults(JSON.parse(JSON.stringify(config)), defaultConfig)
}

// 定义事件
const emit = defineEmits(['config-updated']);

// 监听配置变化，自动通知父组件
import { watch, onMounted } from 'vue';
watch(
  userConfig,
  (newConfig) => {
    emit('config-updated', JSON.parse(JSON.stringify(newConfig)));
  },
  { deep: true }
);

// 初始化时通知父组件
onMounted(() => {
  emit('config-updated', JSON.parse(JSON.stringify(userConfig.value)));
});

// 导出方法
defineExpose({
  exportConfig,
  importConfig,
  resetToDefault,
  handleApplyGlobalFormat
})
</script>

<style scoped>
.config-generator { width: 100%; }
.app-content { display: flex; gap: 20px; width: 100%; }
.config-sections { flex: 1; min-width: 0; }
.yaml-preview { width: 380px; min-width: 260px; position: sticky; top: 72px; align-self: flex-start; max-height: calc(100vh - 100px); overflow-y: auto; }

.form-item { margin-bottom: 5px; display: flex; align-items: center; gap: 8px; }
.form-item label { font-size: 12px; font-weight: 500; color: var(--text-secondary); }
.form-item input[type="text"] { padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; background: var(--ink); color: var(--text); outline: none; flex: 1; max-width: 300px; }
.form-item input[type="text"]:focus { border-color: var(--brass); }

@media (max-width: 1200px) { .app-content { flex-direction: column; } .yaml-preview { width: 100%; position: static; max-height: none; } }
@media (max-width: 768px) { .app-content { gap: 10px; } }
</style>
