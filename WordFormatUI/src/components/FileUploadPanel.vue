<template>
  <div class="upload-group">
    <label for="docx-file" class="upload-btn">选择 docx</label>
    <input type="file" id="docx-file" accept=".docx" class="file-hidden" @change="onDocxChange" />
    <span class="file-tip">{{ docxTip }}</span>

    <template v-if="!generatedConfig">
      <label for="yaml-file" class="upload-btn">选择 yaml</label>
      <input type="file" id="yaml-file" accept=".yaml,.yml" class="file-hidden" @change="onYamlChange" />
      <span class="file-tip">{{ yamlTip }}</span>
    </template>

    <button class="go-btn" :disabled="!docxFile || (!yamlFile && !generatedConfig) || isLoading" @click="$emit('generate-json')">
      生成节点
    </button>
  </div>
</template>

<script setup>
import { computed } from 'vue'
const props = defineProps({ docxFile: File, yamlFile: File, generatedConfig: Object, isLoading: Boolean })
const emit = defineEmits(['update:docxFile', 'update:yamlFile', 'generate-json'])
const docxTip = computed(() => props.docxFile ? `已选择：${props.docxFile.name}` : '未选择')
const yamlTip = computed(() => props.yamlFile ? `已选择：${props.yamlFile.name}` : '未选择')
const onDocxChange = (e) => emit('update:docxFile', e.target.files[0])
const onYamlChange = (e) => emit('update:yamlFile', e.target.files[0])
</script>

<style scoped>
.upload-group { display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap; }
.upload-btn { padding: 6px 12px; font-size: 12px; font-weight: 500; border: 1px solid #475569; border-radius: 6px; background: #1e293b; color: #cbd5e1; cursor: pointer; transition: all .2s; font-family: inherit; }
.upload-btn:hover { background: #334155; border-color: #64748b; color: #e2e8f0; }
.file-hidden { display: none; }
.file-tip { font-size: 11px; color: #64748b; white-space: nowrap; }
.go-btn { padding: 6px 14px; font-size: 12px; font-weight: 600; border: none; border-radius: 6px; background: #22c55e; color: #052e16; cursor: pointer; transition: all .2s; font-family: inherit; }
.go-btn:hover:not(:disabled) { background: #16a34a; }
.go-btn:disabled { opacity: 0.4; cursor: not-allowed; }
</style>
