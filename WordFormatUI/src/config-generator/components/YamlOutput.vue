<template>
  <div class="yaml-output">
    <div class="yaml-header">生成的 YAML 配置</div>
    <div class="yaml-content"><pre>{{ yamlContent }}</pre></div>
    <div class="yaml-actions">
      <button @click="copyToClipboard" class="btn btn-green">复制</button>
      <button @click="importFromFile" class="btn btn-ghost">导入 YAML</button>
      <button @click="$emit('reset-to-default')" class="btn btn-ghost">重置</button>
    </div>
    <input type="file" ref="fileInput" style="display:none" accept=".yaml,.yml" @change="handleFileImport" />
  </div>
</template>

<script setup>
import { ref } from 'vue'
import yaml from 'js-yaml'
const props = defineProps({ yamlContent: { type: String, required: true } })
const emit = defineEmits(['reset-to-default', 'import-yaml'])
const fileInput = ref(null)
const copyToClipboard = () => { navigator.clipboard.writeText(props.yamlContent) }
const importFromFile = () => { fileInput.value.click() }
const handleFileImport = (event) => {
  const file = event.target.files[0]; if (!file) return
  const reader = new FileReader()
  reader.onload = (e) => { try { emit('import-yaml', yaml.load(e.target.result)) } catch (err) { alert('导入失败: ' + err.message) } }
  reader.onerror = () => { alert('文件读取失败') }
  reader.readAsText(file, 'utf-8'); event.target.value = ''
}
</script>

<style scoped>
.yaml-output { border: 1px solid #334155; border-radius: 10px; overflow: hidden; }
.yaml-header { padding: 12px 18px; background: #0f172a; font-size: 14px; font-weight: 600; color: #e2e8f0; }
.yaml-content { padding: 16px 18px; background: #020617; max-height: 360px; overflow-y: auto; }
.yaml-content pre { margin: 0; font-family: 'SFMono-Regular', Menlo, Monaco, Consolas, monospace; font-size: 12px; line-height: 1.6; color: #94a3b8; white-space: pre; word-wrap: break-word; tab-size: 2; }
.yaml-actions { padding: 10px 18px; background: #0f172a; border-top: 1px solid #334155; display: flex; gap: 8px; }
.btn { padding: 6px 14px; border: none; border-radius: 6px; font-size: 12px; cursor: pointer; font-family: inherit; transition: all .2s; }
.btn-green { background: #22c55e; color: #052e16; }
.btn-green:hover { background: #16a34a; }
.btn-ghost { background: transparent; color: #94a3b8; border: 1px solid #475569; }
.btn-ghost:hover { background: #1e293b; color: #e2e8f0; }
</style>
