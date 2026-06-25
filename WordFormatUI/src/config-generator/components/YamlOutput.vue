<template>
  <div class="yaml-output">
    <div class="yaml-header">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      配置预览
    </div>
    <div class="yaml-content"><pre>{{ yamlContent }}</pre></div>
    <div class="yaml-actions">
      <button @click="copyToClipboard" class="yaml-btn yaml-btn-primary">复制</button>
      <button @click="importFromFile" class="yaml-btn">导入</button>
      <button @click="$emit('reset-to-default')" class="yaml-btn">重置</button>
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
.yaml-output {
  border: 1px solid var(--border); border-radius: 10px; overflow: hidden;
  background: var(--paper);
}
.yaml-header {
  padding: 12px 16px; background: var(--surface);
  font-size: 12px; font-weight: 600; color: var(--text-secondary);
  display: flex; align-items: center; gap: 8px;
  border-bottom: 1px solid var(--border);
}
.yaml-header svg { color: var(--brass-dim); }
.yaml-content {
  padding: 14px 16px; background: var(--ink);
  max-height: 420px; overflow-y: auto;
}
.yaml-content pre {
  margin: 0; font-family: 'SF Mono', 'Fira Code', 'Fira Mono', Menlo, Monaco, Consolas, monospace;
  font-size: 11.5px; line-height: 1.7; color: var(--text-secondary); white-space: pre; tab-size: 2;
}
.yaml-actions {
  padding: 10px 16px; border-top: 1px solid var(--border);
  display: flex; gap: 8px;
}
.yaml-btn {
  padding: 5px 13px; border: 1px solid var(--border); border-radius: 6px;
  font-size: 11.5px; font-weight: 500; cursor: pointer; font-family: inherit;
  background: transparent; color: var(--text-muted); transition: all .12s;
}
.yaml-btn:hover { background: var(--surface); color: var(--text); border-color: var(--border-hover); }
.yaml-btn-primary { background: var(--brass); border-color: var(--brass); color: var(--ink); font-weight: 600; }
.yaml-btn-primary:hover { background: #c9a94d; border-color: #c9a94d; }
</style>
