<template>
  <div class="config-sidebar">
    <div class="sidebar-header">
      <h3>配置模板</h3>
      <button class="btn-refresh" @click="fetchConfigs" title="刷新列表">↻</button>
    </div>
    <div class="config-list">
      <div
        v-for="cfg in configs"
        :key="cfg"
        class="config-item"
        :class="{ active: cfg === activeConfig }"
        @click="selectConfig(cfg)"
      >
        <span class="config-name">{{ cfg }}</span>
      </div>
      <div v-if="loading" class="status-text">加载中...</div>
      <div v-if="!loading && configs.length === 0" class="status-text">
        暂无配置文件<br/>将 YAML 放入 configs/ 目录
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'

const emit = defineEmits(['config-selected'])
const configs = ref([])
const loading = ref(false)
const activeConfig = ref('')

const API_BASE = window.__API_BASE__ || ''

async function fetchConfigs() {
  loading.value = true
  try {
    const res = await fetch(`${API_BASE}/configs`)
    const json = await res.json()
    if (json.code === 200) { configs.value = json.data || [] }
  } catch (e) { console.error('获取配置列表失败:', e) }
  finally { loading.value = false }
}

async function selectConfig(filename) {
  try {
    const res = await fetch(`${API_BASE}/configs/${encodeURIComponent(filename)}`)
    const json = await res.json()
    if (json.code === 200) {
      activeConfig.value = filename
      emit('config-selected', { filename, content: json.data })
    }
  } catch (e) { console.error('读取配置失败:', e) }
}

onMounted(fetchConfigs)
</script>

<style scoped>
.config-sidebar {
  width: 200px; min-width: 180px;
  background: var(--paper);
  border-right: 1px solid var(--border);
  padding: 1rem 0;
  height: 100%; display: flex; flex-direction: column;
}
.sidebar-header { display: flex; align-items: center; justify-content: space-between; padding: 0 1rem 0.5rem; }
.sidebar-header h3 { font-size: 0.8rem; color: var(--text-muted); font-weight: 600; letter-spacing: 0.03em; text-transform: uppercase; }
.btn-refresh { background: none; border: 1px solid var(--border); color: var(--text-muted); border-radius: 4px; cursor: pointer; font-size: 0.9rem; padding: 0 5px; line-height: 1.3; }
.btn-refresh:hover { color: var(--text); background: var(--surface); }
.config-list { flex: 1; overflow-y: auto; padding: 0 0.5rem; }
.config-item {
  padding: 0.35rem 0.7rem; cursor: pointer; border-radius: 5px;
  font-size: 0.78rem; color: var(--text-secondary); margin-bottom: 2px;
  transition: background .1s;
}
.config-item:hover { background: var(--surface); }
.config-item.active { background: var(--brass); color: var(--ink); font-weight: 600; }
.status-text { padding: 0.5rem; font-size: 0.7rem; color: var(--text-muted); text-align: center; line-height: 1.5; }
</style>
