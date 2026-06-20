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
      <div v-if="loading" class="loading-text">加载中...</div>
      <div v-if="!loading && configs.length === 0" class="empty-text">
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
    if (json.code === 200) {
      configs.value = json.data || []
    }
  } catch (e) {
    console.error('获取配置列表失败:', e)
  } finally {
    loading.value = false
  }
}

async function selectConfig(filename) {
  try {
    const res = await fetch(`${API_BASE}/configs/${encodeURIComponent(filename)}`)
    const json = await res.json()
    if (json.code === 200) {
      activeConfig.value = filename
      emit('config-selected', { filename, content: json.data })
    }
  } catch (e) {
    console.error('读取配置失败:', e)
  }
}

onMounted(fetchConfigs)
</script>

<style scoped>
.config-sidebar {
  width: 200px;
  min-width: 180px;
  background: #1e293b;
  border-right: 1px solid #334155;
  padding: 1rem 0;
  height: 100%;
  display: flex;
  flex-direction: column;
}
.sidebar-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 1rem 0.5rem;
}
.sidebar-header h3 {
  font-size: 0.85rem;
  color: #94a3b8;
  font-weight: 600;
}
.btn-refresh {
  background: none;
  border: 1px solid #475569;
  color: #94a3b8;
  border-radius: 4px;
  cursor: pointer;
  font-size: 1rem;
  padding: 0 6px;
}
.btn-refresh:hover { color: #e2e8f0; background: #334155; }
.config-list { flex: 1; overflow-y: auto; padding: 0 0.5rem; }
.config-item {
  padding: 0.4rem 0.75rem;
  cursor: pointer;
  border-radius: 4px;
  font-size: 0.8rem;
  color: #cbd5e1;
  margin-bottom: 2px;
  transition: background 0.15s;
}
.config-item:hover { background: #334155; }
.config-item.active { background: #22c55e; color: #052e16; font-weight: 600; }
.loading-text, .empty-text {
  padding: 0.5rem; font-size: 0.75rem; color: #64748b; text-align: center;
}
</style>
