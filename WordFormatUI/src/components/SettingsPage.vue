<template>
  <div class="settings-page">
    <div class="settings-card">
      <h2 class="settings-title">后端连接设置</h2>
      <p class="settings-desc">配置后端 API 服务的地址和端口，修改后自动保存。</p>

      <div class="connection-bar" :class="connectionClass">
        <span class="connection-dot"></span>
        <span>{{ connectionText }}</span>
        <span class="connection-url" v-if="connectionOk">{{ backendBaseUrl }}</span>
      </div>

      <div class="form-group">
        <label class="form-label">后端 IP 地址</label>
        <input v-model="host" class="form-input" placeholder="127.0.0.1" @input="onFieldChange" />
      </div>

      <div class="form-group">
        <label class="form-label">后端端口</label>
        <input v-model.number="port" type="number" class="form-input" placeholder="8000" min="1" max="65535" @input="onFieldChange" />
      </div>

      <div class="form-actions">
        <button class="btn btn-green" @click="testConnection" :disabled="testing">{{ testing ? '测试中...' : '测试连接' }}</button>
        <button class="btn btn-ghost" @click="resetDefaults">恢复默认</button>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue';
import { backendSettings, backendBaseUrl, saveSettings, resetSettings, loadSettings } from '../utils/settings';

const host = ref(backendSettings.host);
const port = ref(backendSettings.port);
const testing = ref(false);
const connectionOk = ref(null);

onMounted(() => { loadSettings(); host.value = backendSettings.host; port.value = backendSettings.port; });

const connectionClass = computed(() => connectionOk.value === null ? 'bar-untested' : connectionOk.value ? 'bar-ok' : 'bar-fail');
const connectionText = computed(() => connectionOk.value === null ? '未测试连接' : connectionOk.value ? '连接成功' : '连接失败');

function onFieldChange() { backendSettings.host = host.value; backendSettings.port = port.value; saveSettings(); connectionOk.value = null; }

async function testConnection() {
  testing.value = true; connectionOk.value = null;
  try { const resp = await fetch(`${backendBaseUrl.value}/openapi.json`, { signal: AbortSignal.timeout(5000) }); connectionOk.value = resp.ok; }
  catch { connectionOk.value = false; }
  finally { testing.value = false; }
}

function resetDefaults() { resetSettings(); host.value = backendSettings.host; port.value = backendSettings.port; connectionOk.value = null; }
</script>

<style scoped>
.settings-page { max-width: 520px; margin: 0 auto; }
.settings-card { background: #1e293b; border: 1px solid #334155; border-radius: 12px; padding: 2rem; }
.settings-title { font-size: 1.15rem; font-weight: 600; color: #f1f5f9; margin-bottom: 0.35rem; font-family: 'Crimson Pro', serif; }
.settings-desc { font-size: 0.8125rem; color: #64748b; margin-bottom: 1.5rem; }
.connection-bar { display: flex; align-items: center; gap: 0.5rem; padding: 0.75rem 1rem; border-radius: 8px; font-size: 0.8125rem; margin-bottom: 1.5rem; }
.connection-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.bar-untested { background: #0f172a; border: 1px solid #334155; color: #64748b; } .bar-untested .connection-dot { background: #64748b; }
.bar-ok { background: #052e16; border: 1px solid #166534; color: #4ade80; } .bar-ok .connection-dot { background: #22c55e; }
.bar-fail { background: #450a0a; border: 1px solid #7f1d1d; color: #fca5a5; } .bar-fail .connection-dot { background: #ef4444; }
.connection-url { margin-left: auto; font-family: monospace; font-size: 0.75rem; opacity: 0.8; }
.form-group { margin-bottom: 1rem; }
.form-label { display: block; font-size: 0.8125rem; font-weight: 500; color: #94a3b8; margin-bottom: 0.375rem; }
.form-input { width: 100%; padding: 0.625rem 0.75rem; font-size: 0.875rem; border: 1px solid #475569; border-radius: 8px; outline: none; background: #0f172a; color: #e2e8f0; transition: border-color .2s; font-family: inherit; box-sizing: border-box; }
.form-input:focus { border-color: #22c55e; box-shadow: 0 0 0 3px rgba(34,197,94,.15); }
.form-actions { display: flex; gap: 0.75rem; margin-top: 1.5rem; }
.btn { padding: 0.5rem 1rem; font-size: 0.8125rem; font-weight: 500; border: none; border-radius: 6px; cursor: pointer; transition: all .2s; font-family: inherit; }
.btn:disabled { opacity: 0.4; cursor: not-allowed; }
.btn-green { background: #22c55e; color: #052e16; }
.btn-green:hover:not(:disabled) { background: #16a34a; }
.btn-ghost { background: transparent; color: #94a3b8; border: 1px solid #475569; }
.btn-ghost:hover { background: #1e293b; color: #e2e8f0; }
</style>
