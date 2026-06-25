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
        <button class="btn btn-primary" @click="testConnection" :disabled="testing">{{ testing ? '测试中...' : '测试连接' }}</button>
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
.settings-card { background: var(--paper); border: 1px solid var(--border); border-radius: 10px; padding: 2rem; }
.settings-title { font-size: 1.05rem; font-weight: 600; color: var(--text); margin-bottom: 0.3rem; }
.settings-desc { font-size: 0.8rem; color: var(--text-muted); margin-bottom: 1.25rem; }
.connection-bar { display: flex; align-items: center; gap: 0.5rem; padding: 0.65rem 0.9rem; border-radius: 7px; font-size: 0.78rem; margin-bottom: 1.25rem; }
.connection-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
.bar-untested { background: var(--surface); border: 1px solid var(--border); color: var(--text-muted); } .bar-untested .connection-dot { background: var(--text-muted); }
.bar-ok { background: var(--green-dim); border: 1px solid var(--green); color: #6ee7b7; } .bar-ok .connection-dot { background: var(--green); }
.bar-fail { background: var(--red-dim); border: 1px solid var(--red); color: #fca5a5; } .bar-fail .connection-dot { background: var(--red); }
.connection-url { margin-left: auto; font-family: 'SF Mono', Menlo, Monaco, monospace; font-size: 0.72rem; opacity: 0.7; }
.form-group { margin-bottom: 0.9rem; }
.form-label { display: block; font-size: 0.78rem; font-weight: 500; color: var(--text-secondary); margin-bottom: 0.3rem; }
.form-input { width: 100%; padding: 0.55rem 0.7rem; font-size: 0.85rem; border: 1px solid var(--border); border-radius: 7px; outline: none; background: var(--ink); color: var(--text); transition: border-color .12s; font-family: inherit; box-sizing: border-box; }
.form-input:focus { border-color: var(--brass); box-shadow: 0 0 0 3px rgba(184,153,59,.08); }
.form-actions { display: flex; gap: 0.65rem; margin-top: 1.25rem; }
.btn { padding: 0.45rem 0.95rem; font-size: 0.78rem; font-weight: 500; border: none; border-radius: 6px; cursor: pointer; transition: all .12s; font-family: inherit; }
.btn:disabled { opacity: 0.35; cursor: not-allowed; }
.btn-primary { background: var(--brass); color: var(--ink); font-weight: 600; }
.btn-primary:hover:not(:disabled) { background: #c9a94d; }
.btn-ghost { background: transparent; color: var(--text-muted); border: 1px solid var(--border); }
.btn-ghost:hover { background: var(--surface); color: var(--text); }
</style>
