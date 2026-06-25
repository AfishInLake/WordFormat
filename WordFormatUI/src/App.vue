<template>
  <div class="app-container">
    <GlobalToast ref="toastRef" />

    <div class="nav-bar">
      <div class="nav-content">
        <h1 class="app-title">WordFormat</h1>
        <div class="nav-actions">
          <div class="config-actions" v-if="activeTab === 'config'">
            <button class="btn btn-ghost" @click="saveConfig" :disabled="!generatedConfig">保存配置</button>
            <button class="btn btn-ghost" @click="loadConfig">加载配置</button>
            <button class="btn btn-ghost" @click="saveConfigToServer" :disabled="!generatedConfig">保存到服务器</button>
          </div>
          <div class="nav-tabs">
            <button class="nav-tab" :class="{ active: activeTab === 'config' }" @click="activeTab = 'config'">配置生成器</button>
            <button class="nav-tab" :class="{ active: activeTab === 'checker' }" @click="activeTab = 'checker'">文档标签核对</button>
            <button class="nav-tab" :class="{ active: activeTab === 'settings' }" @click="activeTab = 'settings'">设置</button>
          </div>
        </div>
      </div>
    </div>

    <div class="main-area">
      <ConfigSidebar v-if="activeTab === 'config'" @config-selected="onServerConfigSelected" />
      <div class="content-area">
        <ConfigGenerator ref="configGeneratorRef" v-show="activeTab === 'config'" @config-updated="handleConfigUpdated"/>
        <DocTagChecker v-show="activeTab === 'checker'" :generated-config="generatedConfig"/>
        <SettingsPage v-show="activeTab === 'settings'" />
      </div>
    </div>

    <input ref="fileInputRef" type="file" accept=".yaml,.yml" style="display: none" @change="onConfigFileSelected" />
  </div>
</template>

<script setup>
import {ref, onMounted} from 'vue';
import GlobalToast from "./components/GlobalToast.vue";
import DocTagChecker from "./components/DocTagChecker.vue";
import ConfigGenerator from "./config-generator/ConfigGenerator.vue";
import ConfigSidebar from "./components/ConfigSidebar.vue";
import SettingsPage from "./components/SettingsPage.vue";
import yaml from 'js-yaml';
import {defaultConfig, mergeWithDefaults} from "./config-generator/utils";
import { loadSettings } from './utils/settings';

const toastRef = ref(null);
const activeTab = ref('config');
const generatedConfig = ref(JSON.parse(JSON.stringify(defaultConfig)));
const configGeneratorRef = ref(null);
const fileInputRef = ref(null);

const handleConfigUpdated = (config) => { generatedConfig.value = config; };

onMounted(() => {
  loadSettings();
  generatedConfig.value = JSON.parse(JSON.stringify(defaultConfig));
  if (toastRef.value) { window.__toast = toastRef.value.toast; }
});

const saveConfig = () => {
  if (!generatedConfig.value) return;
  try {
    const yamlContent = yaml.dump(generatedConfig.value, { indent: 2, skipInvalid: true });
    const blob = new Blob([yamlContent], { type: 'application/x-yaml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'wordformat-config.yaml';
    document.body.appendChild(a); a.click();
    URL.revokeObjectURL(url); document.body.removeChild(a);
    toastRef.value?.toast.success('配置已下载');
  } catch (error) {
    toastRef.value?.toast.error('保存失败：' + error.message);
  }
};

const loadConfig = () => { fileInputRef.value?.click(); };

const onConfigFileSelected = async (e) => {
  const file = e.target.files?.[0]; if (!file) return;
  try {
    const yamlContent = await file.text();
    const config = yaml.load(yamlContent);
    const merged = mergeWithDefaults(config, defaultConfig);
    generatedConfig.value = merged;
    if (configGeneratorRef.value) { configGeneratorRef.value.importConfig(merged); }
    toastRef.value?.toast.success('配置加载成功');
  } catch (error) {
    toastRef.value?.toast.error('加载失败：' + error.message);
  } finally { e.target.value = ''; }
};

const API_BASE = window.__API_BASE__ || '';
const saveConfigToServer = async () => {
  if (!generatedConfig.value) return;
  try {
    const yamlContent = yaml.dump(generatedConfig.value, { indent: 2, skipInvalid: true });
    const fn = prompt('请输入配置文件名（如 my-thesis.yaml）：', 'custom.yaml');
    if (!fn) return;
    const res = await fetch(`${API_BASE}/configs/save`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: fn, content: yamlContent }),
    });
    const json = await res.json();
    if (json.code === 200) { toastRef.value?.toast.success(json.msg); }
    else { toastRef.value?.toast.error(json.msg || '保存失败'); }
  } catch (error) { toastRef.value?.toast.error('保存失败：' + error.message); }
};

const onServerConfigSelected = ({ filename, content }) => {
  try {
    const config = yaml.load(content);
    const merged = mergeWithDefaults(config, defaultConfig);
    generatedConfig.value = merged;
    if (configGeneratorRef.value) { configGeneratorRef.value.importConfig(merged); }
    toastRef.value?.toast.success(`已加载: ${filename}`);
  } catch (error) { toastRef.value?.toast.error('解析失败：' + error.message); }
};
</script>

<style>
/* ══════════════════════════════════════════════════════
   Design tokens — "Typesetter's Cabinet"
   ══════════════════════════════════════════════════════ */
:root {
  --ink: #0d0b09;
  --paper: #161411;
  --surface: #1d1a16;
  --raised: #25221d;
  --border: #302c26;
  --border-hover: #423d35;
  --brass: #b8993b;
  --brass-dim: #8a7230;
  --text: #e5e0d6;
  --text-secondary: #9e988c;
  --text-muted: #6b655b;
  --green: #4a9e6e;
  --green-dim: #2d5c3e;
  --red: #c2574a;
  --red-dim: #5c2a24;
}

* { margin: 0; padding: 0; box-sizing: border-box; }

body {
  background-color: var(--ink);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'PingFang SC', 'Microsoft YaHei', sans-serif;
  -webkit-font-smoothing: antialiased;
  font-size: 14px;
  line-height: 1.5;
}

.app-container { width: 100%; min-height: 100vh; display: flex; flex-direction: column; }

/* ── Nav ── */
.nav-bar {
  background: var(--paper);
  border-bottom: 1px solid var(--border);
  padding: 0 1.75rem;
  position: sticky;
  top: 0;
  z-index: 100;
}
.nav-content {
  max-width: 1480px; margin: 0 auto; width: 100%;
  display: flex; justify-content: space-between; align-items: center; height: 56px;
}
.app-title {
  font-family: 'Georgia', 'Times New Roman', serif;
  font-size: 1.1rem; font-weight: 700; color: var(--text);
  letter-spacing: 0.02em;
}
.nav-actions { display: flex; align-items: center; gap: 0.75rem; }
.config-actions { display: flex; gap: 0.375rem; }

/* ── Buttons ── */
.btn {
  padding: 0.4rem 0.9rem; font-size: 0.78rem; font-weight: 500;
  border: none; border-radius: 6px; cursor: pointer; transition: all .12s;
  font-family: inherit; line-height: 1.4;
}
.btn:disabled { opacity: 0.3; cursor: not-allowed; }
.btn-primary { background: var(--brass); color: var(--ink); font-weight: 600; }
.btn-primary:hover:not(:disabled) { background: #c9a94d; }
.btn-secondary {
  background: var(--surface); color: var(--text-secondary);
  border: 1px solid var(--border);
}
.btn-secondary:hover:not(:disabled) { background: var(--raised); color: var(--text); border-color: var(--border-hover); }
.btn-ghost { background: transparent; color: var(--text-muted); }
.btn-ghost:hover:not(:disabled) { background: var(--surface); color: var(--text); }

/* ── Tabs ── */
.nav-tabs { display: flex; gap: 0; background: var(--surface); border-radius: 8px; padding: 3px; border: 1px solid var(--border); }
.nav-tab {
  padding: 0.35rem 1rem; font-size: 0.78rem; font-weight: 500;
  border: none; border-radius: 6px; background: none;
  color: var(--text-muted); cursor: pointer; transition: all .12s;
  font-family: inherit; position: relative;
}
.nav-tab:hover { color: var(--text); }
.nav-tab.active { background: var(--raised); color: var(--brass); font-weight: 600; }

/* ── Layout ── */
.main-area { display: flex; flex: 1; overflow: hidden; }
.content-area {
  flex: 1; padding: 1.25rem 1.5rem;
  max-width: 1480px; margin: 0 auto; width: 100%;
}

@media (max-width: 768px) {
  .nav-content { flex-direction: column; height: auto; padding: 0.5rem 0; gap: 0.5rem; }
  .nav-tabs { width: 100%; justify-content: center; }
  .content-area { padding: 0.75rem; }
}
</style>
