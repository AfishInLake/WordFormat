<template>
  <div class="app-container">
    <!-- 全局提示框 -->
    <GlobalToast ref="toastRef" />

    <!-- 导航栏 -->
    <div class="nav-bar">
      <div class="nav-content">
        <h1 class="app-title">WordFormat 工具</h1>
        <div class="nav-actions">
          <!-- 配置操作按钮 -->
          <div class="config-actions" v-if="activeTab === 'config'">
            <button class="btn secondary-btn" @click="saveConfig" :disabled="!generatedConfig">
              保存配置
            </button>
            <button class="btn secondary-btn" @click="loadConfig">
              加载配置
            </button>
            <button class="btn secondary-btn" @click="saveConfigToServer" :disabled="!generatedConfig">
              保存到服务器
            </button>
          </div>

          <!-- 标签页切换 -->
          <div class="nav-tabs">
            <button class="nav-tab" :class="{ active: activeTab === 'config' }" @click="activeTab = 'config'">
              配置生成器
            </button>
            <button class="nav-tab" :class="{ active: activeTab === 'checker' }" @click="activeTab = 'checker'">
              文档标签核对
            </button>
            <button class="nav-tab" :class="{ active: activeTab === 'settings' }" @click="activeTab = 'settings'">
              设置
            </button>
          </div>
        </div>
      </div>
    </div>

    <!-- 主区域：侧边栏 + 内容 -->
    <div class="main-area">
      <ConfigSidebar v-if="activeTab === 'config'" @config-selected="onServerConfigSelected" />

      <div class="content-area">
        <!-- 配置生成器 -->
        <ConfigGenerator ref="configGeneratorRef" v-show="activeTab === 'config'" @config-updated="handleConfigUpdated"/>

        <!-- 文档标签核对工具 -->
        <DocTagChecker v-show="activeTab === 'checker'" :generated-config="generatedConfig"/>

        <!-- 设置页面 -->
        <SettingsPage v-show="activeTab === 'settings'" />
      </div>
    </div>

    <!-- 隐藏的 file input 用于加载配置 -->
    <input
      ref="fileInputRef"
      type="file"
      accept=".yaml,.yml"
      style="display: none"
      @change="onConfigFileSelected"
    />
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

// 全局 toast 引用
const toastRef = ref(null);

// 活动标签页
const activeTab = ref('config');

// 生成的配置
const generatedConfig = ref(JSON.parse(JSON.stringify(defaultConfig)));

// 配置生成器引用
const configGeneratorRef = ref(null);

// 隐藏的 file input 引用（用于加载配置）
const fileInputRef = ref(null);

// 处理配置更新
const handleConfigUpdated = (config) => {
  generatedConfig.value = config;
};

// 生命周期：组件挂载后自动执行
onMounted(() => {
  // 从 localStorage 恢复用户的后端连接配置
  loadSettings();

  // 初始化默认配置
  generatedConfig.value = JSON.parse(JSON.stringify(defaultConfig));

  // 注册全局 toast（供 main.js errorHandler 使用）
  if (toastRef.value) {
    window.__toast = toastRef.value.toast;
  }
});

// 保存配置 → 浏览器下载
const saveConfig = () => {
  if (!generatedConfig.value) return;
  try {
    const yamlContent = yaml.dump(generatedConfig.value, { indent: 2, skipInvalid: true });
    const blob = new Blob([yamlContent], { type: 'application/x-yaml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'wordformat-config.yaml';
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(url);
    document.body.removeChild(a);
    toastRef.value?.toast.success('配置已下载！');
  } catch (error) {
    console.error('保存配置失败:', error);
    toastRef.value?.toast.error('保存配置失败：' + error.message);
  }
};

// 加载配置 → 触发隐藏 file input
const loadConfig = () => {
  fileInputRef.value?.click();
};

// file input 选择文件后的回调
const onConfigFileSelected = async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const yamlContent = await file.text();
    const config = yaml.load(yamlContent);
    const merged = mergeWithDefaults(config, defaultConfig);
    generatedConfig.value = merged;
    if (configGeneratorRef.value) {
      configGeneratorRef.value.importConfig(merged);
    }
    toastRef.value?.toast.success('配置加载成功！');
  } catch (error) {
    console.error('加载配置失败:', error);
    toastRef.value?.toast.error('加载配置失败：' + error.message);
  } finally {
    // 重置 input 以便重复选择同一文件
    e.target.value = '';
  }
};

// 保存配置到服务器 configs 目录
const API_BASE = window.__API_BASE__ || '';
const saveConfigToServer = async () => {
  if (!generatedConfig.value) return;
  try {
    const yamlContent = yaml.dump(generatedConfig.value, { indent: 2, skipInvalid: true });
    const fn = prompt('请输入配置文件名（如 my-thesis.yaml）：', 'custom.yaml');
    if (!fn) return;
    const res = await fetch(`${API_BASE}/configs/save`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: fn, content: yamlContent }),
    });
    const json = await res.json();
    if (json.code === 200) {
      toastRef.value?.toast.success(json.msg);
    } else {
      toastRef.value?.toast.error(json.msg || '保存失败');
    }
  } catch (error) {
    toastRef.value?.toast.error('保存配置失败：' + error.message);
  }
};

// 从服务器侧边栏加载配置
const onServerConfigSelected = ({ filename, content }) => {
  try {
    const config = yaml.load(content);
    const merged = mergeWithDefaults(config, defaultConfig);
    generatedConfig.value = merged;
    if (configGeneratorRef.value) {
      configGeneratorRef.value.importConfig(merged);
    }
    toastRef.value?.toast.success(`已加载配置: ${filename}`);
  } catch (error) {
    toastRef.value?.toast.error('解析配置失败：' + error.message);
  }
};

</script>
<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  background-color: #0f172a;
  color: #e2e8f0;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'PingFang SC', 'Microsoft YaHei', sans-serif;
  -webkit-font-smoothing: antialiased;
}

.app-container {
  width: 100%;
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

/* ── 导航栏 ── */
.nav-bar {
  background-color: #1e293b;
  border-bottom: 1px solid #334155;
  padding: 0 2rem;
  position: sticky;
  top: 0;
  z-index: 100;
  backdrop-filter: blur(12px);
}

.nav-content {
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
  display: flex;
  justify-content: space-between;
  align-items: center;
  height: 64px;
}

.app-title {
  font-family: 'Georgia', 'Times New Roman', serif;
  font-size: 1.4rem;
  font-weight: 600;
  color: #f1f5f9;
  margin: 0;
  letter-spacing: -0.02em;
}

.nav-actions {
  display: flex;
  align-items: center;
  gap: 0.75rem;
}

.config-actions {
  display: flex;
  gap: 0.5rem;
}

.nav-tabs {
  display: flex;
  gap: 0.25rem;
  background: #0f172a;
  border-radius: 8px;
  padding: 3px;
}

/* ── 通用按钮 ── */
.btn {
  padding: 0.5rem 1rem;
  font-size: 0.8125rem;
  font-weight: 500;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.2s ease;
  font-family: inherit;
}

.btn:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}

.primary-btn {
  background-color: #22c55e;
  color: #052e16;
}

.primary-btn:hover:not(:disabled) {
  background-color: #16a34a;
}

.secondary-btn {
  background-color: #334155;
  color: #cbd5e1;
  border: 1px solid #475569;
}

.secondary-btn:hover:not(:disabled) {
  background-color: #475569;
  border-color: #64748b;
}

/* ── 标签页 ── */
.nav-tab {
  padding: 0.45rem 1.25rem;
  font-size: 0.875rem;
  font-weight: 500;
  border: none;
  border-radius: 6px;
  background-color: transparent;
  color: #94a3b8;
  cursor: pointer;
  transition: all 0.2s ease;
  font-family: inherit;
}

.nav-tab:hover {
  background-color: #1e293b;
  color: #e2e8f0;
}

.nav-tab.active {
  background-color: #22c55e;
  color: #052e16;
}

/* ── 主区域（侧边栏 + 内容）── */
.main-area {
  display: flex;
  flex: 1;
  overflow: hidden;
}

/* ── 内容区域 ── */
.content-area {
  flex: 1;
  padding: 1.5rem 2rem;
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
}

/* ── 响应式 ── */
@media (max-width: 768px) {
  .nav-content {
    flex-direction: column;
    height: auto;
    padding: 1rem 0;
    gap: 0.75rem;
  }

  .nav-tabs {
    width: 100%;
    justify-content: center;
  }

  .content-area {
    padding: 1rem;
  }
}
</style>
