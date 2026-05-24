/**
 * 应用设置 — 响应式 + localStorage 持久化
 *
 * 用户可在「设置」页面修改后端地址和端口，修改后即时生效并自动保存。
 * 关闭应用再打开时，自动恢复上次的配置。
 */

import { reactive, computed } from 'vue';

const STORAGE_KEY = 'wordformat-settings';

// ── 响应式状态 ──
export const backendSettings = reactive({
  host: '127.0.0.1',
  port: 8000,
});

// ── 派生值 ──
export const backendBaseUrl = computed(() => {
  const h = backendSettings.host || '127.0.0.1';
  const p = backendSettings.port || 8000;
  return `http://${h}:${p}`;
});

// ── 从 localStorage 加载 ──
export function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const saved = JSON.parse(raw);
      if (saved.host) backendSettings.host = saved.host;
      if (saved.port != null) backendSettings.port = saved.port;
    }
  } catch {
    // 解析失败使用默认值
  }
}

// ── 保存到 localStorage ──
export function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    host: backendSettings.host,
    port: backendSettings.port,
  }));
}

// ── 重置为默认值 ──
export function resetSettings() {
  backendSettings.host = '127.0.0.1';
  backendSettings.port = 8000;
  saveSettings();
}
