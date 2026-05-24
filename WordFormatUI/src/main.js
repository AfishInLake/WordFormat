import { createApp } from "vue";
import App from "./App.vue";

const app = createApp(App);

// ===== 全局 toast =====
const $toast = {
  info: (msg, dur) => window.__toast?.info(msg, dur),
  success: (msg, dur) => window.__toast?.success(msg, dur),
  warn: (msg, dur) => window.__toast?.warn(msg, dur),
  error: (msg, dur) => window.__toast?.error(msg, dur),
};

app.config.globalProperties.$toast = $toast;

// ===== 全局未捕获异常 =====
app.config.errorHandler = (err, _instance, info) => {
  console.error('[Vue Error]', err, info);
  $toast.error(`[${info}] ${err.message || err}`);
};

window.addEventListener('error', (e) => {
  console.error('[Global Error]', e.error);
  $toast.error(e.error?.message || e.message);
});

window.addEventListener('unhandledrejection', (e) => {
  console.error('[Unhandled Rejection]', e.reason);
  $toast.error(e.reason?.message || String(e.reason));
});

app.mount("#app");
