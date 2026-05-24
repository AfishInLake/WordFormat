<template>
  <Teleport to="body">
    <div class="toast-container">
      <TransitionGroup name="toast">
        <div
          v-for="item in toasts"
          :key="item.id"
          class="toast-item"
          :class="`toast-${item.type}`"
        >
          <span class="toast-msg">{{ item.message }}</span>
          <button class="toast-close" @click="remove(item.id)">&times;</button>
        </div>
      </TransitionGroup>
    </div>
  </Teleport>
</template>

<script setup>
import { ref } from 'vue';

const toasts = ref([]);
let nextId = 0;

const iconMap = {
  info: '',
  success: '',
  warn: '',
  error: '',
};

function add(type, message, duration = 4000) {
  const id = ++nextId;
  toasts.value.push({ id, type, message });
  if (duration > 0) {
    setTimeout(() => remove(id), duration);
  }
}

function remove(id) {
  const i = toasts.value.findIndex(t => t.id === id);
  if (i >= 0) toasts.value.splice(i, 1);
}

const toast = {
  info: (msg, dur) => add('info', msg, dur),
  success: (msg, dur) => add('success', msg, dur),
  warn: (msg, dur) => add('warn', msg, dur),
  error: (msg, dur) => add('error', msg, dur ?? 8000),
};

defineExpose({ toast });
</script>

<style>
.toast-container {
  position: fixed;
  top: 16px;
  right: 16px;
  z-index: 99999;
  display: flex;
  flex-direction: column;
  gap: 8px;
  pointer-events: none;
}
.toast-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 16px;
  border-radius: 10px;
  background: #1e293b;
  box-shadow: 0 8px 32px rgba(0,0,0,.35);
  min-width: 280px;
  max-width: 480px;
  pointer-events: auto;
  font-size: 14px;
  border: 1px solid #334155;
}
.toast-msg { flex: 1; color: #e2e8f0; }
.toast-close {
  background: none; border: none; font-size: 22px;
  cursor: pointer; color: #64748b; padding: 0 4px; line-height: 1;
}
.toast-close:hover { color: #e2e8f0; }

.toast-info    { border-left: 4px solid #3b82f6; }
.toast-success { border-left: 4px solid #22c55e; }
.toast-warn    { border-left: 4px solid #f59e0b; }
.toast-error   { border-left: 4px solid #ef4444; }

.toast-enter-active { transition: all .3s ease; }
.toast-leave-active { transition: all .2s ease; }
.toast-enter-from { opacity: 0; transform: translateX(40px); }
.toast-leave-to   { opacity: 0; transform: translateX(40px); }
</style>
