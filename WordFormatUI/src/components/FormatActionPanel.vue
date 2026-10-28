<template>
  <div class="format-btn-group">
    <button class="btn btn-amber" :disabled="isLoading" @click="$emit('check-format')">格式校验</button>
    <button class="btn btn-green" :disabled="isLoading" @click="$emit('apply-format')">自动格式化</button>
    <label class="sync-toggle" title="等价 af --sync：按 JSON 行的增删改在原段落/表格身份上落盘，而不是整篇重新排版">
      <input type="checkbox" v-model="syncProxy" />
      <span>结构同步</span>
      <span v-if="structureDirty" class="sync-badge">{{ dirtyHint }}</span>
    </label>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  isLoading: Boolean,
  syncMode: Boolean,
  structureDirty: Boolean,
  dirtyHint: String
})
const emit = defineEmits(['check-format', 'apply-format', 'update:syncMode'])

const syncProxy = computed({
  get: () => props.syncMode,
  set: (value) => emit('update:syncMode', value)
})
</script>

<style scoped>
.format-btn-group { display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap; }
.btn { padding: 6px 12px; font-size: 12px; font-weight: 500; border: none; border-radius: 6px; cursor: pointer; transition: all .2s; font-family: inherit; }
.btn:disabled { opacity: 0.4; cursor: not-allowed; }
.btn-amber { background: var(--brass); color: var(--ink); font-weight: 600; }
.btn-amber:hover:not(:disabled) { background: var(--brass-dim); }
.btn-green { background: var(--green); color: var(--ink); font-weight: 600; }
.btn-green:hover:not(:disabled) { background: var(--green-dim); }
.sync-toggle { display: inline-flex; align-items: center; gap: 0.35rem; font-size: 12px; color: var(--text-secondary); cursor: pointer; white-space: nowrap; }
.sync-toggle input { accent-color: var(--brass); cursor: pointer; }
.sync-badge { padding: 1px 6px; border-radius: 999px; font-size: 11px; background: var(--red-dim); color: var(--red); border: 1px solid var(--red); }
</style>
