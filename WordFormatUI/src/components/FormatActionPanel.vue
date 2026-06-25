<template>
  <div class="format-btn-group">
    <button class="btn btn-amber" :disabled="isLoading" @click="$emit('check-format')">格式校验</button>
    <button class="btn btn-green" :disabled="isLoading" @click="$emit('apply-format')">自动格式化</button>
  </div>
</template>

<script setup>
defineProps({ isLoading: Boolean })
defineEmits(['check-format', 'apply-format'])
</script>

<style scoped>
.format-btn-group { display: flex; align-items: center; gap: 0.5rem; flex-wrap: wrap; }
.btn { padding: 6px 12px; font-size: 12px; font-weight: 500; border: none; border-radius: 6px; cursor: pointer; transition: all .2s; font-family: inherit; }
.btn:disabled { opacity: 0.4; cursor: not-allowed; }
.btn-amber { background: var(--brass); color: var(--ink); font-weight: 600; }
.btn-amber:hover:not(:disabled) { background: var(--brass-dim); }
.btn-green { background: var(--green); color: var(--ink); font-weight: 600; }
.btn-green:hover:not(:disabled) { background: var(--green-dim); }
</style>
