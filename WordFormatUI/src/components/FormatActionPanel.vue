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
.btn-amber { background: #d97706; color: #fef3c7; }
.btn-amber:hover:not(:disabled) { background: #b45309; }
.btn-green { background: #22c55e; color: #052e16; }
.btn-green:hover:not(:disabled) { background: #16a34a; }
</style>
