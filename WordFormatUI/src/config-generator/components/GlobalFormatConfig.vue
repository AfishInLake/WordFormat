<template>
  <div class="global-format-config">
    <FormatConfig :config="config" />
    <div v-if="showApplyButton" class="mt-3">
      <button @click="$emit('apply-to-all')" class="apply-btn">应用到所有配置</button>
    </div>
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'
defineProps({ config: { type: Object, required: true }, showApplyButton: { type: Boolean, default: false } })
defineEmits(['apply-to-all'])
</script>

<style scoped>
.apply-btn {
  padding: 7px 16px; border: 1px solid var(--brass-dim); border-radius: 6px;
  font-size: 12px; font-weight: 500; cursor: pointer; font-family: inherit;
  background: transparent; color: var(--brass); transition: all .12s;
}
.apply-btn:hover { background: var(--brass); color: var(--ink); }
.mt-3 { margin-top: 12px; }
</style>
