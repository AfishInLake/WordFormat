<template>
  <div class="global-format-config">
    <FormatConfig :config="config" />
    <div v-if="showApplyButton" class="mt-4">
      <button @click="$emit('apply-to-all')" class="btn btn-green">应用到所有配置</button>
    </div>
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'
defineProps({ config: { type: Object, required: true }, showApplyButton: { type: Boolean, default: false } })
defineEmits(['apply-to-all'])
</script>

<style scoped>
.btn-green { padding: 8px 16px; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; background: #22c55e; color: #052e16; font-family: inherit; transition: background .2s; }
.btn-green:hover { background: #16a34a; }
.mt-4 { margin-top: 14px; }
</style>
