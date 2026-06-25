<template>
  <div class="references-config">
    <h3>参考文献标题</h3>
    <FormatConfig :config="config.title" />
    <h3 class="mt-4">参考文献内容</h3>
    <FormatConfig :config="config.content" />
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
h3 { margin: 0 0 10px 0; font-size: 13px; font-weight: 600; color: var(--text); }
.mt-4 { margin-top: 16px; }
</style>
