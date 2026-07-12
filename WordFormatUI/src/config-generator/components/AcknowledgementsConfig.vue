<template>
  <div class="acknowledgements-config">
    <h3>致谢标题</h3>
    <FormatConfig :config="config.title" />
    <h3 class="mt-4">致谢内容</h3>
    <FormatConfig :config="config.body" />
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
