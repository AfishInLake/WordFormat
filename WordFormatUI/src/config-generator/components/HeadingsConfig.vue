<template>
  <div class="headings-config">
    <h3>一级标题</h3>
    <FormatConfig :config="config.level_1" />
    <h3 class="mt-4">二级标题</h3>
    <FormatConfig :config="config.level_2" />
    <h3 class="mt-4">三级标题</h3>
    <FormatConfig :config="config.level_3" />
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
