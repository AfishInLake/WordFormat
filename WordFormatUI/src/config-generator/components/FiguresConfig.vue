<template>
  <div class="figures-config">
    <div class="rules-grid">
      <div class="rule-field"><label>编号前缀</label><input type="text" v-model="config.caption_prefix"></div>
      <label class="rule-check"><input type="checkbox" v-model="config.rules.caption_numbering.enabled"> 启用编号检查</label>
      <div class="rule-field"><label>分隔符</label><input type="text" v-model="config.rules.caption_numbering.separator" :disabled="!config.rules.caption_numbering.enabled"></div>
      <label class="rule-check"><input type="checkbox" v-model="config.rules.caption_numbering.label_number_space" :disabled="!config.rules.caption_numbering.enabled"> 标签-编号间空格</label>
    </div>
    <h5>题注格式</h5>
    <FormatConfig :config="config" />
    <h5 class="mt-3">图片段落格式</h5>
    <FormatConfig :config="config.image" />
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
.rules-grid { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-bottom: 10px; padding: 10px 12px; border: 1px solid var(--border); border-radius: 6px; background: var(--ink); }
.rule-check { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; color: var(--text-secondary); cursor: pointer; user-select: none; }
.rule-check input[type="checkbox"] { width: 14px; height: 14px; accent-color: var(--brass); cursor: pointer; }
.rule-field { display: flex; align-items: center; gap: 4px; }
.rule-field label { font-size: 11px; color: var(--text-muted); }
.rule-field input[type="text"] { width: 60px; padding: 4px 7px; border: 1px solid var(--border); border-radius: 4px; font-size: 12px; background: var(--paper); color: var(--text); outline: none; }
.rule-field input:focus { border-color: var(--brass); }
.rule-field input:disabled { opacity: .35; }
h5 { font-size: 11px; font-weight: 600; color: var(--text-muted); margin-bottom: 6px; }
.mt-3 { margin-top: 10px; }
</style>
