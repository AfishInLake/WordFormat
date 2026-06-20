<template>
  <div class="figures-config">
    <div class="grid grid-cols-2 gap-4 mb-4">
      <div class="form-item">
        <label>图注编号前缀:</label>
        <input type="text" v-model="config.caption_prefix">
      </div>
      <div class="form-item">
        <label><input type="checkbox" v-model="config.rules.caption_numbering.enabled"> 启用编号检查</label>
      </div>
    </div>
    <h5 class="subsection-title">题注格式</h5>
    <FormatConfig :config="config" />
    <h5 class="subsection-title mt-3">图片段落格式（包含内联图片的段落）</h5>
    <FormatConfig :config="config.image" />
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'

const props = defineProps({
  config: {
    type: Object,
    required: true
  }
})
</script>

<style scoped>
.form-item {
  margin-bottom: 5px;
  display: flex;
  align-items: center;
  gap: 8px;
}
.form-item label {
  font-weight: 500; color: #555; font-size: 13px; white-space: nowrap; min-width: 100px;
}
.form-item input[type="text"], .form-item input[type="number"], .form-item select {
  width: 90px; min-width: 80px; padding: 6px 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px;
}
.form-item input[type="checkbox"] { margin-right: 5px; }
.grid { display: grid; }
.grid-cols-2 { grid-template-columns: repeat(4, 1fr); }
.gap-4 { gap: 10px; }
.mb-4 { margin-bottom: 12px; }
.subsection-title { font-size: 12px; font-weight: 600; color: #94a3b8; margin-bottom: 4px; }
.mt-3 { margin-top: 8px; }
@media (min-width: 1024px) { .grid-cols-2 { grid-template-columns: repeat(5, 1fr); } }
@media (min-width: 1200px) { .grid-cols-2 { grid-template-columns: repeat(6, 1fr); } }
@media (max-width: 768px) {
  .grid-cols-2 { grid-template-columns: 1fr; }
  .form-item { flex-direction: column; align-items: flex-start; }
  .form-item label { min-width: auto; }
  .form-item input[type="text"], .form-item input[type="number"], .form-item select { width: 120px; }
}
</style>
