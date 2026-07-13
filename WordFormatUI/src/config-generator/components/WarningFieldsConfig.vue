<template>
  <div class="warning-fields-config">
    <div class="select-all-bar">
      <label class="select-all-label">
        <input type="checkbox" :checked="allChecked" @change="toggleAll" />
        <span>{{ allChecked ? '取消全选' : '全选' }}</span>
      </label>
    </div>
    <div class="grid grid-cols-3 gap-2">
      <div class="form-item"><label><input type="checkbox" v-model="config.bold"> 加粗</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.italic"> 斜体</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.underline"> 下划线</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.font_size"> 字号</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.font_name_cn"> 中文字体名称</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.font_name_en"> 英文字体名称</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.font_color"> 字体颜色</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.alignment"> 对齐方式</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.space_before"> 段前间距</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.space_after"> 段后间距</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.line_spacing"> 行距</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.line_spacing_rule"> 行距类型</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.left_indent"> 文本之前</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.right_indent"> 文本之后</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.first_line_indent"> 段落首行缩进</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.builtin_style_name"> 内置样式名称</label></div>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
const props = defineProps({ config: { type: Object, required: true } })
const FIELDS = ['bold','italic','underline','font_size','font_name_cn','font_name_en','font_color','alignment','space_before','space_after','line_spacing','line_spacing_rule','left_indent','right_indent','first_line_indent','builtin_style_name']
const allChecked = computed(() => FIELDS.every(f => props.config[f]))
function toggleAll() { const n = !allChecked.value; FIELDS.forEach(f => { props.config[f] = n }) }
</script>

<style scoped>
.select-all-bar { margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px solid var(--border); }
.select-all-label { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; font-weight: 600; font-size: 13px; color: var(--brass); user-select: none; }
.select-all-label input[type="checkbox"] { width: 15px; height: 15px; cursor: pointer; accent-color: var(--brass); }
.form-item { margin-bottom: 7px; display: flex; flex-direction: row; align-items: center; gap: 4px; }
.form-item label { font-weight: 500; color: var(--text-secondary); font-size: 13px; white-space: nowrap; cursor: pointer; user-select: none; }
.form-item input[type="checkbox"] { margin-right: 6px; accent-color: var(--brass); width: 14px; height: 14px; cursor: pointer; }
.grid { display: grid; }
.grid-cols-3 { grid-template-columns: repeat(5, 1fr); }
.gap-2 { gap: 4px; }
@media (min-width: 1024px) { .grid-cols-3 { grid-template-columns: repeat(6, 1fr); } }
@media (min-width: 1400px) { .grid-cols-3 { grid-template-columns: repeat(7, 1fr); } }
@media (max-width: 768px) { .grid-cols-3 { grid-template-columns: repeat(2, 1fr); } }
</style>
