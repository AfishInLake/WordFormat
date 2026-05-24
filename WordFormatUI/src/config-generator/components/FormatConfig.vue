<template>
  <div class="format-config">
    <div class="grid grid-cols-2 gap-4">
      <div class="form-item"><label>对齐方式:</label><select v-model="config.alignment"><option v-for="o in alignmentOptions" :key="o" :value="o">{{ o }}</option></select></div>
      <div class="form-item"><label>段前间距:</label><input type="text" v-model="config.space_before"></div>
      <div class="form-item"><label>段后间距:</label><input type="text" v-model="config.space_after"></div>
      <div class="form-item"><label>行距类型:</label><select v-model="config.line_spacingrule"><option v-for="o in lineSpacingRuleOptions" :key="o" :value="o">{{ o }}</option></select></div>
      <div class="form-item"><label>行距参数:</label><input type="text" v-model="config.line_spacing"></div>
      <div class="form-item"><label>文本之前:</label><input type="text" v-model="config.left_indent"></div>
      <div class="form-item"><label>文本之后:</label><input type="text" v-model="config.right_indent"></div>
      <div class="form-item"><label>段落首行缩进:</label><input type="text" v-model="config.first_line_indent"></div>
      <div class="form-item"><label>内置样式名称:</label><input type="text" v-model="config.builtin_style_name"></div>
      <div class="form-item"><label>中文字体名称:</label><select v-model="config.chinese_font_name"><option v-for="o in chineseFontOptions" :key="o" :value="o">{{ o }}</option></select></div>
      <div class="form-item"><label>英文字体名称:</label><select v-model="config.english_font_name"><option v-for="o in englishFontOptions" :key="o" :value="o">{{ o }}</option></select></div>
      <div class="form-item"><label>字号:</label><select v-model="config.font_size"><option v-for="o in fontSizeOptions" :key="o" :value="o">{{ o }}</option></select></div>
      <div class="form-item"><label>字体颜色:</label><input type="text" v-model="config.font_color"></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.bold"> 加粗</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.italic"> 斜体</label></div>
      <div class="form-item"><label><input type="checkbox" v-model="config.underline"> 下划线</label></div>
    </div>
  </div>
</template>

<script setup>
import { alignmentOptions, lineSpacingRuleOptions, chineseFontOptions, englishFontOptions, fontSizeOptions } from '../utils'
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
.form-item { margin-bottom: 6px; display: flex; align-items: center; gap: 8px; }
.form-item label { font-weight: 500; color: #94a3b8; font-size: 13px; white-space: nowrap; min-width: 80px; }
.form-item input[type="text"], .form-item input[type="number"], .form-item select { width: 90px; min-width: 80px; padding: 6px 8px; border: 1px solid #475569; border-radius: 6px; font-size: 13px; background: #0f172a; color: #e2e8f0; outline: none; transition: border-color 0.2s; }
.form-item input:focus, .form-item select:focus { border-color: #22c55e; box-shadow: 0 0 0 2px rgba(34,197,94,.15); }
.form-item input[type="text"]:disabled, .form-item select:disabled { opacity: 0.4; }
.form-item:has(input[type="checkbox"]) { width: auto; }
.form-item:has(input[type="checkbox"]) label { min-width: auto; }
.form-item input[type="checkbox"] { margin-right: 5px; accent-color: #22c55e; }
.grid { display: grid; }
.grid-cols-2 { grid-template-columns: repeat(4, 1fr); }
.gap-4 { gap: 10px; }
@media (min-width: 1200px) { .grid-cols-2 { grid-template-columns: repeat(5, 1fr); } }
@media (min-width: 1600px) { .grid-cols-2 { grid-template-columns: repeat(6, 1fr); } }
@media (max-width: 768px) { .grid-cols-2 { grid-template-columns: 1fr; } .form-item { flex-direction: column; align-items: flex-start; } .form-item label { min-width: auto; } .form-item input[type="text"], .form-item input[type="number"], .form-item select { width: 120px; } }
</style>
