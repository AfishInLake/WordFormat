<template>
  <div class="format-config">
    <div class="field-group">
      <div class="group-label">段落格式</div>
      <div class="field-row">
        <div class="field"><label>对齐</label><select v-model="config.alignment"><option v-for="o in alignmentOptions" :key="o" :value="o">{{ o }}</option></select></div>
        <div class="field"><label>段前距</label><input type="text" v-model="config.space_before"></div>
        <div class="field"><label>段后距</label><input type="text" v-model="config.space_after"></div>
        <div class="field"><label>行距类型</label><select v-model="config.line_spacingrule"><option v-for="o in lineSpacingRuleOptions" :key="o" :value="o">{{ o }}</option></select></div>
        <div class="field"><label>行距值</label><input type="text" v-model="config.line_spacing"></div>
      </div>
      <div class="field-row">
        <div class="field"><label>左缩进</label><input type="text" v-model="config.left_indent"></div>
        <div class="field"><label>右缩进</label><input type="text" v-model="config.right_indent"></div>
        <div class="field"><label>首行缩进</label><input type="text" v-model="config.first_line_indent"></div>
        <div class="field"><label>样式名</label><input type="text" v-model="config.builtin_style_name"></div>
      </div>
    </div>

    <div class="field-group">
      <div class="group-label">文字格式</div>
      <div class="field-row">
        <div class="field"><label>中文字体</label><select v-model="config.chinese_font_name"><option v-for="o in chineseFontOptions" :key="o" :value="o">{{ o }}</option></select></div>
        <div class="field"><label>英文字体</label><select v-model="config.english_font_name"><option v-for="o in englishFontOptions" :key="o" :value="o">{{ o }}</option></select></div>
        <div class="field"><label>字号</label><select v-model="config.font_size"><option v-for="o in fontSizeOptions" :key="o" :value="o">{{ o }}</option></select></div>
        <div class="field"><label>颜色</label><input type="text" v-model="config.font_color"></div>
      </div>
      <div class="field-row checks">
        <label class="check-item"><input type="checkbox" v-model="config.bold"><span>加粗</span></label>
        <label class="check-item"><input type="checkbox" v-model="config.italic"><span>斜体</span></label>
        <label class="check-item"><input type="checkbox" v-model="config.underline"><span>下划线</span></label>
      </div>
    </div>
  </div>
</template>

<script setup>
import { alignmentOptions, lineSpacingRuleOptions, chineseFontOptions, englishFontOptions, fontSizeOptions } from '../utils'
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
.format-config { display: flex; flex-direction: column; gap: 12px; }

.field-group {
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 12px 14px;
  background: var(--ink);
}
.group-label {
  font-size: 10px; font-weight: 600; color: var(--text-muted);
  text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 10px;
}

.field-row { display: flex; gap: 10px; flex-wrap: wrap; }
.field-row.checks { gap: 20px; padding-top: 6px; }

.field { display: flex; flex-direction: column; gap: 3px; }
.field label {
  font-size: 11px; font-weight: 500; color: var(--text-muted);
  padding-left: 1px;
}
.field input[type="text"],
.field input[type="number"],
.field select {
  width: 130px; padding: 6px 10px;
  border: 1px solid var(--border); border-radius: 6px;
  font-size: 13px; background: var(--paper); color: var(--text);
  outline: none; transition: border-color .12s, box-shadow .12s;
  font-family: inherit;
}
.field input:focus, .field select:focus {
  border-color: var(--brass); box-shadow: 0 0 0 3px rgba(184,153,59,.08);
}
.field input:disabled, .field select:disabled { opacity: .35; }

.check-item {
  display: inline-flex; align-items: center; gap: 6px;
  font-size: 13px; color: var(--text-secondary); cursor: pointer; user-select: none;
}
.check-item input[type="checkbox"] {
  width: 15px; height: 15px; accent-color: var(--brass); cursor: pointer;
}
</style>
