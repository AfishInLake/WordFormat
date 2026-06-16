<template>
  <div class="numbering-config">
    <div class="master-switch">
      <label class="switch-label">
        <input type="checkbox" v-model="config.enabled" />
        <span>启用标题自动编号</span>
      </label>
    </div>
    <div v-if="config.enabled" class="levels-config">
      <div class="level-item">
        <h4>一级标题</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.level_1.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>编号模板</label><input type="text" v-model="config.level_1.template" :disabled="!config.level_1.enabled" /><span class="hint">如 %1</span></div>
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>后缀</label><select v-model="config.level_1.suffix" :disabled="!config.level_1.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_1.numbering_indent" :disabled="!config.level_1.enabled" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_1.text_indent" :disabled="!config.level_1.enabled" /><span class="hint">如 0.75cm</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>二级标题</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.level_2.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>编号模板</label><input type="text" v-model="config.level_2.template" :disabled="!config.level_2.enabled" /><span class="hint">如 %1.%2</span></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>后缀</label><select v-model="config.level_2.suffix" :disabled="!config.level_2.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_2.numbering_indent" :disabled="!config.level_2.enabled" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_2.text_indent" :disabled="!config.level_2.enabled" /><span class="hint">如 0.75cm</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>三级标题</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.level_3.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>编号模板</label><input type="text" v-model="config.level_3.template" :disabled="!config.level_3.enabled" /><span class="hint">如 %1.%2.%3</span></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>后缀</label><select v-model="config.level_3.suffix" :disabled="!config.level_3.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_3.numbering_indent" :disabled="!config.level_3.enabled" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_3.text_indent" :disabled="!config.level_3.enabled" /><span class="hint">如 0.75cm</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>参考文献</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.references.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>编号模板</label><input type="text" v-model="config.references.template" :disabled="!config.references.enabled" /><span class="hint">如 [%1]</span></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>后缀</label><select v-model="config.references.suffix" :disabled="!config.references.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>编号缩进</label><input type="text" v-model="config.references.numbering_indent" :disabled="!config.references.enabled" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>文本缩进</label><input type="text" v-model="config.references.text_indent" :disabled="!config.references.enabled" /><span class="hint">如 0.75cm</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>题注编号</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.captions.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.captions.enabled }"><label>分隔符</label><input type="text" v-model="config.captions.separator" :disabled="!config.captions.enabled" /><span class="hint">章节号与编号间的分隔符，如 . - :</span></div>
          <label class="inline-check" :class="{ disabled: !config.captions.enabled }"><input type="checkbox" v-model="config.captions.label_number_space" :disabled="!config.captions.enabled" /><span>标签与编号间加空格</span></label>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
.master-switch { margin-bottom: 16px; }
.switch-label { display: inline-flex; align-items: center; gap: 8px; cursor: pointer; font-weight: 600; font-size: 13px; color: #22c55e; }
.switch-label input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; accent-color: #22c55e; }
.levels-config { display: flex; flex-direction: column; gap: 12px; }
.level-item { border: 1px solid #334155; border-radius: 8px; padding: 12px; background: #0f172a; }
.level-item h4 { margin: 0 0 10px 0; font-size: 13px; font-weight: 600; color: #e2e8f0; }
.level-controls { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.inline-check { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; font-size: 13px; color: #cbd5e1; white-space: nowrap; }
.inline-check input[type="checkbox"] { width: 14px; height: 14px; cursor: pointer; accent-color: #22c55e; }
.field { display: flex; align-items: center; gap: 6px; }
.field.disabled { opacity: 0.4; pointer-events: none; }
.field label { font-size: 12px; color: #64748b; white-space: nowrap; }
.field input[type="text"] { width: 80px; padding: 4px 8px; border: 1px solid #475569; border-radius: 5px; font-size: 13px; font-family: monospace; background: #1e293b; color: #e2e8f0; outline: none; }
.field input[type="text"]:focus { border-color: #22c55e; box-shadow: 0 0 0 2px rgba(34,197,94,.15); }
.field select { padding: 4px 8px; border: 1px solid #475569; border-radius: 5px; font-size: 13px; background: #1e293b; color: #e2e8f0; outline: none; }
.hint { font-size: 11px; color: #64748b; white-space: nowrap; }
</style>
