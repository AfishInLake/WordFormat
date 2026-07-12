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
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_1.numbering_indent" :disabled="!config.level_1.enabled" placeholder="可选" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_1.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_1.text_indent" :disabled="!config.level_1.enabled" placeholder="可选" /><span class="hint">悬挂缩进量</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>二级标题</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.level_2.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>编号模板</label><input type="text" v-model="config.level_2.template" :disabled="!config.level_2.enabled" /><span class="hint">如 %1.%2</span></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>后缀</label><select v-model="config.level_2.suffix" :disabled="!config.level_2.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_2.numbering_indent" :disabled="!config.level_2.enabled" placeholder="可选" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_2.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_2.text_indent" :disabled="!config.level_2.enabled" placeholder="可选" /><span class="hint">悬挂缩进量</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>三级标题</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.level_3.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>编号模板</label><input type="text" v-model="config.level_3.template" :disabled="!config.level_3.enabled" /><span class="hint">如 %1.%2.%3</span></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>后缀</label><select v-model="config.level_3.suffix" :disabled="!config.level_3.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>编号缩进</label><input type="text" v-model="config.level_3.numbering_indent" :disabled="!config.level_3.enabled" placeholder="可选" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.level_3.enabled }"><label>文本缩进</label><input type="text" v-model="config.level_3.text_indent" :disabled="!config.level_3.enabled" placeholder="可选" /><span class="hint">悬挂缩进量</span></div>
        </div>
      </div>
      <div class="level-item">
        <h4>参考文献</h4>
        <div class="level-controls">
          <label class="inline-check"><input type="checkbox" v-model="config.references.enabled" /><span>启用</span></label>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>编号模板</label><input type="text" v-model="config.references.template" :disabled="!config.references.enabled" /><span class="hint">如 [%1]</span></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>后缀</label><select v-model="config.references.suffix" :disabled="!config.references.enabled"><option value="space">空格</option><option value="tab">制表符</option><option value="nothing">无</option></select></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>编号缩进</label><input type="text" v-model="config.references.numbering_indent" :disabled="!config.references.enabled" placeholder="可选" /><span class="hint">如 0.75cm</span></div>
          <div class="field" :class="{ disabled: !config.references.enabled }"><label>文本缩进</label><input type="text" v-model="config.references.text_indent" :disabled="!config.references.enabled" placeholder="可选" /><span class="hint">悬挂缩进量</span></div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
.master-switch { margin-bottom: 14px; }
.switch-label { display: inline-flex; align-items: center; gap: 8px; cursor: pointer; font-weight: 600; font-size: 13px; color: var(--brass); user-select: none; }
.switch-label input[type="checkbox"] { width: 15px; height: 15px; cursor: pointer; accent-color: var(--brass); }
.levels-config { display: flex; flex-direction: column; gap: 10px; }
.level-item { border: 1px solid var(--border); border-radius: 8px; padding: 12px; background: var(--surface); }
.level-item h4 { margin: 0 0 8px 0; font-size: 13px; font-weight: 600; color: var(--text); }
.level-controls { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.inline-check { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; font-size: 13px; color: var(--text-secondary); white-space: nowrap; user-select: none; }
.inline-check input[type="checkbox"] { width: 14px; height: 14px; cursor: pointer; accent-color: var(--brass); }
.inline-check.disabled { opacity: .35; pointer-events: none; }
.field { display: flex; align-items: center; gap: 5px; }
.field.disabled { opacity: .35; pointer-events: none; }
.field label { font-size: 11px; color: var(--text-muted); white-space: nowrap; }
.field input[type="text"] { width: 80px; padding: 4px 8px; border: 1px solid var(--border); border-radius: 5px; font-size: 12px; font-family: 'SF Mono', Menlo, Monaco, Consolas, monospace; background: var(--ink); color: var(--text); outline: none; transition: border-color .12s; }
.field input[type="text"]:focus { border-color: var(--brass); }
.field select { padding: 4px 8px; border: 1px solid var(--border); border-radius: 5px; font-size: 13px; background: var(--ink); color: var(--text); outline: none; }
.hint { font-size: 11px; color: var(--text-muted); white-space: nowrap; }
</style>
