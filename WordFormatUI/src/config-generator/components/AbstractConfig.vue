<template>
  <div class="abstract-config">
    <h3>中文摘要</h3>
    <h4>中文标题</h4>
    <FormatConfig :config="config.chinese.title" />
    <h4 class="mt-4">中文内容</h4>
    <FormatConfig :config="config.chinese.body" />

    <h3 class="mt-4">英文摘要</h3>
    <h4>英文标题</h4>
    <FormatConfig :config="config.english.title" />
    <h4 class="mt-4">英文内容</h4>
    <FormatConfig :config="config.english.body" />

    <h3 class="mt-4">关键词配置</h3>
    <h4>中文关键词</h4>
    <div class="rules-grid">
      <label class="rule-check"><input type="checkbox" v-model="config.chinese.keywords.rules.keyword_count.enabled"> 启用数量检查</label>
      <div class="rule-field"><label>最少</label><input type="number" v-model.number="config.chinese.keywords.rules.keyword_count.count_min"></div>
      <div class="rule-field"><label>最多</label><input type="number" v-model.number="config.chinese.keywords.rules.keyword_count.count_max"></div>
      <label class="rule-check"><input type="checkbox" v-model="config.chinese.keywords.rules.trailing_punctuation.enabled"> 禁止末尾标点</label>
      <div class="rule-field"><label>禁止字符</label><input type="text" v-model="config.chinese.keywords.rules.trailing_punctuation.forbidden_chars" :disabled="!config.chinese.keywords.rules.trailing_punctuation.enabled"></div>
    </div>
    <h5>中文关键词内容格式</h5>
    <FormatConfig :config="config.chinese.keywords" />
    <h5 class="mt-3">中文关键词标签（"关键词："）</h5>
    <FormatConfig :config="config.chinese.keywords.label" />

    <h4 class="mt-4">英文关键词</h4>
    <div class="rules-grid">
      <label class="rule-check"><input type="checkbox" v-model="config.english.keywords.rules.keyword_count.enabled"> 启用数量检查</label>
      <div class="rule-field"><label>最少</label><input type="number" v-model.number="config.english.keywords.rules.keyword_count.count_min"></div>
      <div class="rule-field"><label>最多</label><input type="number" v-model.number="config.english.keywords.rules.keyword_count.count_max"></div>
      <label class="rule-check"><input type="checkbox" v-model="config.english.keywords.rules.trailing_punctuation.enabled"> 禁止末尾标点</label>
      <div class="rule-field"><label>禁止字符</label><input type="text" v-model="config.english.keywords.rules.trailing_punctuation.forbidden_chars" :disabled="!config.english.keywords.rules.trailing_punctuation.enabled"></div>
    </div>
    <h5>英文关键词内容格式</h5>
    <FormatConfig :config="config.english.keywords" />
    <h5 class="mt-3">英文关键词标签（"Keywords:"）</h5>
    <FormatConfig :config="config.english.keywords.label" />
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'
defineProps({ config: { type: Object, required: true } })
</script>

<style scoped>
h3 { margin: 0 0 12px 0; font-size: 13px; font-weight: 600; color: var(--text); }
h4 { margin: 10px 0 6px 0; font-size: 11px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }
h5 { font-size: 11px; font-weight: 600; color: var(--text-muted); margin-bottom: 6px; }

.rules-grid {
  display: flex; gap: 12px; flex-wrap: wrap; align-items: center;
  margin-bottom: 8px; padding: 10px 12px;
  border: 1px solid var(--border); border-radius: 6px; background: var(--ink);
}
.rule-check {
  display: inline-flex; align-items: center; gap: 5px;
  font-size: 12px; color: var(--text-secondary); cursor: pointer; user-select: none;
}
.rule-check input[type="checkbox"] { width: 14px; height: 14px; accent-color: var(--brass); cursor: pointer; }
.rule-field { display: flex; align-items: center; gap: 4px; }
.rule-field label { font-size: 11px; color: var(--text-muted); }
.rule-field input[type="number"],
.rule-field input[type="text"] {
  width: 60px; padding: 4px 7px;
  border: 1px solid var(--border); border-radius: 4px;
  font-size: 12px; background: var(--paper); color: var(--text); outline: none;
}
.rule-field input:focus { border-color: var(--brass); }
.rule-field input:disabled { opacity: .35; }
.mt-3 { margin-top: 8px; }
.mt-4 { margin-top: 14px; }
</style>
